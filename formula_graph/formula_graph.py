"""
formula_graph.py
----------------
Parse an Excel formula string into a NetworkX directed graph (top → bottom).

Usage
-----
    from formula_graph import parse_formula
    G = parse_formula("=SUM(A1, B1*2)")

Graph structure
---------------
Nodes  – one per AST node, keyed by integer id (root = 0).
         Every node carries:
           id         int    unique node id
           type       str    FunctionCall | BinaryOp | UnaryOp |
                             Reference | Number | Text | Bool
           label      str    human-readable display label
         Type-specific fields (only present when relevant):
           name       str    function name          (FunctionCall)
           op         str    operator symbol        (BinaryOp, UnaryOp)
           ref        str    cell / range string    (Reference)
           value      any    literal value          (Number, Text, Bool)

Edges  – directed from parent → child.
         Every edge carries:
           arg_index  int    0-based position among the parent's children
"""

import re
import networkx as nx
from dataclasses import dataclass, field
from typing import Any

# ── AST nodes ──────────────────────────────────────────────────────────────────

@dataclass
class _Number:
    value: float | int

@dataclass
class _Text:
    value: str

@dataclass
class _Bool:
    value: bool

@dataclass
class _Reference:
    ref: str

@dataclass
class _FunctionCall:
    name: str
    args: list = field(default_factory=list)

@dataclass
class _BinaryOp:
    op: str
    left: Any
    right: Any

@dataclass
class _UnaryOp:
    op: str
    expr: Any

# ── Tokenizer ──────────────────────────────────────────────────────────────────

_QUOTED_SHEET = r"'(?:[^']|'')*'"
_UNQUOTED_SHEET = r"[^\s,;(){}+\-*/^=<>&%!]+"
_SHEET_PREFIX = rf"(?:(?:{_QUOTED_SHEET}|{_UNQUOTED_SHEET})!)"
_CELL_REF = r"\$?[A-Za-z]{1,3}\$?\d+"
_COLUMN_REF = r"\$?[A-Za-z]{1,3}"
_ROW_REF = r"\$?\d+"
_AREA_REF = (
    rf"(?:{_CELL_REF}(?::{_CELL_REF})?"
    rf"|{_COLUMN_REF}:{_COLUMN_REF}"
    rf"|{_ROW_REF}:{_ROW_REF})"
)
_REF_RE = rf"(?:{_SHEET_PREFIX})?{_AREA_REF}"

_PATTERNS = [
    ("NUMBER", r"\d+(\.\d*)?([eE][+-]?\d+)?"),
    ("STRING", r'"(?:[^"]|"")*"'),
    ("BOOL",   r"\b(?:TRUE|FALSE)\b"),
    ("REF",    _REF_RE),                                     # A1, Sheet1!A1, 'Book'!A:E
    ("NAME",   r"[^\W\d][\w.]*"),                            # function / named range
    ("OP",     r"<=|>=|<>|[+\-*/^=<>&%]"),
    ("LPAREN", r"\("),
    ("RPAREN", r"\)"),
    ("COMMA",  r"[,;]"),
    ("LBRACE", r"\{"),
    ("RBRACE", r"\}"),
    ("WS",     r"[ \t\r\n]+"),
]

_TOKEN_RE = re.compile("|".join(f"(?P<{n}>{p})" for n, p in _PATTERNS))


def _tokenize(formula: str) -> list[tuple[str, str]]:
    src = formula[1:] if formula.startswith("=") else formula
    tokens = []
    pos = 0
    while pos < len(src):
        m = _TOKEN_RE.match(src, pos)
        if not m:
            raise SyntaxError(f"Unexpected character at pos {pos}: {src[pos]!r}")
        if m.lastgroup != "WS":
            tokens.append((m.lastgroup, m.group()))
        pos = m.end()
    return tokens

# ── Parser (recursive descent) ─────────────────────────────────────────────────

class _Parser:
    def __init__(self, tokens):
        self.tokens = tokens
        self.pos = 0

    def peek(self):
        return self.tokens[self.pos] if self.pos < len(self.tokens) else None

    def consume(self, kind=None):
        tok = self.tokens[self.pos]
        if kind and tok[0] != kind:
            raise SyntaxError(f"Expected {kind}, got {tok!r}")
        self.pos += 1
        return tok

    def at_op(self, *ops):
        t = self.peek()
        return t is not None and t[0] == "OP" and t[1] in ops

    def parse(self):
        node = self._comparison()
        if self.peek():
            raise SyntaxError(f"Unexpected token: {self.peek()!r}")
        return node

    def _comparison(self):
        left = self._concat()
        while self.at_op("<", ">", "<=", ">=", "=", "<>"):
            op = self.consume()[1]
            left = _BinaryOp(op, left, self._concat())
        return left

    def _concat(self):
        left = self._addition()
        while self.at_op("&"):
            op = self.consume()[1]
            left = _BinaryOp(op, left, self._addition())
        return left

    def _addition(self):
        left = self._multiplication()
        while self.at_op("+", "-"):
            op = self.consume()[1]
            left = _BinaryOp(op, left, self._multiplication())
        return left

    def _multiplication(self):
        left = self._exponent()
        while self.at_op("*", "/"):
            op = self.consume()[1]
            left = _BinaryOp(op, left, self._exponent())
        return left

    def _exponent(self):
        left = self._unary()
        if self.at_op("^"):
            op = self.consume()[1]
            return _BinaryOp(op, left, self._exponent())   # right-associative
        return left

    def _unary(self):
        if self.at_op("+", "-"):
            op = self.consume()[1]
            return _UnaryOp(op, self._unary())
        return self._primary()

    def _primary(self):
        tok = self.peek()
        if tok is None:
            raise SyntaxError("Unexpected end of formula")
        kind, val = tok

        if kind == "NUMBER":
            self.consume()
            f = float(val)
            return _Number(int(f) if f == int(f) else f)

        if kind == "STRING":
            self.consume()
            return _Text(val[1:-1].replace('""', '"'))

        if kind == "BOOL":
            self.consume()
            return _Bool(val.upper() == "TRUE")

        if kind == "REF":
            self.consume()
            return _Reference(val)

        if kind == "NAME":
            name = self.consume()[1]
            if self.peek() and self.peek()[0] == "LPAREN":
                self.consume("LPAREN")
                args = []
                if self.peek() and self.peek()[0] != "RPAREN":
                    args.append(self._comparison())
                    while self.peek() and self.peek()[0] == "COMMA":
                        self.consume()
                        args.append(self._comparison())
                self.consume("RPAREN")
                return _FunctionCall(name.upper(), args)
            return _Reference(name)   # named range

        if kind == "LPAREN":
            self.consume()
            node = self._comparison()
            self.consume("RPAREN")
            return node

        raise SyntaxError(f"Unexpected token: {tok!r}")

# ── Graph builder ──────────────────────────────────────────────────────────────

def _children(node) -> list:
    if isinstance(node, _FunctionCall): return node.args
    if isinstance(node, _BinaryOp):    return [node.left, node.right]
    if isinstance(node, _UnaryOp):     return [node.expr]
    return []

def _add_node(G: nx.DiGraph, node, counter: list, parent_id=None, arg_index=None) -> int:
    nid = counter[0]
    counter[0] += 1

    attrs = {"id": nid}

    if isinstance(node, _FunctionCall):
        attrs.update(type="FunctionCall", name=node.name,
                     label=f'FunctionCall("{node.name}")')
    elif isinstance(node, _BinaryOp):
        attrs.update(type="BinaryOp", op=node.op,
                     label=f'BinaryOp("{node.op}")')
    elif isinstance(node, _UnaryOp):
        attrs.update(type="UnaryOp", op=node.op,
                     label=f'UnaryOp("{node.op}")')
    elif isinstance(node, _Reference):
        attrs.update(type="Reference", ref=node.ref,
                     label=f'Reference("{node.ref}")')
    elif isinstance(node, _Number):
        attrs.update(type="Number", value=node.value,
                     label=f"Number({node.value})")
    elif isinstance(node, _Text):
        attrs.update(type="Text", value=node.value,
                     label=f'Text("{node.value}")')
    elif isinstance(node, _Bool):
        attrs.update(type="Bool", value=node.value,
                     label=f"Bool({node.value})")

    G.add_node(nid, **attrs)

    if parent_id is not None:
        G.add_edge(parent_id, nid, arg_index=arg_index)

    for i, child in enumerate(_children(node)):
        _add_node(G, child, counter, parent_id=nid, arg_index=i)

    return nid

# ── Public API ─────────────────────────────────────────────────────────────────

def parse_formula(formula: str) -> nx.DiGraph:
    """
    Parse an Excel formula string into a directed graph (parent → child).

    Parameters
    ----------
    formula : str
        Excel formula, with or without leading '='.

    Returns
    -------
    nx.DiGraph
        Directed graph where node 0 is the root.
        Node attributes: id, type, label, and type-specific fields
        (name / op / ref / value).
        Edge attributes: arg_index (0-based child position).
    """
    tokens = _tokenize(formula)
    ast = _Parser(tokens).parse()
    G = nx.DiGraph()
    _add_node(G, ast, counter=[0])
    return G
