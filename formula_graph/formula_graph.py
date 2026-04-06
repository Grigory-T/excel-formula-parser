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
Nodes  – one per AST node, keyed by UUIDv4.
         Every node carries:
           id         str    unique UUIDv4 node id
           type       str    FunctionCall | BinaryOp | UnaryOp |
                             Reference | Number | Text | Bool
           label      str    human-readable display label
         Type-specific fields (only present when relevant):
           name       str    function name          (FunctionCall)
           op         str    operator symbol        (BinaryOp, UnaryOp)
           ref        str    cell / range string    (Reference)
           reference_class  str   parsed Excel reference kind (Reference)
           reference_scope  str   current_sheet | other_sheet |
                                  external_workbook           (Reference)
           value      any    literal value          (Number, Text, Bool)

Edges  – directed from parent → child.
         Every edge carries:
           arg_index  int    0-based position among the parent's children
"""

import re
from uuid import uuid4

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

# ── Reference parser ───────────────────────────────────────────────────────────

_CELL_PART_RE = re.compile(r"^(?P<abs_col>\$?)(?P<col>[A-Za-z]{1,3})(?P<abs_row>\$?)(?P<row>\d+)$")
_COLUMN_PART_RE = re.compile(r"^(?P<abs>\$?)(?P<col>[A-Za-z]{1,3})$")
_ROW_PART_RE = re.compile(r"^(?P<abs>\$?)(?P<row>\d+)$")


def _col_to_index(col: str) -> int:
    value = 0
    for char in col.upper():
        value = value * 26 + (ord(char) - ord("A") + 1)
    return value


def _split_ref_prefix(ref: str) -> tuple[str | None, str]:
    in_quotes = False
    i = 0
    while i < len(ref):
        char = ref[i]
        if char == "'":
            if in_quotes and i + 1 < len(ref) and ref[i + 1] == "'":
                i += 2
                continue
            in_quotes = not in_quotes
        elif char == "!" and not in_quotes:
            return ref[:i], ref[i + 1:]
        i += 1
    return None, ref


def _unquote_sheet_prefix(prefix: str) -> tuple[str, bool]:
    if len(prefix) >= 2 and prefix[0] == "'" and prefix[-1] == "'":
        return prefix[1:-1].replace("''", "'"), True
    return prefix, False


def _parse_sheet_prefix(prefix: str | None) -> dict[str, Any]:
    if prefix is None:
        return {
            "reference_scope": "current_sheet",
            "sheet_name": None,
            "sheet_quoted": False,
            "workbook_name": None,
            "workbook_path": None,
            "has_sheet": False,
            "has_workbook": False,
            "is_external": False,
        }

    raw_prefix, quoted = _unquote_sheet_prefix(prefix)
    workbook_name = None
    workbook_path = None
    sheet_name = raw_prefix

    if "[" in raw_prefix and "]" in raw_prefix:
        open_idx = raw_prefix.find("[")
        close_idx = raw_prefix.find("]", open_idx + 1)
        if close_idx != -1:
            workbook_path = raw_prefix[:open_idx] or None
            workbook_name = raw_prefix[open_idx + 1:close_idx] or None
            sheet_name = raw_prefix[close_idx + 1:] or None

    return {
        "reference_scope": "external_workbook" if workbook_name else "other_sheet",
        "sheet_name": sheet_name,
        "sheet_quoted": quoted,
        "workbook_name": workbook_name,
        "workbook_path": workbook_path,
        "has_sheet": True,
        "has_workbook": workbook_name is not None,
        "is_external": workbook_name is not None,
    }


def _parse_ref_part(part: str) -> dict[str, Any] | None:
    match = _CELL_PART_RE.match(part)
    if match:
        col = match.group("col").upper()
        row = int(match.group("row"))
        return {
            "part_kind": "cell",
            "text": part,
            "column": col,
            "column_index": _col_to_index(col),
            "row": row,
            "abs_column": bool(match.group("abs_col")),
            "abs_row": bool(match.group("abs_row")),
        }

    match = _COLUMN_PART_RE.match(part)
    if match:
        col = match.group("col").upper()
        return {
            "part_kind": "column",
            "text": part,
            "column": col,
            "column_index": _col_to_index(col),
            "abs_column": bool(match.group("abs")),
        }

    match = _ROW_PART_RE.match(part)
    if match:
        return {
            "part_kind": "row",
            "text": part,
            "row": int(match.group("row")),
            "abs_row": bool(match.group("abs")),
        }

    return None


def parse_reference(ref: str) -> dict[str, Any]:
    """
    Parse an Excel reference into structured metadata.

    Supported forms include:
    - single cells: A1, $B$2
    - cell ranges: A1:C3, Sheet1!$A$1:$B$2
    - whole columns: C:C, $A:$E
    - whole rows: 1:10
    - quoted/unquoted sheet names
    - external workbook references such as '[Book.xlsx]Sheet1'!A1 or
      'C:\\path\\[Book.xlsx]Sheet1'!A1

    Named ranges are returned as class "named_range".
    """
    prefix, area = _split_ref_prefix(ref)
    sheet_meta = _parse_sheet_prefix(prefix)

    if ":" in area:
        start_text, end_text = area.split(":", 1)
        start = _parse_ref_part(start_text)
        end = _parse_ref_part(end_text)
        if start and end:
            if start["part_kind"] == end["part_kind"] == "cell":
                ref_class = "cell_range"
            elif start["part_kind"] == end["part_kind"] == "column":
                ref_class = "column_range"
            elif start["part_kind"] == end["part_kind"] == "row":
                ref_class = "row_range"
            else:
                ref_class = "mixed_range"
            return {
                "reference_class": ref_class,
                "normalized_ref": ref,
                "reference_parts": {"start": start, "end": end},
                **sheet_meta,
            }
    else:
        part = _parse_ref_part(area)
        if part and part["part_kind"] == "cell":
            return {
                "reference_class": "cell",
                "normalized_ref": ref,
                "reference_parts": {"value": part},
                **sheet_meta,
            }

    return {
        "reference_class": "named_range",
        "normalized_ref": ref,
        "reference_parts": {"value": {"part_kind": "name", "text": area}},
        "reference_scope": "current_sheet" if prefix is None else sheet_meta["reference_scope"],
        **sheet_meta,
    }

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

def _add_node(G: nx.DiGraph, node, parent_id=None, arg_index=None) -> str:
    nid = str(uuid4())

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
                     **parse_reference(node.ref),
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
        _add_node(G, child, parent_id=nid, arg_index=i)

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
        Directed tree with exactly one root node.
        Node attributes: id, type, label, and type-specific fields
        (name / op / ref / value).
        Edge attributes: arg_index (0-based child position).
        Graph attribute: root_id (UUIDv4 of the root node).
    """
    tokens = _tokenize(formula)
    ast = _Parser(tokens).parse()
    G = nx.DiGraph()
    G.graph["root_id"] = _add_node(G, ast)
    return G
