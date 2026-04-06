#!/usr/bin/env python3
"""Parse an Excel formula into an AST and print it as a tree."""

import re
import sys
from dataclasses import dataclass
from typing import Any, List

# ── AST nodes ──────────────────────────────────────────────────────────────────

@dataclass
class Number:
    value: float | int

@dataclass
class Text:
    value: str

@dataclass
class Bool:
    value: bool

@dataclass
class Reference:
    ref: str

@dataclass
class FunctionCall:
    name: str
    args: List[Any]

@dataclass
class BinaryOp:
    op: str
    left: Any
    right: Any

@dataclass
class UnaryOp:
    op: str
    expr: Any

# ── Tokenizer ──────────────────────────────────────────────────────────────────

_PATTERNS = [
    ("NUMBER", r"\d+(\.\d*)?([eE][+-]?\d+)?"),
    ("STRING", r'"(?:[^"]|"")*"'),
    ("BOOL",   r"\b(?:TRUE|FALSE)\b"),
    ("REF",    r"\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?"),  # A1, $B$2, A1:C3
    ("NAME",   r"[A-Za-z_][A-Za-z0-9_.]*"),               # function / named range
    ("OP",     r"<=|>=|<>|[+\-*/^=<>&%]"),
    ("LPAREN", r"\("),
    ("RPAREN", r"\)"),
    ("COMMA",  r"[,;]"),                                   # ; is regional comma
    ("LBRACE", r"\{"),
    ("RBRACE", r"\}"),
    ("WS",     r"[ \t\r\n]+"),
]

_TOKEN_RE = re.compile("|".join(f"(?P<{n}>{p})" for n, p in _PATTERNS))


def tokenize(formula: str) -> list[tuple[str, str]]:
    src = formula[1:] if formula.startswith("=") else formula
    tokens = []
    pos = 0
    while pos < len(src):
        m = _TOKEN_RE.match(src, pos)
        if not m:
            raise SyntaxError(f"Unexpected character at pos {pos}: {src[pos]!r}")
        kind = m.lastgroup
        if kind != "WS":
            tokens.append((kind, m.group()))
        pos = m.end()
    return tokens

# ── Parser (recursive descent) ─────────────────────────────────────────────────

class Parser:
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

    # Grammar: formula → comparison
    # comparison  → concat  (('<'|'>'|'<='|'>='|'='|'<>') concat)*
    # concat      → addition ('&' addition)*
    # addition    → mult     (('+' | '-') mult)*
    # mult        → exponent (('*' | '/') exponent)*
    # exponent    → unary    ('^' unary)*          right-assoc
    # unary       → ('+' | '-') unary | primary
    # primary     → number | string | bool | ref | func '(' args ')' | '(' expr ')'

    def parse(self):
        node = self.comparison()
        if self.peek():
            raise SyntaxError(f"Unexpected token: {self.peek()!r}")
        return node

    def comparison(self):
        left = self.concat()
        while self.at_op("<", ">", "<=", ">=", "=", "<>"):
            op = self.consume()[1]
            left = BinaryOp(op, left, self.concat())
        return left

    def concat(self):
        left = self.addition()
        while self.at_op("&"):
            op = self.consume()[1]
            left = BinaryOp(op, left, self.addition())
        return left

    def addition(self):
        left = self.multiplication()
        while self.at_op("+", "-"):
            op = self.consume()[1]
            left = BinaryOp(op, left, self.multiplication())
        return left

    def multiplication(self):
        left = self.exponent()
        while self.at_op("*", "/"):
            op = self.consume()[1]
            left = BinaryOp(op, left, self.exponent())
        return left

    def exponent(self):
        left = self.unary()
        if self.at_op("^"):
            op = self.consume()[1]
            right = self.exponent()          # right-associative
            return BinaryOp(op, left, right)
        return left

    def unary(self):
        if self.at_op("+", "-"):
            op = self.consume()[1]
            return UnaryOp(op, self.unary())
        return self.primary()

    def primary(self):
        tok = self.peek()
        if tok is None:
            raise SyntaxError("Unexpected end of formula")

        kind, val = tok

        if kind == "NUMBER":
            self.consume()
            f = float(val)
            return Number(int(f) if f == int(f) else f)

        if kind == "STRING":
            self.consume()
            return Text(val[1:-1].replace('""', '"'))

        if kind == "BOOL":
            self.consume()
            return Bool(val == "TRUE")

        if kind == "REF":
            self.consume()
            return Reference(val)

        if kind == "NAME":
            name = self.consume()[1]
            if self.peek() and self.peek()[0] == "LPAREN":
                self.consume("LPAREN")
                args = []
                if self.peek() and self.peek()[0] != "RPAREN":
                    args.append(self.comparison())
                    while self.peek() and self.peek()[0] == "COMMA":
                        self.consume()
                        args.append(self.comparison())
                self.consume("RPAREN")
                return FunctionCall(name.upper(), args)
            return Reference(name)           # named range

        if kind == "LPAREN":
            self.consume()
            node = self.comparison()
            self.consume("RPAREN")
            return node

        raise SyntaxError(f"Unexpected token: {tok!r}")

# ── Tree renderer ──────────────────────────────────────────────────────────────

def _label(node) -> str:
    if isinstance(node, FunctionCall): return f'FunctionCall("{node.name}")'
    if isinstance(node, BinaryOp):    return f'BinaryOp("{node.op}")'
    if isinstance(node, UnaryOp):     return f'UnaryOp("{node.op}")'
    if isinstance(node, Reference):   return f'Reference("{node.ref}")'
    if isinstance(node, Number):      return f'Number({node.value})'
    if isinstance(node, Text):        return f'Text("{node.value}")'
    if isinstance(node, Bool):        return f'Bool({node.value})'
    return repr(node)

def _children(node) -> list:
    if isinstance(node, FunctionCall): return node.args
    if isinstance(node, BinaryOp):    return [node.left, node.right]
    if isinstance(node, UnaryOp):     return [node.expr]
    return []

def _render(node, prefix: str, is_last: bool) -> list[str]:
    connector   = "└── " if is_last else "├── "
    child_pfx   = prefix + ("    " if is_last else "│   ")
    lines = [prefix + connector + _label(node)]
    kids = _children(node)
    for i, child in enumerate(kids):
        lines += _render(child, child_pfx, i == len(kids) - 1)
    return lines

def tree_str(node) -> str:
    lines = [_label(node)]
    kids = _children(node)
    for i, child in enumerate(kids):
        lines += _render(child, "", i == len(kids) - 1)
    return "\n".join(lines)

# ── Entry point ────────────────────────────────────────────────────────────────

def parse_formula(formula: str) -> str:
    tokens = tokenize(formula)
    ast = Parser(tokens).parse()
    return tree_str(ast)

if __name__ == "__main__":
    formula = sys.argv[1] if len(sys.argv) > 1 else input("Formula: ").strip()
    print(parse_formula(formula))
