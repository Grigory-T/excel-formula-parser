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
class _ArrayConstant:
    rows: list[list[Any]]

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
_STRUCTURED_REF = r"(?:[^\W\d][\w.]*)?(?:\[(?:[^\[\]]*|\[[^\[\]]*\])*\])+"
_AREA_REF = (
    rf"(?:{_CELL_REF}(?::{_CELL_REF})?"
    rf"|{_COLUMN_REF}:{_COLUMN_REF}"
    rf"|{_ROW_REF}:{_ROW_REF})"
)
_A1_REF_RE = rf"(?:{_SHEET_PREFIX})?{_AREA_REF}#?"
_TABLE_REF_RE = rf"(?:{_SHEET_PREFIX})?{_STRUCTURED_REF}#?"
_REF_RE = rf"(?:{_A1_REF_RE}|{_TABLE_REF_RE})"

_PATTERNS = [
    ("NUMBER", r"\d+(\.\d*)?([eE][+-]?\d+)?"),
    ("STRING", r'"(?:[^"]|"")*"'),
    ("BOOL",   r"\b(?:TRUE|FALSE)\b"),
    ("REF",    _REF_RE),                                     # A1, Sheet1!A1, 'Book'!A:E
    ("NAME",   r"[^\W\d][\w.]*"),                            # function / named range
    ("OP",     r"<=|>=|<>|[@+\-*/^=<>&%]"),
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


def _split_formula_wrapper(formula: str) -> tuple[str, dict[str, Any]]:
    src = formula.strip()
    meta = {
        "formula_type": "normal_formula",
        "formula_wrapper": None,
    }
    if src.startswith("{=") and src.endswith("}"):
        meta["formula_type"] = "array_formula"
        meta["formula_wrapper"] = "array"
        return src[2:-1], meta
    if src.startswith("="):
        return src[1:], meta
    return src, meta

# ── Reference parser ───────────────────────────────────────────────────────────

_CELL_PART_RE = re.compile(r"^(?P<abs_col>\$?)(?P<col>[A-Za-z]{1,3})(?P<abs_row>\$?)(?P<row>\d+)$")
_COLUMN_PART_RE = re.compile(r"^(?P<abs>\$?)(?P<col>[A-Za-z]{1,3})$")
_ROW_PART_RE = re.compile(r"^(?P<abs>\$?)(?P<row>\d+)$")
_TABLE_REF_RE_FULL = re.compile(r"^(?P<table>[^\W\d][\w.]*)?(?P<selectors>(?:\[(?:[^\[\]]*|\[[^\[\]]*\])*\])+)$")
_ERROR_REF_VALUES = {
    "#REF!",
    "#VALUE!",
    "#NAME?",
    "#N/A",
    "#NUM!",
    "#NULL!",
    "#DIV/0!",
    "#SPILL!",
    "#CALC!",
    "#FIELD!",
    "#BLOCKED!",
    "#UNKNOWN!",
    "#CONNECT!",
    "#GETTING_DATA",
}


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


def _split_top_level(ref: str, separators: set[str]) -> list[str]:
    parts: list[str] = []
    start = 0
    quote = False
    bracket_depth = 0
    paren_depth = 0
    brace_depth = 0
    i = 0
    while i < len(ref):
        char = ref[i]
        if char == "'":
            if quote and i + 1 < len(ref) and ref[i + 1] == "'":
                i += 2
                continue
            quote = not quote
        elif not quote:
            if char == "[":
                bracket_depth += 1
            elif char == "]" and bracket_depth > 0:
                bracket_depth -= 1
            elif char == "(":
                paren_depth += 1
            elif char == ")" and paren_depth > 0:
                paren_depth -= 1
            elif char == "{":
                brace_depth += 1
            elif char == "}" and brace_depth > 0:
                brace_depth -= 1
            elif (
                char in separators
                and bracket_depth == 0
                and paren_depth == 0
                and brace_depth == 0
            ):
                parts.append(ref[start:i].strip())
                start = i + 1
        i += 1
    parts.append(ref[start:].strip())
    return [part for part in parts if part]


def _split_top_level_intersection(ref: str) -> list[str]:
    parts: list[str] = []
    buf: list[str] = []
    quote = False
    bracket_depth = 0
    paren_depth = 0
    brace_depth = 0
    i = 0
    while i < len(ref):
        char = ref[i]
        if char == "'":
            if quote and i + 1 < len(ref) and ref[i + 1] == "'":
                buf.extend([char, ref[i + 1]])
                i += 2
                continue
            quote = not quote
            buf.append(char)
        elif not quote:
            if char == "[":
                bracket_depth += 1
                buf.append(char)
            elif char == "]":
                if bracket_depth > 0:
                    bracket_depth -= 1
                buf.append(char)
            elif char == "(":
                paren_depth += 1
                buf.append(char)
            elif char == ")":
                if paren_depth > 0:
                    paren_depth -= 1
                buf.append(char)
            elif char == "{":
                brace_depth += 1
                buf.append(char)
            elif char == "}":
                if brace_depth > 0:
                    brace_depth -= 1
                buf.append(char)
            elif (
                char == " "
                and bracket_depth == 0
                and paren_depth == 0
                and brace_depth == 0
            ):
                if buf and buf[-1] != " ":
                    parts.append("".join(buf).strip())
                    buf = []
                while i + 1 < len(ref) and ref[i + 1] == " ":
                    i += 1
            else:
                buf.append(char)
        else:
            buf.append(char)
        i += 1

    final = "".join(buf).strip()
    if final:
        parts.append(final)
    return parts if len(parts) > 1 else [ref.strip()]


def _unquote_sheet_prefix(prefix: str) -> tuple[str, bool]:
    if len(prefix) >= 2 and prefix[0] == "'" and prefix[-1] == "'":
        return prefix[1:-1].replace("''", "'"), True
    return prefix, False


def _parse_sheet_prefix(prefix: str | None) -> dict[str, Any]:
    if prefix is None:
        return {
            "reference_scope": "current_sheet",
            "sheet_name": None,
            "sheet_range_start": None,
            "sheet_range_end": None,
            "sheet_quoted": False,
            "workbook_name": None,
            "workbook_path": None,
            "has_sheet": False,
            "has_workbook": False,
            "is_external": False,
            "is_3d_reference": False,
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

    sheet_range_start = None
    sheet_range_end = None
    is_3d_reference = False
    if sheet_name and ":" in sheet_name:
        sheet_range_start, sheet_range_end = sheet_name.split(":", 1)
        is_3d_reference = True

    return {
        "reference_scope": "external_workbook" if workbook_name else "other_sheet",
        "sheet_name": sheet_name,
        "sheet_range_start": sheet_range_start,
        "sheet_range_end": sheet_range_end,
        "sheet_quoted": quoted,
        "workbook_name": workbook_name,
        "workbook_path": workbook_path,
        "has_sheet": True,
        "has_workbook": workbook_name is not None,
        "is_external": workbook_name is not None,
        "is_3d_reference": is_3d_reference,
    }


def _parse_table_reference(area: str) -> dict[str, Any] | None:
    match = _TABLE_REF_RE_FULL.match(area)
    if not match:
        return None

    selectors = match.group("selectors")
    raw_items = [item.strip() for item in re.findall(r"\[([^\[\]]+)\]", selectors) if item.strip()]

    table_selectors = [item for item in raw_items if item.startswith("#")]
    column_refs = [item for item in raw_items if not item.startswith("#")]
    return {
        "table_name": match.group("table"),
        "table_selectors": table_selectors,
        "table_columns": column_refs,
        "is_table_reference": True,
    }


def _empty_scope_meta() -> dict[str, Any]:
    return {
        "reference_scope": "current_sheet",
        "sheet_name": None,
        "sheet_range_start": None,
        "sheet_range_end": None,
        "sheet_quoted": False,
        "workbook_name": None,
        "workbook_path": None,
        "has_sheet": False,
        "has_workbook": False,
        "is_external": False,
        "is_3d_reference": False,
    }


def _parse_external_name_reference(ref: str) -> dict[str, Any] | None:
    if not (len(ref) >= 2 and ref[0] == "'" and ref[-1] == "'"):
        return None
    raw, quoted = _unquote_sheet_prefix(ref)
    if "[" not in raw or "]" not in raw:
        return None
    open_idx = raw.find("[")
    close_idx = raw.find("]", open_idx + 1)
    if close_idx == -1:
        return None
    workbook_path = raw[:open_idx] or None
    workbook_name = raw[open_idx + 1:close_idx] or None
    defined_name = raw[close_idx + 1:] or None
    if not workbook_name or not defined_name:
        return None
    return {
        "reference_class": "named_range",
        "normalized_ref": ref,
        "reference_operator": None,
        "reference_parts": {"value": {"part_kind": "name", "text": defined_name}},
        "is_spill_reference": False,
        "spill_anchor": None,
        "is_table_reference": False,
        "reference_scope": "external_workbook",
        "sheet_name": None,
        "sheet_range_start": None,
        "sheet_range_end": None,
        "sheet_quoted": quoted,
        "workbook_name": workbook_name,
        "workbook_path": workbook_path,
        "has_sheet": False,
        "has_workbook": True,
        "is_external": True,
        "is_3d_reference": False,
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
    ref = ref.strip()
    implicit_intersection = ref.startswith("@")
    if implicit_intersection:
        ref = ref[1:].strip()

    if ref.upper() in _ERROR_REF_VALUES:
        result = {
            "reference_class": "error_reference",
            "normalized_ref": ref,
            "reference_operator": None,
            "reference_parts": {"value": {"part_kind": "error", "text": ref}},
            "is_spill_reference": False,
            "spill_anchor": None,
            "is_table_reference": False,
            **_empty_scope_meta(),
        }
        result["implicit_intersection"] = implicit_intersection
        return result

    external_name = _parse_external_name_reference(ref)
    if external_name is not None:
        external_name["implicit_intersection"] = implicit_intersection
        return external_name

    bang_reference = ref.startswith("!")
    if bang_reference:
        ref = ref[1:]

    union_parts = _split_top_level(ref, {","})
    if len(union_parts) > 1:
        operands = [parse_reference(part) for part in union_parts]
        result = {
            "reference_class": "union",
            "normalized_ref": ref,
            "reference_operator": "union",
            "operands": operands,
            "reference_scope": "mixed",
            "is_external": any(item.get("is_external") for item in operands),
            "is_3d_reference": any(item.get("is_3d_reference") for item in operands),
            "is_table_reference": any(item.get("is_table_reference") for item in operands),
        }
        result["implicit_intersection"] = implicit_intersection
        return result

    intersection_parts = _split_top_level_intersection(ref)
    if len(intersection_parts) > 1:
        operands = [parse_reference(part) for part in intersection_parts]
        result = {
            "reference_class": "intersection",
            "normalized_ref": ref,
            "reference_operator": "intersection",
            "operands": operands,
            "reference_scope": "mixed",
            "is_external": any(item.get("is_external") for item in operands),
            "is_3d_reference": any(item.get("is_3d_reference") for item in operands),
            "is_table_reference": any(item.get("is_table_reference") for item in operands),
        }
        result["implicit_intersection"] = implicit_intersection
        return result

    prefix, area = _split_ref_prefix(ref)
    sheet_meta = _parse_sheet_prefix(prefix)
    if bang_reference:
        sheet_meta = {**sheet_meta, **_empty_scope_meta(), "reference_scope": "bang_reference"}
    spill_anchor = None
    is_spill_reference = area.endswith("#")
    if is_spill_reference:
        spill_anchor = area[:-1]
        area = spill_anchor

    table_meta = _parse_table_reference(area)
    if table_meta:
        result = {
            "reference_class": "table_reference",
            "normalized_ref": ref,
            "reference_operator": None,
            "reference_parts": {"value": {"part_kind": "table", "text": area}},
            "is_spill_reference": is_spill_reference,
            "spill_anchor": spill_anchor,
            **sheet_meta,
            **table_meta,
        }
        if is_spill_reference:
            result["reference_class"] = "spill_reference"
            result["spill_source"] = parse_reference(f"{prefix}!{spill_anchor}" if prefix else spill_anchor)
        result["implicit_intersection"] = implicit_intersection
        return result

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
                "reference_operator": "range",
                "reference_parts": {"start": start, "end": end},
                "is_spill_reference": is_spill_reference,
                "spill_anchor": spill_anchor,
                "is_table_reference": False,
                "implicit_intersection": implicit_intersection,
                **sheet_meta,
            }
    else:
        part = _parse_ref_part(area)
        if part and part["part_kind"] == "cell":
            return {
                "reference_class": "spill_reference" if is_spill_reference else "cell",
                "normalized_ref": ref,
                "reference_operator": None,
                "reference_parts": {"value": part},
                "is_spill_reference": is_spill_reference,
                "spill_anchor": spill_anchor,
                "is_table_reference": False,
                "implicit_intersection": implicit_intersection,
                **sheet_meta,
            }

    return {
        "reference_class": "named_range",
        "normalized_ref": ref,
        "reference_operator": None,
        "reference_parts": {"value": {"part_kind": "name", "text": area}},
        "reference_scope": "current_sheet" if prefix is None else sheet_meta["reference_scope"],
        "is_spill_reference": is_spill_reference,
        "spill_anchor": spill_anchor,
        "is_table_reference": False,
        "implicit_intersection": implicit_intersection,
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
        if self.at_op("+", "-", "@"):
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

        if kind == "LBRACE":
            return self._array_constant()

        raise SyntaxError(f"Unexpected token: {tok!r}")

    def _array_constant(self):
        self.consume("LBRACE")
        rows: list[list[Any]] = [[]]
        if self.peek() and self.peek()[0] != "RBRACE":
            rows[-1].append(self._comparison())
            while self.peek() and self.peek()[0] != "RBRACE":
                tok = self.consume("COMMA")
                if tok[1] == ";":
                    rows.append([self._comparison()])
                else:
                    rows[-1].append(self._comparison())
        self.consume("RBRACE")
        return _ArrayConstant(rows)

# ── Graph builder ──────────────────────────────────────────────────────────────

def _children(node) -> list:
    if isinstance(node, _FunctionCall): return node.args
    if isinstance(node, _BinaryOp):    return [node.left, node.right]
    if isinstance(node, _UnaryOp):     return [node.expr]
    if isinstance(node, _ArrayConstant):
        return [item for row in node.rows for item in row]
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
    elif isinstance(node, _ArrayConstant):
        attrs.update(
            type="ArrayConstant",
            rows=len(node.rows),
            cols=max((len(row) for row in node.rows), default=0),
            label=f"ArrayConstant({len(node.rows)}x{max((len(row) for row in node.rows), default=0)})",
        )
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
    src, meta = _split_formula_wrapper(formula)
    tokens = _tokenize(src)
    ast = _Parser(tokens).parse()
    G = nx.DiGraph()
    G.graph["root_id"] = _add_node(G, ast)
    G.graph.update(meta)
    return G
