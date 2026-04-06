# excel-formula-parser

Small Python library for parsing Excel formulas into a NetworkX tree.

## API

```python
from formula_graph import parse_formula, parse_reference
```

- `parse_formula(formula: str) -> nx.DiGraph`
- `parse_reference(ref: str) -> dict`

There is no CLI entry point in this repo. The package is intended to be imported.

## Install

```bash
python3 -m venv venv
venv/bin/pip install networkx
```

## Formula Graph

`parse_formula()` returns a directed tree with one root:

- graph keys are UUIDv4 strings
- `G.graph["root_id"]` stores the root node id
- node fields are preserved in the NetworkX node attribute dict

Reference nodes keep the original `ref` string and also include parsed metadata such as:

- `reference_class`
- `reference_scope`
- `sheet_name`
- `workbook_name`
- `workbook_path`
- `reference_parts`

## Examples

### `parse_reference()`

```python
from formula_graph import parse_reference

ref_info = parse_reference("'Sheet 1'!$A:$E")
print(ref_info)
```

```python
{
    "reference_class": "column_range",
    "normalized_ref": "'Sheet 1'!$A:$E",
    "reference_parts": {
        "start": {
            "part_kind": "column",
            "text": "$A",
            "column": "A",
            "column_index": 1,
            "abs_column": True,
        },
        "end": {
            "part_kind": "column",
            "text": "$E",
            "column": "E",
            "column_index": 5,
            "abs_column": True,
        },
    },
    "reference_scope": "other_sheet",
    "sheet_name": "Sheet 1",
    "sheet_quoted": True,
    "workbook_name": None,
    "workbook_path": None,
    "has_sheet": True,
    "has_workbook": False,
    "is_external": False,
}
```

This shows all parsed reference fields:

- reference kind: `reference_class`
- normalized source text: `normalized_ref`
- parsed coordinates: `reference_parts`
- scope: `reference_scope`
- sheet/workbook metadata: `sheet_name`, `sheet_quoted`, `workbook_name`, `workbook_path`
- flags: `has_sheet`, `has_workbook`, `is_external`

### `parse_formula()`

```python
from formula_graph import parse_formula

G = parse_formula("=SUM(A1,B1*2)")
```

Graph-level information:

```python
G.graph == {
    "root_id": "<uuid-v4-of-root>"
}
```

Node data stored in NetworkX:

```python
{
    "<uuid-1>": {
        "id": "<uuid-1>",
        "type": "FunctionCall",
        "name": "SUM",
        "label": 'FunctionCall("SUM")',
    },
    "<uuid-2>": {
        "id": "<uuid-2>",
        "type": "Reference",
        "ref": "A1",
        "reference_class": "cell",
        "normalized_ref": "A1",
        "reference_parts": {
            "value": {
                "part_kind": "cell",
                "text": "A1",
                "column": "A",
                "column_index": 1,
                "row": 1,
                "abs_column": False,
                "abs_row": False,
            }
        },
        "reference_scope": "current_sheet",
        "sheet_name": None,
        "sheet_quoted": False,
        "workbook_name": None,
        "workbook_path": None,
        "has_sheet": False,
        "has_workbook": False,
        "is_external": False,
        "label": 'Reference("A1")',
    },
    "<uuid-3>": {
        "id": "<uuid-3>",
        "type": "BinaryOp",
        "op": "*",
        "label": 'BinaryOp("*")',
    },
    "<uuid-4>": {
        "id": "<uuid-4>",
        "type": "Reference",
        "ref": "B1",
        "reference_class": "cell",
        "normalized_ref": "B1",
        "reference_parts": {
            "value": {
                "part_kind": "cell",
                "text": "B1",
                "column": "B",
                "column_index": 2,
                "row": 1,
                "abs_column": False,
                "abs_row": False,
            }
        },
        "reference_scope": "current_sheet",
        "sheet_name": None,
        "sheet_quoted": False,
        "workbook_name": None,
        "workbook_path": None,
        "has_sheet": False,
        "has_workbook": False,
        "is_external": False,
        "label": 'Reference("B1")',
    },
    "<uuid-5>": {
        "id": "<uuid-5>",
        "type": "Number",
        "value": 2,
        "label": "Number(2)",
    },
}
```

Edges:

```python
[
    ("<uuid-1>", "<uuid-2>", {"arg_index": 0}),
    ("<uuid-1>", "<uuid-3>", {"arg_index": 1}),
    ("<uuid-3>", "<uuid-4>", {"arg_index": 0}),
    ("<uuid-3>", "<uuid-5>", {"arg_index": 1}),
]
```

Tree view of the same formula:

```text
FunctionCall("SUM")
├── Reference("A1")
└── BinaryOp("*")
    ├── Reference("B1")
    └── Number(2)
```

This example shows all information produced during formula parsing:

- graph root: `G.graph["root_id"]`
- one NetworkX node per AST node
- all parsed fields stored directly on each node
- argument order preserved on edges via `arg_index`
- a single-root tree for each formula

## Reference Parser

`parse_reference()` handles:

- single cells like `A1`, `$B$2`
- cell ranges like `A1:C3`
- full columns like `A:E`, `$A:$E`
- full rows like `1:10`
- same-sheet and other-sheet refs
- quoted sheet names
- Cyrillic and Latin sheet names
- external workbook refs such as `'[Book.xlsx]Sheet1'!A1`
- full-path workbook refs such as `'C:\path\[Book.xlsx]Sheet1'!A1`

## Layout

```text
formula_graph/
  __init__.py
  formula_graph.py
```
