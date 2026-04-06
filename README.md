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
