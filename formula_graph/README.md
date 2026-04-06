# formula_graph

Parse an Excel formula into a NetworkX directed graph (AST, top → bottom).

## Setup

```bash
python3 -m venv venv
venv/bin/pip install networkx
```

## Usage

```python
from formula_graph import parse_formula

G = parse_formula("=SUM(A1, B1*2)")
```

## Graph structure

**Nodes** — one per AST node, keyed by UUIDv4.

| Attribute | Type | Always present | Description |
|-----------|------|----------------|-------------|
| `id`      | str  | yes | unique UUIDv4 node id |
| `type`    | str  | yes | `FunctionCall`, `BinaryOp`, `UnaryOp`, `Reference`, `Number`, `Text`, `Bool` |
| `label`   | str  | yes | human-readable display string |
| `name`    | str  | FunctionCall | function name (e.g. `"SUM"`) |
| `op`      | str  | BinaryOp, UnaryOp | operator symbol (e.g. `"*"`, `"-"`) |
| `ref`     | str  | Reference | cell or range string (e.g. `"$A2"`, `"A1:B3"`) |
| `reference_class` | str | Reference | `cell`, `cell_range`, `column_range`, `row_range`, or `named_range` |
| `reference_scope` | str | Reference | `current_sheet`, `other_sheet`, or `external_workbook` |
| `sheet_name` | str or null | Reference | parsed sheet name, if present |
| `workbook_name` | str or null | Reference | parsed workbook file name, if present |
| `workbook_path` | str or null | Reference | parsed external workbook path, if present |
| `reference_parts` | dict | Reference | parsed row/column/cell coordinates for the ref |
| `value`   | any  | Number, Text, Bool | literal value |

**Edges** — directed from parent → child.

| Attribute   | Type | Description |
|-------------|------|-------------|
| `arg_index` | int  | 0-based position among siblings |

**Graph** — each parsed formula is a tree, so there is always exactly one root.

| Attribute | Type | Description |
|-----------|------|-------------|
| `root_id` | str  | UUIDv4 of the root node |

## Example

```
=SUM(A1, B1*2)

nodes:
  550e8400-e29b-41d4-a716-446655440000  FunctionCall  name="SUM"
  1a4f6c08-2e8b-4103-b9c9-4e7a3afc8f91  Reference     ref="A1"
  28cf18e8-98b7-40ad-8410-5d279d40f247  BinaryOp      op="*"
  2c8f9ce1-6c5d-4897-99dc-55d65d0bf0fd  Reference     ref="B1"
  9da8fd2a-1ed3-4236-b3fc-c5ec9f0a6ea6  Number        value=2

edges:
  550e8400-e29b-41d4-a716-446655440000 → 1a4f6c08-2e8b-4103-b9c9-4e7a3afc8f91  arg_index=0
  550e8400-e29b-41d4-a716-446655440000 → 28cf18e8-98b7-40ad-8410-5d279d40f247  arg_index=1
  28cf18e8-98b7-40ad-8410-5d279d40f247 → 2c8f9ce1-6c5d-4897-99dc-55d65d0bf0fd  arg_index=0
  28cf18e8-98b7-40ad-8410-5d279d40f247 → 9da8fd2a-1ed3-4236-b3fc-c5ec9f0a6ea6  arg_index=1
```

## Supported syntax

- Functions: `SUM(...)`, `IF(...)`, nested calls
- Binary operators: `+ - * / ^ & = <> < > <= >=`
- Unary operators: `+ -`
- Cell references: `A1`, `$A2`, `$A$2`, `A1:B3`, `$M$2:$Q$200`
- Whole-column and whole-row references: `A:E`, `$A:$E`, `1:10`
- Sheet-qualified references, including quoted sheet names and Cyrillic names
- External workbook references, including full paths like `'C:\path\[Book.xlsx]Sheet1'!A1`
- Named ranges
- String literals: `"text"`
- Numbers: integers, decimals, scientific notation
- Booleans: `TRUE`, `FALSE`
- Whitespace and newlines inside formulas
