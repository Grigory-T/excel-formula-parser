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

**Nodes** — one per AST node, root is always node `0`.

| Attribute | Type | Always present | Description |
|-----------|------|----------------|-------------|
| `id`      | int  | yes | unique node id |
| `type`    | str  | yes | `FunctionCall`, `BinaryOp`, `UnaryOp`, `Reference`, `Number`, `Text`, `Bool` |
| `label`   | str  | yes | human-readable display string |
| `name`    | str  | FunctionCall | function name (e.g. `"SUM"`) |
| `op`      | str  | BinaryOp, UnaryOp | operator symbol (e.g. `"*"`, `"-"`) |
| `ref`     | str  | Reference | cell or range string (e.g. `"$A2"`, `"A1:B3"`) |
| `value`   | any  | Number, Text, Bool | literal value |

**Edges** — directed from parent → child.

| Attribute   | Type | Description |
|-------------|------|-------------|
| `arg_index` | int  | 0-based position among siblings |

## Example

```
=SUM(A1, B1*2)

nodes:
  0  FunctionCall  name="SUM"
  1  Reference     ref="A1"
  2  BinaryOp      op="*"
  3  Reference     ref="B1"
  4  Number        value=2

edges:
  0 → 1  arg_index=0
  0 → 2  arg_index=1
  2 → 3  arg_index=0
  2 → 4  arg_index=1
```

## Supported syntax

- Functions: `SUM(...)`, `IF(...)`, nested calls
- Binary operators: `+ - * / ^ & = <> < > <= >=`
- Unary operators: `+ -`
- Cell references: `A1`, `$A2`, `$A$2`, `A1:B3`, `$M$2:$Q$200`
- Named ranges
- String literals: `"text"`
- Numbers: integers, decimals, scientific notation
- Booleans: `TRUE`, `FALSE`
- Whitespace and newlines inside formulas
