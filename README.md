# openpyxl (xtend fork)

This is an **experimental fork** of [openpyxl](https://foss.heptapod.net/openpyxl/openpyxl) with extensions to handle more of the full OOXML (.xlsx) format.

The upstream *openpyxl* library provides excellent support for basic spreadsheet operations. This fork extends it with preservation and support for modern Excel 365 features commonly found in real-world spreadsheets.

## Extensions

This fork adds the following capabilities:

| Feature | Description |
|---------|-------------|
| **Dynamic Arrays** | Preserve metadata.xml for spill formulas (UNIQUE, SORT, FILTER, etc.) |
| **Threaded Comments** | Read/preserve modern Excel 365 comments with replies and authors |
| **Sparklines** | Full read/write support for in-cell mini-charts |
| **Slicers & Timelines** | Preserve interactive filter controls on round-trip |
| **Modern Charts** | Preserve chartex charts (waterfall, funnel, treemap, sunburst, etc.) |
| **Rich Data Types** | Preserve Stock/Geography linked data types |
| **Excel 365 Functions** | 147 modern functions added (XLOOKUP, LET, LAMBDA, etc.) |
| **Comment Formatting** | Preserve rich text formatting in legacy comments |
| **Extension Preservation** | Stop losing extLst content on round-trip |
| **Unknown Element Preservation** | Preserve unrecognized XML elements |

## Installation

```bash
pip install git+https://github.com/rekrevs/openpyxl.git@xtend
```

## Example

```python
>>> from openpyxl import Workbook, load_workbook

>>> wb = Workbook()
>>> ws = wb.active
>>> ws['A1'] = 42
>>> ws.append([1, 2, 3])
>>> wb.save("sample.xlsx")

>>> wb = load_workbook("sample.xlsx")
>>> ws = wb.active
>>> ws['A1'].value
42
```

### Round-Trip Preservation

Files with modern Excel features now survive round-trip without data loss:

```python
>>> from openpyxl import load_workbook

# Load a file with dynamic arrays, sparklines, slicers, etc.
>>> wb = load_workbook("modern-features.xlsx")

# Make changes
>>> wb.active['Z1'] = "New data"

# Save - modern features are preserved!
>>> wb.save("modified.xlsx")
```

### Sparklines

```python
>>> from openpyxl import load_workbook

>>> wb = load_workbook("sparklines.xlsx")
>>> ws = wb.active

# Sparklines are preserved via worksheet extensions
>>> if ws.extensions:
...     print(f"Extensions preserved: {len(ws.extensions.ext)}")
```

### Threaded Comments

```python
>>> from openpyxl import load_workbook

>>> wb = load_workbook("with-comments.xlsx")

# Access comment authors
>>> if wb.persons:
...     for person in wb.persons:
...         print(f"{person.displayName} ({person.providerId})")

# Legacy comment text shows threaded comment fallback
>>> ws = wb.active
>>> if ws['A1'].comment:
...     print(ws['A1'].comment.text)
```

## Documentation

For core openpyxl functionality, see the [openpyxl documentation](https://openpyxl.readthedocs.io/en/stable/).

Extension features are documented in the `WOTAN/` directory:

- [WOTAN/docs/vision.md](WOTAN/docs/vision.md) - Project goals and target capabilities
- [WOTAN/docs/architecture.md](WOTAN/docs/architecture.md) - openpyxl structure and extension points
- [WOTAN/docs/openpyxl-api.md](WOTAN/docs/openpyxl-api.md) - Comprehensive API reference
- [WOTAN/docs/backlog.md](WOTAN/docs/backlog.md) - Feature status and roadmap
- [WOTAN/docs/research/](WOTAN/docs/research/) - OOXML specification research

## Status

This is an experimental fork focused on **WOTAN** (Working On The Actual Nuances) - bringing openpyxl to full modern XLSX compatibility.

### Completed

- Preserve extLst/extension content on round-trip
- Preserve unknown worksheet XML elements
- Dynamic array metadata.xml infrastructure
- 147 Excel 365/2019+ functions added to FORMULAE
- Sparkline read/write support
- Slicer and timeline preservation
- Modern chart (chartex) preservation
- Rich data type preservation
- Comment formatting preservation
- Pivot table builder

### In Progress

- Full threaded comments support (read/write/create)
- Dynamic array cell metadata (`cm` attribute)
- Spilled range operator (`#`) in formulas

### Backlog

- DrawingML shapes preservation
- Table creation API
- Full chartex chart creation

All original openpyxl tests pass. Extensions are additive and should not break existing functionality.

## Security

By default openpyxl does not guard against quadratic blowup or billion laughs XML attacks. To guard against these attacks install defusedxml.

## Mailing List

The user list can be found at http://groups.google.com/group/openpyxl-users
