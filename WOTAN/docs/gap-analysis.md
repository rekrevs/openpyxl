# WOTAN Gap Analysis

This document catalogs the gaps between openpyxl v3.1.5 and modern Excel 365 XLSX format.

## Executive Summary

| Aspect | Coverage | Notes |
|--------|----------|-------|
| **Core XLSX (2010)** | 95% | Excellent |
| **Excel 2016** | 85% | Good, some chart gaps |
| **Excel 2019** | 70% | Missing newer features |
| **Excel 365** | 40% | Major gaps |
| **Excel 365 (2024)** | 30% | Critical gaps in dynamic arrays |

---

## Specification Context

- **ECMA-376**: 5th Edition (December 2016) - last official update
- **ISO/IEC 29500**: Aligned with ECMA-376
- **MS-XLSX**: Microsoft's extension specification (continuously updated)
  - Current version: 29.0 (September 2025)
  - Documents proprietary extensions beyond ECMA-376

---

## Feature Support Matrix

### Fully Supported (No Action Needed)

| Feature | Location | Notes |
|---------|----------|-------|
| Cell values | `openpyxl/cell/` | All Python types |
| Fonts | `openpyxl/styles/fonts.py` | 13+ attributes |
| Fills | `openpyxl/styles/fills.py` | 19 patterns + gradients |
| Borders | `openpyxl/styles/borders.py` | 13 styles |
| Alignment | `openpyxl/styles/alignment.py` | Full support |
| Number formats | `openpyxl/styles/` | 166+ built-in + custom |
| Standard formulas | `openpyxl/formula/` | Full tokenizer |
| Array formulas | `openpyxl/worksheet/formula.py` | `ArrayFormula` class |
| Data table formulas | `openpyxl/worksheet/formula.py` | `DataTableFormula` class |
| Conditional formatting | `openpyxl/formatting/` | 14+ rule types |
| Data validation | `openpyxl/worksheet/datavalidation.py` | All 6 types |
| AutoFilter | `openpyxl/worksheet/filters.py` | Full support |
| Classic charts | `openpyxl/chart/` | 25+ types |
| Images | `openpyxl/drawing/image.py` | PNG, JPEG, GIF |
| Drawings/shapes | `openpyxl/drawing/` | Full support |
| Named ranges | `openpyxl/workbook/defined_name.py` | Global/local |
| Hyperlinks | `openpyxl/worksheet/hyperlink.py` | All types |
| Merged cells | `openpyxl/worksheet/merge.py` | Full support |
| Sheet/workbook protection | `openpyxl/worksheet/protection.py` | SHA-512 |
| Print settings | `openpyxl/worksheet/print_settings.py` | Full support |
| Rich text | `openpyxl/cell/rich_text.py` | Multi-format text |
| External links | `openpyxl/workbook/external_link/` | Full support |
| Document properties | `openpyxl/packaging/core.py` | Core + custom |

### Partially Supported (Needs Enhancement)

| Feature | Current State | Gap |
|---------|--------------|-----|
| **Pivot Tables** | Read/modify only | Cannot create new pivot tables |
| **Tables** | Read/modify only | Cannot create programmatically |
| **Comments** | Text only | Formatting lost on read; dimensions lost |
| **VBA/Macros** | Preserve with `keep_vba=True` | Cannot create/execute |
| **Sparklines** | GUID recognized | No read/write/create implementation |
| **Slicers** | GUID recognized | No implementation |
| **Timelines** | GUID recognized | No implementation |

### Not Supported (Critical Gaps)

#### 1. Dynamic Arrays

**Impact**: HIGH - Fundamental to modern Excel

| Component | Status | Required Changes |
|-----------|--------|------------------|
| Cell metadata (`cm="1"`) | Missing | Add to cell writer |
| `metadata.xml` | Missing | New file, new relationship |
| `futureMetadata XLDAPR` | Missing | New XML structure |
| Spilled range operator (`#`) | Not parsed | Update tokenizer |
| `calcChain` for bi-directional spill | Missing | New handler |

**Files to create/modify**:
- `openpyxl/packaging/metadata.py` (NEW)
- `openpyxl/cell/_writer.py` - add `cm` attribute
- `openpyxl/workbook/_writer.py` - add metadata relationship
- `openpyxl/xml/constants.py` - add `METADATA_NS`
- `openpyxl/formula/tokenizer.py` - handle `#` operator

**Specification**: [MS-XLSX Metadata](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/3dd44d53-847b-402f-a8c7-41a85024caf7)

#### 2. Modern Functions

**Impact**: HIGH - Required for reading modern Excel files

Missing from `openpyxl/utils/formulas.py` FORMULAE tuple:

| Category | Functions |
|----------|-----------|
| **Lambda/UDF** | `LAMBDA`, `LET`, `MAKEARRAY`, `MAP`, `REDUCE`, `SCAN`, `BYROW`, `BYCOL` |
| **Dynamic Array** | `FILTER`, `SORT`, `SORTBY`, `UNIQUE`, `SEQUENCE`, `RANDARRAY` |
| **Lookup** | `XLOOKUP`, `XMATCH` |
| **Text** | `TEXTJOIN`, `CONCAT`, `TEXTBEFORE`, `TEXTAFTER`, `TEXTSPLIT`, `VALUETOTEXT`, `ARRAYTOTEXT` |
| **Aggregate** | `GROUPBY`, `PIVOTBY`, `PERCENTOF` |
| **Logic** | `SWITCH`, `IFS`, `XOR`, `IFNA`, `MAXIFS`, `MINIFS` |
| **Image** | `IMAGE` |
| **Error** | `IFERROR`, `IFNA` (IFNA missing) |

**Notes**:
- All require `_xlfn.` prefix in XML
- Some require `_xlfn._xlws.` prefix (worksheet-scoped)
- Openpyxl already handles `_xlfn.` prefix correctly - just needs function list update

#### 3. Threaded Comments

**Impact**: MEDIUM - Required for modern comment support

| Component | Status | Required Changes |
|-----------|--------|------------------|
| `xl/threadedComments/*.xml` | Not parsed | New reader/writer |
| `xl/persons/person.xml` | Not parsed | New reader/writer |
| `CT_ThreadedComment` | Not implemented | New class |
| `@mentions` | Not supported | Parse mention syntax |
| Legacy comment fallback | Not handled | Preserve legacy copy |

**Files to create**:
- `openpyxl/comments/threaded.py` (NEW)
- `openpyxl/comments/person.py` (NEW)
- Update `openpyxl/packaging/manifest.py`

**Specification**: [MS-XLSX Threaded Comments](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/e0fb917a-1107-409a-852f-13b47aea70dc)

#### 4. Modern Chart Types

**Impact**: MEDIUM - Visual features

Missing chart types:
- Waterfall charts
- Funnel charts
- Treemap charts
- Sunburst charts
- Box & Whisker
- Histogram
- Map charts (geographic)

**Location**: `openpyxl/chart/`

#### 5. Rich Data Types

**Impact**: MEDIUM - Linked data functionality

| Component | Status | Required Changes |
|-----------|--------|------------------|
| `xlRichValue` content type | Not supported | New handler |
| `rdRichValue.xml` | Not parsed | New reader/writer |
| Stocks data type | Not supported | Implement rich value |
| Geography data type | Not supported | Implement rich value |
| `FIELDVALUE` function | Not in FORMULAE | Add to list |

**Specification**: [MS-XLSX Rich Value Data](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/896934fd-8df7-43f4-b154-2d39371c270d)

#### 6. Slicers & Timelines

**Impact**: MEDIUM - Interactive filtering

| Component | Status | Notes |
|-----------|--------|-------|
| Slicer definitions | GUID recognized | `{A8765BA9-456A-4DAB-B4F3-ACF838C121DE}` |
| Slicer cache | Not implemented | |
| Table slicers | Not implemented | |
| Pivot slicers | Not implemented | |
| Timeline | GUID recognized | `{7E03D99C-DC04-49d9-9315-930204A7B6E9}` |

---

## Extension Framework

Openpyxl has infrastructure for extensions but doesn't fully use it:

### Existing Extension Support

```python
# openpyxl/xml/constants.py:114-124
EXT_TYPES = {
    '{78C0D931-6437-407D-A8EE-F0AAD7539E65}': 'Conditional Formatting',
    '{CCE6A557-97BC-4B89-ADB6-D9C93CAAB3DF}': 'Data Validation',
    '{05C60535-1F16-4FD2-B633-F4F36F0B64E0}': 'Sparkline Group',
    '{A8765BA9-456A-4DAB-B4F3-ACF838C121DE}': 'Slicer List',
    '{FC87AEE6-9EDD-4A0A-B7FB-166176984837}': 'Protected Range',
    '{01252117-D84E-4E92-8308-4BE1C098FCBB}': 'Ignored Error',
    '{F7C9EE02-42E1-4005-9D12-6889AFFD525C}': 'Web Extension',
    '{3A4CF648-6AED-40f4-86FF-DC5316D8AED3}': 'Slicer List',
    '{7E03D99C-DC04-49d9-9315-930204A7B6E9}': 'Timeline Ref',
}
```

### extLst Usage

`extLst` (Extension List) appears in 20+ modules but is often just preserved, not parsed:
- `openpyxl/workbook/views.py`
- `openpyxl/chart/*.py` (multiple)
- `openpyxl/worksheet/filters.py`
- etc.

---

## Namespaces

### Current Namespaces (`openpyxl/xml/constants.py`)

```python
SHEET_MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
CHART_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
# ... etc
```

### Missing Namespaces (Need to Add)

```python
# For dynamic arrays and modern features
X14_NS = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
X15_NS = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"
XLDAPR_NS = "http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray"
XLRD_NS = "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"
XLRD2_NS = "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2"
XLTHREADED_NS = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"
```

---

## Recommended Implementation Order

### Phase 1: Foundation
1. Add missing namespaces to constants
2. Update FORMULAE tuple with modern functions
3. Implement metadata.xml infrastructure

### Phase 2: Dynamic Arrays
4. Add cell metadata (`cm` attribute) support
5. Implement `futureMetadata XLDAPR`
6. Update formula tokenizer for `#` operator
7. Handle calcChain for dynamic arrays

### Phase 3: Comments
8. Implement threaded comments reader
9. Implement person.xml support
10. Implement threaded comments writer
11. Handle legacy comment fallback

### Phase 4: Charts & Visualization
12. Add waterfall chart
13. Add funnel chart
14. Add treemap chart
15. Implement sparklines fully

### Phase 5: Advanced Features
16. Implement slicers (read/write)
17. Implement timelines
18. Implement rich data types
19. Enable pivot table creation
20. Enable table creation

---

## Testing Strategy

Each feature needs:

1. **Unit tests**: Test individual classes/functions
2. **Round-trip tests**: Read XLSX, modify, write, read again
3. **Excel validation**: Open in Excel to verify no corruption
4. **Spec compliance**: Validate XML against schemas

### Test Data Sources

- Create test files in Excel 365
- Use existing files in `openpyxl/tests/data/`
- Save minimal examples for each feature

---

## Sources

- [ECMA-376 5th Edition](https://ecma-international.org/publications-and-standards/standards/ecma-376/)
- [MS-XLSX Specification](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/2c5dee00-eff2-4b22-92b6-0738acd4475e)
- [MS-XLSX Metadata](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/3dd44d53-847b-402f-a8c7-41a85024caf7)
- [MS-XLSX Threaded Comments](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/e0fb917a-1107-409a-852f-13b47aea70dc)
- [MS-XLSX Rich Value Data](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/896934fd-8df7-43f4-b154-2d39371c270d)
- [OpenPyXL Dynamic Arrays Discussion](https://groups.google.com/g/openpyxl-users/c/aacD5eRiP7w)
- [XlsxWriter Dynamic Arrays](https://xlsxwriter.readthedocs.io/working_with_formulas.html)
- [SheetJS Comments Documentation](https://docs.sheetjs.com/docs/csf/features/comments/)
