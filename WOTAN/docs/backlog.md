# WOTAN Backlog

Work items for bringing openpyxl to full XLSX compatibility.

## Categories

| Code | Category | Description |
|------|----------|-------------|
| `PRESERVE` | Preservation | Stop losing data on round-trip |
| `DYNARR` | Dynamic Arrays | Spill, metadata, calcChain |
| `FUNC` | Functions | Modern Excel functions |
| `COMMENT` | Comments | Threaded comments, person |
| `CHART` | Charts | New chart types |
| `PIVOT` | Pivot Tables | Creation, full support |
| `RICH` | Rich Data | Stocks, geography, etc |
| `SLICER` | Slicers | Table/pivot slicers |
| `SPARK` | Sparklines | Mini charts in cells |
| `TABLE` | Tables | Table creation |
| `CORE` | Core | Bug fixes, improvements |
| `DOC` | Documentation | Docs and research |
| `TEST` | Testing | Test infrastructure |

---

## Priority 0: Foundation (Stop Losing Data)

**Goal**: Before adding new features, stop destroying existing features on round-trip.

### B-PRESERVE-01: Preserve extLst content `[READY]`

**Intent**: Stop discarding extension list content. This is the #1 cause of data loss.

**Next**: T-PRESERVE-01

**Details**:
Currently `openpyxl/descriptors/excel.py` Extension class only stores URI:
```python
class Extension(Serialisable):
    uri = String()  # Content inside <ext> is DISCARDED
```

Need to:
- Store raw XML content in Extension class
- Write back unchanged content on save
- Still issue warning but PRESERVE the data
- Affects: sparklines, slicers, timelines, etc.

### B-PRESERVE-02: Preserve unknown worksheet XML elements `[READY]`

**Intent**: Stop dropping unknown elements in worksheet reader.

**Next**: Depends on B-PRESERVE-01

**Details**:
Currently `openpyxl/worksheet/_reader.py` silently drops unknown elements:
```python
for _, element in it:
    if tag_name in dispatcher:
        dispatcher[tag_name](element)
    # ELSE: Element gone forever
```

Need to:
- Collect unknown elements during parse
- Store on worksheet object
- Write back in correct position
- Allow future parsing without losing current data

### B-PRESERVE-03: Add Serialisable raw XML storage `[NEEDS-SPEC]`

**Intent**: Allow any Serialisable to preserve unknown child elements.

**Next**: None assigned

**Details**:
- Add `_unknown_children` list to Serialisable base
- Modify `from_tree()` to collect unrecognized children
- Modify `to_tree()` to include unknown children
- Enables incremental feature support

### B-PRESERVE-04: Preserve DrawingML shapes `[NEEDS-SPEC]`

**Intent**: Stop losing textboxes, shapes, annotations.

**Next**: None assigned

**Details**:
Current warning: "DrawingML support is incomplete... Shapes and drawings will be lost."
- Identify what shape types exist
- Store as raw XML if not fully parsed
- Write back unchanged

### B-PRESERVE-05: Create round-trip test suite `[READY]`

**Intent**: Test that modern Excel files survive round-trip without data loss.

**Next**: None assigned

**Details**:
- Create Excel 365 files with each modern feature
- Run through load_workbook() then save()
- Compare input/output XML
- Fail test if content lost (beyond known gaps)
- Store in `openpyxl/tests/data/wotan/`

---

## Priority 1: Critical (Dynamic Arrays & Functions)

### B-DYNARR-01: Add metadata.xml infrastructure `[READY]`

**Intent**: Create the foundation for dynamic array support by implementing metadata.xml file handling.

**Next**: T-DYNARR-01

**Details**:
- Create `openpyxl/packaging/metadata.py`
- Add `METADATA_NS` to `xml/constants.py`
- Add metadata content type
- Add workbook relationship for metadata
- Parse and preserve metadata.xml in round-trips

### B-DYNARR-02: Add cell metadata attribute support `[NEEDS-SPEC]`

**Intent**: Add `cm` attribute to cell elements for dynamic array marking.

**Next**: Depends on B-DYNARR-01

**Details**:
- Modify `openpyxl/cell/_writer.py` to write `cm` attribute
- Modify cell reader to parse `cm` attribute
- Add `metadata_index` property to Cell class
- Link to metadata.xml indices

### B-DYNARR-03: Implement futureMetadata XLDAPR `[NEEDS-SPEC]`

**Intent**: Implement the dynamic array property metadata structure.

**Next**: Depends on B-DYNARR-01

**Details**:
- Implement `FutureMetadata` class
- Handle XLDAPR namespace
- Parse dynamic array flags
- Support bi-directional spill tracking

### B-DYNARR-04: Handle spilled range operator `[NEEDS-SPEC]`

**Intent**: Parse and preserve the `#` operator in formulas.

**Next**: None assigned

**Details**:
- Modify `openpyxl/formula/tokenizer.py`
- `A1#` means "spilled range from A1"
- Ensure round-trip preservation

---

### B-FUNC-01: Add Excel 365 functions to FORMULAE `[READY]`

**Intent**: Update the formula validation list with modern Excel functions.

**Next**: T-FUNC-01

**Details**:
Add to `openpyxl/utils/formulas.py`:
- Lambda functions: `LAMBDA`, `LET`, `MAKEARRAY`, `MAP`, `REDUCE`, `SCAN`, `BYROW`, `BYCOL`
- Dynamic array: `FILTER`, `SORT`, `SORTBY`, `UNIQUE`, `SEQUENCE`, `RANDARRAY`
- Lookup: `XLOOKUP`, `XMATCH`
- Text: `TEXTJOIN`, `CONCAT`, `TEXTBEFORE`, `TEXTAFTER`, `TEXTSPLIT`, `VALUETOTEXT`, `ARRAYTOTEXT`
- Aggregate: `GROUPBY`, `PIVOTBY`, `PERCENTOF`
- Logic: `SWITCH`, `IFS`, `XOR`, `IFNA`, `MAXIFS`, `MINIFS`
- Other: `IMAGE`, `FIELDVALUE`, `STOCKHISTORY`

### B-FUNC-02: Document _xlfn prefix handling `[READY]`

**Intent**: Verify and document how openpyxl handles the `_xlfn.` prefix.

**Next**: None assigned

**Details**:
- Current handling in tokenizer looks correct
- Verify behavior with test files
- Document in architecture.md
- Some functions need `_xlfn._xlws.` prefix

---

## Priority 2: High Value (Comments & Charts)

### B-COMMENT-01: Implement threaded comments reader `[NEEDS-SPEC]`

**Intent**: Read threaded comments from modern Excel files.

**Next**: None assigned

**Details**:
- Create `openpyxl/comments/threaded.py`
- Parse `xl/threadedComments/threadedComment*.xml`
- Implement `ThreadedComment` class per MS-XLSX spec
- Handle parent-child comment relationships

### B-COMMENT-02: Implement person.xml support `[NEEDS-SPEC]`

**Intent**: Read and write person data for threaded comments.

**Next**: Depends on B-COMMENT-01

**Details**:
- Create `openpyxl/comments/person.py`
- Parse `xl/persons/person.xml`
- Track authors across workbook
- Generate new person entries

### B-COMMENT-03: Implement threaded comments writer `[NEEDS-SPEC]`

**Intent**: Write threaded comments to XLSX files.

**Next**: Depends on B-COMMENT-01, B-COMMENT-02

**Details**:
- Add content type for threaded comments
- Write threadedComment*.xml files
- Maintain legacy comment fallback
- Generate proper relationships

### B-COMMENT-04: Fix comment formatting preservation `[READY]`

**Intent**: Preserve font/color formatting when reading comments.

**Next**: None assigned

**Details**:
- Current: text and dimensions only, formatting lost
- Parse rich text formatting in comment text
- Preserve on round-trip

---

### B-CHART-01: Implement waterfall chart `[NEEDS-SPEC]`

**Intent**: Add support for waterfall charts.

**Next**: None assigned

**Details**:
- Create `openpyxl/chart/waterfall_chart.py`
- Research MS-XLSX waterfall spec
- Handle connectors and subtotals

### B-CHART-02: Implement funnel chart `[NEEDS-SPEC]`

**Intent**: Add support for funnel charts.

**Next**: None assigned

**Details**:
- Create `openpyxl/chart/funnel_chart.py`
- Research MS-XLSX funnel spec

### B-CHART-03: Implement treemap chart `[NEEDS-SPEC]`

**Intent**: Add support for treemap charts.

**Next**: None assigned

**Details**:
- Create `openpyxl/chart/treemap_chart.py`
- Hierarchical data visualization

### B-CHART-04: Implement sunburst chart `[NEEDS-SPEC]`

**Intent**: Add support for sunburst charts.

**Next**: None assigned

**Details**:
- Create `openpyxl/chart/sunburst_chart.py`
- Hierarchical ring visualization

### B-CHART-05: Implement box & whisker chart `[NEEDS-SPEC]`

**Intent**: Add support for box and whisker charts.

**Next**: None assigned

**Details**:
- Statistical visualization
- Quartiles, outliers

### B-CHART-06: Implement histogram chart `[NEEDS-SPEC]`

**Intent**: Add support for histogram charts.

**Next**: None assigned

**Details**:
- Frequency distribution
- Automatic binning

---

## Priority 3: Medium Value (Slicers, Rich Data, Sparklines)

### B-SLICER-01: Implement table slicer reader `[NEEDS-SPEC]`

**Intent**: Read slicer definitions from XLSX files.

**Next**: None assigned

**Details**:
- Parse slicer XML structures
- GUID: `{A8765BA9-456A-4DAB-B4F3-ACF838C121DE}`
- Link to table definitions
- Preserve on round-trip

### B-SLICER-02: Implement timeline reader `[NEEDS-SPEC]`

**Intent**: Read timeline definitions from XLSX files.

**Next**: None assigned

**Details**:
- Parse timeline XML structures
- GUID: `{7E03D99C-DC04-49d9-9315-930204A7B6E9}`
- Date-based filtering

### B-SLICER-03: Implement slicer writer `[BLOCKED]`

**Intent**: Write slicers to XLSX files.

**Next**: Depends on B-SLICER-01

**Details**:
- Generate slicer XML
- Create relationships
- Link to tables/pivots

---

### B-RICH-01: Implement xlRichValue reader `[NEEDS-SPEC]`

**Intent**: Read rich data types (stocks, geography).

**Next**: None assigned

**Details**:
- Parse `xlRichValue` content type
- Handle linked data references
- Preserve in round-trips

### B-RICH-02: Add FIELDVALUE function support `[READY]`

**Intent**: Recognize FIELDVALUE function for rich data access.

**Next**: Covered by B-FUNC-01

**Details**:
- Add to FORMULAE tuple
- Used to extract fields from rich data types

---

### B-SPARK-01: Implement sparkline reader `[NEEDS-SPEC]`

**Intent**: Read sparkline definitions.

**Next**: None assigned

**Details**:
- GUID: `{05C60535-1F16-4FD2-B633-F4F36F0B64E0}`
- Parse sparkline group structure
- Track data ranges and display options

### B-SPARK-02: Implement sparkline writer `[BLOCKED]`

**Intent**: Write and create sparklines.

**Next**: Depends on B-SPARK-01

**Details**:
- Generate sparkline XML
- Support line, column, win/loss types

---

## Priority 4: Completeness

### B-PIVOT-01: Enable pivot table creation `[NEEDS-SPEC]`

**Intent**: Allow programmatic creation of pivot tables.

**Next**: None assigned

**Details**:
- Currently read-only
- Complex OOXML structure
- Need cache definition generation

### B-TABLE-01: Enable table creation `[NEEDS-SPEC]`

**Intent**: Allow programmatic creation of tables.

**Next**: None assigned

**Details**:
- Currently read-only
- Simpler than pivots
- Need column definition generation

---

## Infrastructure

### B-TEST-01: Create Excel 365 test file suite `[READY]`

**Intent**: Build comprehensive test files for modern features.

**Next**: None assigned

**Details**:
- Create minimal XLSX files in Excel 365 for each feature
- Dynamic arrays with various functions
- Threaded comments
- Modern charts
- Store in `openpyxl/tests/data/wotan/`

### B-DOC-01: Research MS-XLSX dynamic array spec `[READY]`

**Intent**: Document exact XML structures for dynamic arrays.

**Next**: None assigned

**Details**:
- Download current MS-XLSX PDF
- Extract metadata.xml schema
- Document cell cm attribute
- Create `WOTAN/docs/research/dynamic-arrays.md`

### B-DOC-02: Research MS-XLSX threaded comments spec `[READY]`

**Intent**: Document exact XML structures for threaded comments.

**Next**: None assigned

**Details**:
- Extract threadedComment schema
- Document person.xml format
- Create `WOTAN/docs/research/threaded-comments.md`

---

## Status Legend

- `[READY]` - Can be started immediately
- `[NEEDS-SPEC]` - Requires research/specification work first
- `[BLOCKED]` - Depends on other items
- `[DONE]` - Completed

---

## Task Index

| Task ID | Backlog Item | Status |
|---------|--------------|--------|
| (none yet) | | |
