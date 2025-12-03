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

### B-PRESERVE-01: Preserve extLst content `[DONE]`

**Intent**: Stop discarding extension list content. This is the #1 cause of data loss.

**Completed**: T-PRESERVE-01 (2024-12-03)

**Summary**: Modified Extension class to preserve raw XML via deepcopy. Sparklines, slicers, timelines now survive round-trip.

### B-PRESERVE-02: Preserve unknown worksheet XML elements `[DONE]`

**Intent**: Stop dropping unknown elements in worksheet reader.

**Completed**: T-PRESERVE-02 (2024-12-03)

**Summary**: Added PRESERVE_TAGS for known but unhandled elements (sheetCalcPr, protectedRanges, ignoredErrors, etc.). Elements captured in parse() and written back at correct positions per ECMA-376 schema.

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

### B-PRESERVE-05: Create round-trip test suite `[DONE]`

**Intent**: Test that modern Excel files survive round-trip without data loss.

**Completed**: T-PRESERVE-05 (2024-12-03)

**Summary**: Created `openpyxl/tests/test_roundtrip.py` with 7 tests for round-trip fidelity. Also fixed stdlib/lxml element conversion issue discovered during testing. Test data directory created at `openpyxl/tests/data/wotan/`.

---

## Priority 1: Critical (Dynamic Arrays & Functions)

### B-DYNARR-01: Add metadata.xml infrastructure `[DONE]`

**Intent**: Create the foundation for dynamic array support by implementing metadata.xml file handling.

**Completed**: T-DYNARR-01 (2024-12-03)

**Summary**: Created `openpyxl/packaging/metadata.py` with complete Metadata class hierarchy. Added reading in workbook parser, writing in excel writer with proper relationships and content types. Metadata preserved on round-trip.

### B-DYNARR-02: Add cell metadata attribute support `[DONE]`

**Intent**: Add `cm` attribute to cell elements for dynamic array marking.

**Completed**: 2024-12-03

**Summary**: Cell metadata index (`cm` attribute) fully implemented:
- `openpyxl/cell/_writer.py:26-28` writes `cm` attribute
- `openpyxl/worksheet/_reader.py:231-233` reads `cm` attribute
- `openpyxl/cell/cell.py:302-319` provides `cell_metadata_index` property
- Tests in `openpyxl/cell/tests/test_cell.py` and `test_writer.py`

### B-DYNARR-03: Implement futureMetadata XLDAPR `[NEEDS-SPEC]`

**Intent**: Implement the dynamic array property metadata structure.

**Next**: Depends on B-DYNARR-01

**Details**:
- Implement `FutureMetadata` class
- Handle XLDAPR namespace
- Parse dynamic array flags
- Support bi-directional spill tracking

### B-DYNARR-04: Handle spilled range operator `[DONE]`

**Intent**: Parse and preserve the `#` operator in formulas.

**Completed**: 2024-12-03

**Summary**: Spilled range operator fully implemented in `openpyxl/formula/tokenizer.py:148-177`. The `#` operator is recognized as a postfix operator on range references (e.g., `A1#` means "spill from A1"). Tests in `openpyxl/formula/tests/test_tokenizer.py::test_parse_spilled_range_operator`.

---

### B-FUNC-01: Add Excel 365 functions to FORMULAE `[DONE]`

**Intent**: Update the formula validation list with modern Excel functions.

**Completed**: 2024-12-03

**Summary**: Added 147 Excel 365/2019+ functions to `openpyxl/utils/formulas.py`. Total functions now 499 (352 classic + 147 modern). Includes dynamic array, lambda, lookup, text, statistical, and other modern functions.

### B-FUNC-02: Document _xlfn prefix handling `[DONE]`

**Intent**: Verify and document how openpyxl handles the `_xlfn.` prefix.

**Completed**: 2024-12-03

**Summary**: Verified tokenizer correctly handles `_xlfn.`, `_xlfn._xlws.`, and `_xlpm.` prefixes. Documented behavior in architecture.md including prefix types, tokenization examples, and best practices.

---

## Priority 2: High Value (Comments & Charts)

### B-COMMENT-01: Implement threaded comments reader `[DONE]`

**Intent**: Read threaded comments from modern Excel files.

**Completed**: 2024-12-03

**Summary**: Full threaded comments reader implemented:
- Created `openpyxl/comments/threaded.py` with ThreadedComment, ThreadedCommentList, Mention classes
- Reader in `openpyxl/reader/excel.py:281-284` parses threadedComments XML
- Parent-child relationships tracked via `parentId` attribute
- Tests in `openpyxl/comments/tests/test_threaded.py`

### B-COMMENT-02: Implement person.xml support `[DONE]`

**Intent**: Read and write person data for threaded comments.

**Completed**: 2024-12-03

**Summary**: Full person support implemented:
- Created `openpyxl/comments/person.py` with Person, PersonList classes
- Reader in `openpyxl/reader/excel.py:198-205` parses persons XML
- Writer in `openpyxl/writer/excel.py:233-240` writes persons XML
- Workbook.persons property for author management
- Tests in `openpyxl/comments/tests/test_threaded.py`

### B-COMMENT-03: Implement threaded comments writer `[DONE]`

**Intent**: Write threaded comments to XLSX files.

**Completed**: 2024-12-03

**Summary**: Full threaded comments writer implemented:
- Writer in `openpyxl/writer/excel.py:220-230` writes threadedComments XML
- Content types and relationships properly generated
- Legacy comment fallback maintained (legacy comments preserve threaded comment text)
- Full round-trip verified with real Excel 365 files
- Tests in `openpyxl/comments/tests/test_threaded.py::TestWorkbookIntegration`

### B-COMMENT-04: Fix comment formatting preservation `[DONE]`

**Intent**: Preserve font/color formatting when reading comments.

**Completed**: 2024-12-03

**Summary**: Added `_text_obj` attribute to Comment class to store rich Text object. Modified CommentSheet.comments to preserve Text object with RichText formatting. Modified CommentRecord.from_cell() to use preserved Text object on write. Formatting (bold, font, color) now survives round-trip.

---

### B-CHART-01-06: Preserve modern charts (chartex) `[DONE]`

**Intent**: Preserve modern chart types on round-trip (waterfall, funnel, treemap, sunburst, box & whisker, histogram).

**Completed**: 2024-12-03

**Summary**: Created `openpyxl/chart/chartex.py` with ChartExSpace class for extended chart (chartex) preservation. These charts use a different XML namespace (`http://schemas.microsoft.com/office/drawing/2014/chartex`) than traditional charts. Added CHARTEX_NS, CHARTEX_TYPE, CHARTEX_REL constants to `openpyxl/xml/constants.py`. Charts preserved via raw XML blob pattern for round-trip fidelity. Full parsing would require 500+ lines of new code, so binary preservation approach chosen.

**Note**: Full creation/editing support for these chart types is not yet implemented - only round-trip preservation.

---

## Priority 3: Medium Value (Slicers, Rich Data, Sparklines)

### B-SLICER-01: Implement table slicer reader `[DONE]`

**Intent**: Read slicer definitions from XLSX files.

**Completed**: 2024-12-03

**Summary**: Created `openpyxl/worksheet/slicer.py` with Slicer, SlicerCache, SlicerCacheData classes. Added SLICER_TYPE, SLICER_CACHE_TYPE, SLICER_REL constants. Integrated reading in excel.py reader and writing in excel.py writer. Slicers now preserved on round-trip.

### B-SLICER-02: Implement timeline reader `[DONE]`

**Intent**: Read timeline definitions from XLSX files.

**Completed**: 2024-12-03

**Summary**: Created `openpyxl/worksheet/timeline.py` with Timeline, TimelineCache classes. Added TIMELINE_TYPE, TIMELINE_CACHE_TYPE, TIMELINE_REL constants. Integrated reading in excel.py reader and writing in excel.py writer. Timelines now preserved on round-trip.

### B-SLICER-03: Implement slicer writer `[DONE]`

**Intent**: Write slicers to XLSX files.

**Completed**: 2024-12-03

**Summary**: Implemented `_write_slicer` and `_write_timeline` methods in `openpyxl/writer/excel.py`. Proper content types and relationships generated. Full round-trip support.

---

### B-RICH-01: Implement xlRichValue reader `[DONE]`

**Intent**: Read rich data types (stocks, geography).

**Completed**: 2024-12-03

**Summary**: Created `openpyxl/packaging/richdata.py` with RichDataManager class for binary blob preservation. Handles rdrichvalue.xml, rdrichvaluestructure.xml, rdRichValueTypes.xml, rdarray.xml, richValueRel.xml. Integrated into reader/writer. Rich data preserved on round-trip via binary blob pattern (similar to VBA).

### B-RICH-02: Add FIELDVALUE function support `[DONE]`

**Intent**: Recognize FIELDVALUE function for rich data access.

**Completed**: 2024-12-03 (via B-FUNC-01)

**Summary**: FIELDVALUE included in the 147 Excel 365/2019+ functions added to `openpyxl/utils/formulas.py`.

---

### B-SPARK-01: Implement sparkline reader `[DONE]`

**Intent**: Read sparkline definitions.

**Completed**: 2024-12-03

**Summary**: Created `openpyxl/worksheet/sparkline.py` with Sparkline, SparklineGroup, SparklineGroups, SparklineColor classes. Sparklines read via extensions in worksheet reader and preserved on round-trip. Full color and type support (line, column, stacked/win-loss).

### B-SPARK-02: Implement sparkline writer `[DONE]`

**Intent**: Write and create sparklines.

**Completed**: 2024-12-03

**Summary**: Sparklines written via worksheet extensions. Created test file generation script at `openpyxl/tests/data/wotan/create_sparkline_files.py` demonstrating programmatic sparkline creation.

---

## Priority 4: Completeness

### B-PIVOT-01: Enable pivot table creation `[DONE]`

**Intent**: Allow programmatic creation of pivot tables.

**Completed**: 2024-12-03

**Summary**: Created `openpyxl/pivot/builder.py` with PivotTableConfig and PivotTableBuilder classes. Builder creates TableDefinition and CacheDefinition objects with proper field mappings. Supports row fields, column fields, value fields (with aggregation functions), and filter fields. Fixed issues with CacheSource type parameter and NestedSequence requiring tuples not None.

### B-TABLE-01: Enable table creation `[DONE]`

**Intent**: Allow programmatic creation of tables.

**Completed**: 2024-12-03

**Summary**: Table creation fully implemented:
- `Table.from_headers()` convenience constructor in `openpyxl/worksheet/table.py:296-340`
- Creates tables with proper column definitions, styles, and auto-filters
- Full round-trip support verified
- Tests in `openpyxl/worksheet/tests/test_table.py::TestTable::test_from_headers`

---

## Infrastructure

### B-TEST-01: Create Excel 365 test file suite `[PARTIAL]`

**Intent**: Build comprehensive test files for modern features.

**Status**: Threaded comments test file now available; other features still missing

**Summary**:
- Created directory structure: `openpyxl/tests/data/wotan/`
- Created README.md with specification and status tracking
- **7 programmatic test files** (UNTESTED against real Excel 365):
  - Sparklines (3): Created programmatically
  - Dynamic Arrays (3): From XlsxWriter (MIT License)
  - Pivot Tables (1): From ClosedXML (MIT License)
- **Threaded Comments**: ✅ Real Excel 365 file available
  - `WOTAN/example-docs/Ny sammanställning beräkningsmodeller.xlsx`
  - Contains: 3 authors in person.xml, 15 threaded comments across 6 sheets
  - Features: AD provider IDs, timestamps (`dT`), reply threads (`parentId`)
- **Still missing** (require Excel 365 to create):
  - Modern Charts (chartex), Slicers/Timelines, Rich Data

### B-DOC-01: Research MS-XLSX dynamic array spec `[DONE]`

**Intent**: Document exact XML structures for dynamic arrays.

**Completed**: 2024-12-03

**Summary**: Created comprehensive research document covering metadata.xml structure, XLDAPR namespace, cell `cm` attribute, futureMetadata blocks, and implementation plan.

### B-DOC-02: Research MS-XLSX threaded comments spec `[DONE]`

**Intent**: Document exact XML structures for threaded comments.

**Completed**: 2024-12-03

**Summary**: Created comprehensive research document covering CT_PersonList, CT_Person, CT_ThreadedComments, CT_ThreadedComment, and CT_Mention schemas. Documented content types, relationship URIs, namespaces, GUID format requirements, and implementation plan.

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
| T-PRESERVE-01 | B-PRESERVE-01 | DONE |
| T-PRESERVE-02 | B-PRESERVE-02 | DONE |
| T-DOC-01 | Research | DONE |
| T-FUNC-01 | B-FUNC-01 | DONE |
| T-PRESERVE-05 | B-PRESERVE-05 | DONE |
| T-DOC-01-DA | B-DOC-01 | DONE |
| T-DYNARR-01 | B-DYNARR-01 | DONE |
| T-DYNARR-02 | B-DYNARR-02 | DONE |
| T-DYNARR-04 | B-DYNARR-04 | DONE |
| T-DOC-02 | B-DOC-02 | DONE |
| T-FUNC-02 | B-FUNC-02 | DONE |
| T-COMMENT-01 | B-COMMENT-01 | DONE |
| T-COMMENT-02 | B-COMMENT-02 | DONE |
| T-COMMENT-03 | B-COMMENT-03 | DONE |
| T-COMMENT-04 | B-COMMENT-04 | DONE |
| T-SLICER-01 | B-SLICER-01 | DONE |
| T-SLICER-02 | B-SLICER-02 | DONE |
| T-SLICER-03 | B-SLICER-03 | DONE |
| T-RICH-01 | B-RICH-01 | DONE |
| T-RICH-02 | B-RICH-02 | DONE |
| T-SPARK-01 | B-SPARK-01 | DONE |
| T-SPARK-02 | B-SPARK-02 | DONE |
| T-CHART-01-06 | B-CHART-01-06 | DONE |
| T-PIVOT-01 | B-PIVOT-01 | DONE |
| T-TABLE-01 | B-TABLE-01 | DONE |
| T-TEST-01 | B-TEST-01 | PARTIAL |
