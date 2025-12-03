# Task: T-PRESERVE-02

## Header

| Field | Value |
|-------|-------|
| **ID** | T-PRESERVE-02 |
| **Parent** | B-PRESERVE-02 |
| **State** | DONE |
| **Created** | 2024-12-03 |
| **Updated** | 2024-12-03 |

## Objective

Preserve unknown worksheet XML elements during round-trip. Currently, elements not explicitly handled by the parser are silently dropped, causing data loss.

## Acceptance Criteria

- [x] Unknown elements are captured during parse
- [x] Elements are stored on worksheet object
- [x] Elements are written back in correct XML position
- [x] Unit tests verify preservation
- [x] No regression in existing tests (708 tests pass)

## Context

### Elements NOT currently handled (will be dropped)

Per ECMA-376 CT_Worksheet sequence:
- sheetCalcPr (calculation properties)
- protectedRanges
- sortState
- dataConsolidate
- phoneticPr
- customProperties
- cellWatches
- ignoredErrors
- smartTags
- legacyDrawingHF
- picture (background)
- oleObjects
- controls
- webPublishItems

### Current behavior

```python
# openpyxl/worksheet/_reader.py parse()
for _, element in it:
    if tag_name in dispatcher:
        dispatcher[tag_name](element)
    elif tag_name in properties:
        # ...
    elif tag_name == ROW_TAG:
        # ...
    # ELSE: element silently dropped!
```

### Target behavior

Store unknown elements and write them back in the correct position according to the ECMA-376 schema sequence.

## Implementation Plan

1. Add `_unknown_elements` dict to WorkSheetParser (keyed by tag name)
2. In parse(), capture unknown elements via deepcopy
3. Add `unknown_elements` attribute to Worksheet
4. Bind unknown elements in WorksheetReader
5. In WorksheetWriter, write unknown elements at correct positions

## Outcome

### Implementation Summary

Added preservation for known but unhandled worksheet XML elements. The approach uses explicit PRESERVE_TAGS set to avoid capturing child elements of handled parents.

Elements now preserved:
- sheetCalcPr (calculation properties)
- protectedRanges
- dataConsolidate
- phoneticPr
- customProperties
- cellWatches
- ignoredErrors
- smartTags
- drawing (raw element)
- legacyDrawingHF
- picture
- oleObjects
- controls
- webPublishItems

Note: sortState excluded because it also appears as child of autoFilter (which handles it internally).

### Files Changed

- `openpyxl/worksheet/_reader.py` - Added PRESERVE_TAGS, capture in parse()
- `openpyxl/worksheet/_writer.py` - Added UNKNOWN_ELEMENT_POSITIONS, write_unknown_elements()
- `openpyxl/worksheet/worksheet.py` - Added unknown_elements attribute
- `openpyxl/worksheet/tests/test_read_only.py` - Added unknown_elements to std_only
- `openpyxl/worksheet/tests/test_reader.py` - Added test_unknown_elements_preserved
- `openpyxl/worksheet/tests/test_writer.py` - Added test_unknown_elements

### Test Results

All 708 tests pass (worksheet + descriptors tests).
