# WOTAN Test Files

Test files for modern Excel features. These files must be created in Excel 365/2021+.

## Required Test Files

### Dynamic Arrays (`dynamic-arrays/`)

| File | Description | Features to Test |
|------|-------------|------------------|
| `simple-spill.xlsx` | Single UNIQUE formula | Dynamic array spill, metadata.xml |
| `multi-spill.xlsx` | Multiple spill formulas | SORT, FILTER, UNIQUE on same sheet |
| `spill-reference.xlsx` | Spilled range reference | A1# operator in formulas |
| `nested-dynamic.xlsx` | Nested dynamic formulas | SORT(UNIQUE(...)) |

#### How to Create `simple-spill.xlsx`:
1. Open Excel 365
2. In A1:A5, enter: Apple, Banana, Apple, Cherry, Banana
3. In C1, enter: =UNIQUE(A1:A5)
4. Save as .xlsx

### Threaded Comments (`threaded-comments/`)

| File | Description | Features to Test |
|------|-------------|------------------|
| `single-comment.xlsx` | One threaded comment | Basic threadedComment.xml |
| `comment-thread.xlsx` | Comment with replies | parentId relationships |
| `multi-thread.xlsx` | Multiple threads | Multiple authors, person.xml |
| `comment-mention.xlsx` | Comment with @mention | CT_Mention structure |
| `resolved-comment.xlsx` | Resolved/done comment | done="true" attribute |

#### How to Create `single-comment.xlsx`:
1. Open Excel 365
2. Right-click cell A1
3. Select "New Comment" (not "New Note")
4. Type a comment and click away
5. Save as .xlsx

### Modern Charts (`charts/`)

| File | Description | Features to Test |
|------|-------------|------------------|
| `waterfall.xlsx` | Waterfall chart | c:waterfallChart element |
| `funnel.xlsx` | Funnel chart | c:funnelChart element |
| `treemap.xlsx` | Treemap chart | Hierarchical visualization |
| `sunburst.xlsx` | Sunburst chart | Ring hierarchy |
| `box-whisker.xlsx` | Box and whisker | Statistical chart |
| `histogram.xlsx` | Histogram chart | Frequency distribution |

### Slicers and Timelines (`slicers/`)

| File | Description | Features to Test |
|------|-------------|------------------|
| `table-slicer.xlsx` | Table with slicer | slicer.xml, slicerCache.xml |
| `timeline.xlsx` | Table with timeline | timelineCache.xml |

### Rich Data (`rich-data/`)

| File | Description | Features to Test |
|------|-------------|------------------|
| `stock-data.xlsx` | Stock data type | xlRichValue, FIELDVALUE |
| `geography.xlsx` | Geography data type | Linked data types |

### Sparklines (`sparklines/`)

| File | Description | Features to Test |
|------|-------------|------------------|
| `line-sparkline.xlsx` | Line sparkline | x14:sparklineGroup |
| `column-sparkline.xlsx` | Column sparkline | type="column" |
| `winloss-sparkline.xlsx` | Win/loss sparkline | type="stacked" |

## Validation Checklist

After creating each file, verify:

1. [ ] File opens in Excel without errors
2. [ ] Run `python -c "from openpyxl import load_workbook; wb = load_workbook('file.xlsx'); wb.save('out.xlsx')"`
3. [ ] `out.xlsx` opens in Excel without errors
4. [ ] Feature still works correctly in `out.xlsx`

## File Naming Convention

- Use lowercase with hyphens
- Include feature name in filename
- Keep files minimal (only the feature being tested)
