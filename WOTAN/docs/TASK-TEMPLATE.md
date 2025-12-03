# Task: T-{CATEGORY}-{NN}

## Header

| Field | Value |
|-------|-------|
| **ID** | T-{CATEGORY}-{NN} |
| **Parent** | B-{CATEGORY}-{NN} / T-{CATEGORY}-{NN} / User request |
| **State** | READY / IN_PROGRESS / BLOCKED / DONE / FAILED / PARTIAL |
| **Created** | YYYY-MM-DD |
| **Updated** | YYYY-MM-DD |

## Objective

Clear, concise statement of what this task accomplishes.

## Acceptance Criteria

- [ ] Criterion 1: Specific, testable requirement
- [ ] Criterion 2: Another requirement
- [ ] Criterion 3: Tests pass
- [ ] Criterion 4: Round-trip verification in Excel

## Context

Reference relevant documents:
- `WOTAN/docs/vision.md` - alignment with project goals
- `WOTAN/docs/gap-analysis.md` - specific gap being addressed
- `WOTAN/docs/architecture.md` - extension points to use
- Related backlog items or prior tasks

### Background

Brief technical context. What exists now? What needs to change?

### Related Files

- `openpyxl/path/to/file.py` - description
- `openpyxl/path/to/test.py` - test location

## Subtasks

| ID | Description | State |
|----|-------------|-------|
| T-{CATEGORY}-{NN}-1 | Subtask 1 | READY |
| T-{CATEGORY}-{NN}-2 | Subtask 2 | READY |

(Remove this section if no subtasks)

## Implementation

### Approach

Describe the implementation strategy.

### Changes Made

1. **File**: `path/to/file.py`
   - Change description

2. **File**: `path/to/other.py`
   - Change description

### Code Notes

Any important implementation details, design decisions, or gotchas.

## Testing

### Unit Tests

- `openpyxl/*/tests/test_*.py::TestClass::test_method` - description

### Round-Trip Tests

- Test file: `openpyxl/tests/data/wotan/feature_test.xlsx`
- Steps:
  1. Load file
  2. Modify X
  3. Save
  4. Reload and verify

### Excel Validation

- [ ] Output file opens in Excel without errors
- [ ] Feature behaves correctly in Excel

## Evidence

### Test Output

```
pytest output here
```

### Verification

Description of manual verification performed.

## Obstacles

(Document any issues encountered)

### Issue 1: Description

- **Observed**: What happened
- **Expected**: What should happen
- **Tried**: What was attempted
- **Hypothesis**: Possible cause
- **Resolution**: How it was fixed (or why it's blocked)

## Outcome

**State**: DONE / FAILED / BLOCKED / PARTIAL

### Summary

Brief summary of what was accomplished.

### Remaining Work

(If PARTIAL or BLOCKED) What still needs to be done.

### Follow-up Items

- New backlog item created: B-{CATEGORY}-{NN}
- Related issue discovered: description

---

## Changelog

| Date | Change |
|------|--------|
| YYYY-MM-DD | Created task |
| YYYY-MM-DD | Updated with implementation |
| YYYY-MM-DD | Marked DONE |
