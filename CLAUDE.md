# CLAUDE.md

## Project: WOTAN

**WOTAN** (Working On The Actual Nuances) is our effort to bring openpyxl to full modern XLSX compatibility.

### Vision

Make openpyxl "complete" in its ability to read, understand, and modify XLSX files - including all modern Excel 365 features like dynamic arrays, LAMBDA functions, threaded comments, and rich data types.

### Key Documents

- `wotan/docs/vision.md` - Target capabilities and invariants
- `wotan/docs/gap-analysis.md` - Current gaps vs modern XLSX
- `wotan/docs/architecture.md` - Openpyxl structure and extension points
- `wotan/docs/backlog.md` - Legacy backlog (detailed work items)
- `wotan/docs/implementation-status.md` - Current WOTAN progress
- `wotan/docs/research/*.md` - Spec research documents

## Openpyxl-Specific Guidelines

### Code Style
- Follow existing openpyxl patterns
- Use descriptors for XML attribute mapping
- Maintain backwards compatibility where possible
- Add type hints for new code

### Testing
- Test files go in `openpyxl/*/tests/`
- Use pytest fixtures
- Include both unit tests and round-trip tests with real XLSX files
- Test data files go in `openpyxl/tests/data/`
- Regression check: `pytest openpyxl/ -q`

### XML Namespaces
- New namespaces go in `openpyxl/xml/constants.py`
- Follow existing naming conventions (e.g., `*_NS` suffix)

### Formula Support
- Update `openpyxl/utils/formulas.py` FORMULAE tuple for new functions
- Handle `_xlfn.` prefix for extension functions

### Verification

A Task is complete only when:
1. All new tests pass
2. Regression suite passes (`pytest openpyxl/`)
3. Round-trip behavior demonstrated (read XLSX, modify, write, verify in Excel)
4. Evidence recorded in Task file
5. `wotan/docs/implementation-status.md` updated if features changed

---

## General Rules

- After renaming or deleting files, verify new paths are tracked before removing old ones
- Never delete files unless explicitly instructed or confirmed disposable
- Prefer reading actual XLSX files created by Excel to understand format details
- When in doubt about spec interpretation, test behavior in Excel first
- Significant code changes require a Task file for traceability
