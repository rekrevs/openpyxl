# CLAUDE.md

## Role

Ensure that all non-trivial work is traceable, reproducible, and aligned with the architecture. All meaningful work is a Task or a Backlog item.

## Project: WOTAN

**WOTAN** (Working On The Actual Nuances) is our effort to bring openpyxl to full modern XLSX compatibility. All WOTAN-related files live under `WOTAN/`.

### Vision

Make openpyxl "complete" in its ability to read, understand, and modify XLSX files - including all modern Excel 365 features like dynamic arrays, LAMBDA functions, threaded comments, and rich data types.

### Key Documents

- `WOTAN/docs/vision.md` - Target capabilities and invariants
- `WOTAN/docs/gap-analysis.md` - Current gaps vs modern XLSX
- `WOTAN/docs/architecture.md` - Openpyxl structure and extension points
- `WOTAN/docs/backlog.md` - Future work items
- `WOTAN/docs/implementation-status.md` - Current WOTAN progress
- `WOTAN/dev-log/` - Task files

## 1. Work Types

* **Tasks**
  Self contained work units with objective, acceptance criteria, tests, implementation, evidence, and explicit outcome. May have subtasks.
* **Backlog items**
  Future work seeds in `WOTAN/docs/backlog.md` (`B-*`). Not executable until turned into a Task.

## 2. IDs and Files

All work history lives under `WOTAN/dev-log/`.

* **Backlog items**

  * ID: `B-{CATEGORY}-{NN}`
  * Categories: `DYNARR` (Dynamic Arrays), `FUNC` (Functions), `COMMENT` (Comments), `CHART` (Charts), `PIVOT` (Pivot Tables), `RICH` (Rich Data), `SLICER` (Slicers), `CORE` (Core fixes), `DOC` (Documentation), `TEST` (Testing)

* **Tasks**

  * ID: `T-{CATEGORY}-{NN}`
  * File: `WOTAN/dev-log/T-{CATEGORY}-{NN}.md`
  * Category should match source backlog item.

* **Subtasks**

  * ID: parent ID + `-{N}` (nesting allowed):
    `T-DYNARR-01-1`, `T-DYNARR-01-1-1`, etc.

* **ID allocation**
  For a new Task in a category, find highest `T-{CATEGORY}-NN`, increment `NN`, or start at `01`.

* **Template**
  All tasks and subtasks use `WOTAN/docs/TASK-TEMPLATE.md`.

## 3. Task States

Every Task has one state:

* `READY` - Can be started.
* `IN_PROGRESS` - Being worked on.
* `BLOCKED` - Waiting on subtasks or external dependency.
* `DONE` - All acceptance criteria met.
* `FAILED` - Attempted but not successful.
* `PARTIAL` - Some criteria met, rest deferred.

A parent Task becomes `BLOCKED` when it spawns subtasks and unblocks when they finish.

## 4. Parent Relationships

Each Task declares one parent in its header:

* **Backlog item `B-*`**
  On state change, update `WOTAN/docs/backlog.md`.

* **Task `T-*` (subtask)**
  On state change, update parent Task's Subtasks table.

* **User request**
  No file update required.

Child points to parent; parent lists children.

## 5. When To Create Subtasks

Create subtasks when:

1. Work naturally splits into several distinct subtasks.
2. Debugging is non-trivial and needs structured investigation.
3. A significant part of the task cannot proceed until another non-trivial part completes.

Do not create subtasks when:

* A single focused change will resolve the issue.
* Steps are sequential but tightly cohesive.
* You would create exactly one subtask.

Out of scope discoveries become new backlog items, not subtasks.

## 6. How Tasks Start

Tasks may come from:

1. **Backlog items** `B-*`.
2. **Direct user requests** to fix, debug, implement, or investigate.
3. **Reactive discovery** while working:

   * In scope for current Task: handle within that Task (with subtasks if needed).
   * Out of scope: create or update backlog item.
4. **Subtasks** spawned from parent Tasks.

## 7. Problem Solving and Debugging

When something fails or misbehaves in a non-trivial way (including failing or hanging tests):

1. **Stop** further execution.

2. **Document** under `## Obstacles` in the current Task:
   * Observed error.
   * Expected behavior.
   * What has already been tried.
   * Hypotheses what could be the problem.

3. **Remind yourself about the context**:

   * Read `WOTAN/docs/vision.md` and ensure to keep on track.
   * Find and read other documents relevant to the components you are working on.

4. **Write a problem statement** and a proposed plan.

5. **Decide**:
   * Clear, in scope resolution: continue.
   * Out of scope for task: create or update backlog item, keep main Task focused.
   * No clear path: stop and present full context and problem description to the user.

Never iterate through multiple blind fix attempts without this pause and documentation.

---

## 8. Mapping Natural Language To Actions

The agent infers intent and chooses one of:

1. **Create Task now**
   Phrases like "Fix this now", "Implement dynamic arrays".

   * Create new Task using template and next `T-{CATEGORY}-{NN}`.
   * Fill Objective, Acceptance Criteria, Parent, Context.
   * Start Task lifecycle.

2. **Add to backlog**
   Phrases like "Later", "Add to backlog", "Park this".

   * Create or update `B-XXX-NN` in `WOTAN/docs/backlog.md`.
   * Do not start a Task unless user also requests that.

3. **Work on existing Task**
   Phrases like "Continue that task", "Pick up T-DYNARR-01".

   * Resolve Task by ID or recent context.
   * Open its file and continue lifecycle.

4. **Work from backlog**
   Phrases like "Work on B-FUNC-02", "Take next function backlog item".

   * From backlog item, create Task (with subtasks if multi step).
   * Set Parent to backlog ID.
   * Run Task lifecycle.

5. **Continue working, agent chooses**
   Phrases like "Continue working", "Do the next sensible step".
   Priority:

   1. Resume most recent open Task (Outcome not set).
   2. Otherwise choose a `[READY]` backlog item, create Task.
   3. If nothing is ready, report that and suggest options.

When intent is unclear, ask a brief clarifying question.

## 9. Task Lifecycle

All Tasks follow this sequence.

### 9.1 Create Task file

Using `WOTAN/docs/TASK-TEMPLATE.md`, fill:

* `ID` - next `T-{CATEGORY}-{NN}`.
* `Parent` - backlog ID, Task ID, or "User request".
* `Objective`.
* `Acceptance Criteria`.
* `Context` referencing:

  * `vision.md`, `architecture.md`, `gap-analysis.md`,
  * `implementation-status.md`,
  * relevant backlog or research doc details.

The agent may implicitly create a Task file when a conversation implies substantial code or test work.

### 9.2 Tests first

Design or update tests required to prove the Acceptance Criteria:

* Unit tests in `openpyxl/*/tests/`.
* Integration tests with real XLSX files.
* Regression tests.

Never weaken or remove regression tests without explicit justification in the Task file.

### 9.3 Minimal implementation

Make clean, well-designed and concise implementations that satisfy the Acceptance Criteria and are supportive of the project long term goals.

* Respect openpyxl's existing architecture and patterns.
* Use descriptor-based XML serialization where applicable.
* Avoid speculative abstractions, unused configuration, or unrelated refactoring.

### 9.4 Verification

A Task is complete only when:

1. All new tests pass.
2. The relevant regression suite passes (`pytest openpyxl/`).
3. Round-trip behavior is demonstrated (read XLSX, modify, write, verify in Excel).
4. Evidence (test outputs, file comparisons) is recorded in the Task file.
5. `WOTAN/docs/implementation-status.md` is updated if feature status changed.

Tests may not be skipped without explaining why.

### 9.5 Outcome

Set one Outcome:

* `DONE` - All acceptance criteria met, evidence recorded.
* `FAILED` - Criteria not met; document why.
* `BLOCKED` - Cannot proceed; document blocker.
* `PARTIAL` - Some criteria met; record remaining work.

If not `DONE`:

* Explain in the Task.
* Update related backlog items (`[BLOCKED]`, refined description).
* Mention this explicitly when user next asks to "continue working".

### 9.6 Parent updates

On Task state change:

* Parent `B-*`: update `WOTAN/docs/backlog.md`.
* Parent `T-*`: update parent Task's Subtasks table.
* Parent "User request": no extra file changes.

### 9.7 Commit policy

One Task per commit (or a tight commit cluster).

* Commit message:
  `T-{CATEGORY}-{NN}: short description`
  Example: `T-DYNARR-01: implement dynamic array metadata support`.

## 10. `WOTAN/docs/backlog.md`

Each backlog item contains:

* ID: `B-{CATEGORY}-{NN}`.
* Readiness tag: `[READY]`, `[NEEDS-SPEC]`, `[BLOCKED]`, `[DONE]`.
* Short intent.
* "Next" pointer to Task(s).
* A Details section with context.

Using backlog items:

* Header informs Task Objective and Parent reference.
* Details inform Context and Acceptance Criteria.
* If item is multi step, plan subtasks accordingly.

Maintaining `WOTAN/docs/backlog.md`:

* When a Task resolves an item, mark `[DONE]`.
* If Task partially addresses it, update Details.
* If Task ends `BLOCKED`, mark backlog item `[BLOCKED]` with explanation.

## 11. Agent Behavior Checklist

While working:

* Always map user language to an intent type (create Task, backlog, existing Task, backlog Task, or continue).
* Never perform significant code or test edits without a Task file.
* Use the error handling protocol for non-trivial issues.
* If stuck with no clear plan, stop and present full context to the user.

After finishing work:

* Ensure Task file has updated Objective, Acceptance Criteria, Evidence, and Outcome.
* Ensure parent backlog item or Task is updated.
* Update `WOTAN/docs/implementation-status.md` if capabilities changed.
* Leave the repository in a coherent state (tests passing, docs consistent).

## 12. Other Documents

The basic flow of information through docs is:
vision -> research -> gap-analysis -> architecture -> tasks -> implementation -> reviews

New tasks must reference and align with:

* `WOTAN/docs/vision.md` - target capabilities and invariants.
* `WOTAN/docs/research/*.md` - specification insights.
* `WOTAN/docs/gap-analysis.md` - what's missing.
* `WOTAN/docs/architecture.md` - openpyxl structure and extension points.
* `WOTAN/docs/implementation-status.md` - current WOTAN progress.
* `WOTAN/docs/backlog.md` - future work.

Research docs answer "what does the spec say" for specific features:

* `ms-xlsx-extensions.md` - Microsoft's XLSX extensions
* `dynamic-arrays.md` - Dynamic array specification
* `threaded-comments.md` - Threaded comment format
* `rich-data-types.md` - Stock/Geography/etc data types

Workflow with research docs:

* **Before implementing**

  * Check for relevant research doc, read it, note constraints.
  * If significant work lacks research, consider creating a doc.
* **During implementation**

  * Update research doc if you discover new spec details.
  * Link to Task IDs in evidence fields.
* **After implementation**

  * Update `implementation-status.md` to reflect new features.

## 13. Openpyxl-Specific Guidelines

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

### XML Namespaces
- New namespaces go in `openpyxl/xml/constants.py`
- Follow existing naming conventions (e.g., `*_NS` suffix)

### Formula Support
- Update `openpyxl/utils/formulas.py` FORMULAE tuple for new functions
- Handle `_xlfn.` prefix for extension functions

## General Rules

- After renaming or deleting files, always verify that the new paths are tracked before removing old ones.
- Never delete files unless explicitly instructed or after confirming they are disposable.
- Prefer reading actual XLSX files created by Excel to understand format details.
- When in doubt about spec interpretation, test behavior in Excel first.
