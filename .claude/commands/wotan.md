# Wotan Task Management

Manage tasks in the current project using the wotan/ directory system.

## Structure

```
project-root/
└── wotan/
    ├── backlog.json      # Task index
    └── dev-log/
        └── T-NNNN.md     # Task details
```

## Commands

Parse the user's command from: $ARGUMENTS

### `/wotan` (no arguments)
Show READY and ONGOING tasks as a clean table: ID, Status, Phase (if ONGOING), Summary.

### `/wotan all`
Show all tasks grouped by status.

### `/wotan add "summary"`
Create a new task.

1. Read `wotan/backlog.json`
2. Create task with next ID (T-NNNN format, zero-padded)
3. Determine dependencies from context (user may describe them naturally)
   - If dependencies specified and any are not DONE → BLOCKED
   - Otherwise → READY
4. Write updated `backlog.json`
5. Create `wotan/dev-log/T-NNNN.md` from template
6. Report: "Created T-NNNN: {summary}"

### `/wotan start [T-NNNN]`
Execute a task through its full lifecycle.

**If no ID given**: Pick first ONGOING task, or first READY task if none ongoing.

**If task not found or none available**: Report and stop.

#### Execution Protocol

1. **Load task** from `backlog.json` and `wotan/dev-log/T-NNNN.md`

2. **Check status**:
   - BLOCKED → Report blockers and stop
   - DONE → Report already complete and stop
   - READY → Set to ONGOING, phase to PLANNING
   - ONGOING → Resume from current phase

3. **Execute phases**:

   **PLANNING**:
   - Review Context and Acceptance Criteria
   - Fill Verification Plan (tests, regression check, manual steps)
   - Write Approach
   - Assess scope: commit-sized? If too big, decompose into predecessor tasks and BLOCK
   - → TESTING_RED (or IMPLEMENTING for non-code tasks)

   **TESTING_RED**:
   - Write tests for each criterion
   - Confirm tests FAIL (red)
   - → IMPLEMENTING

   **IMPLEMENTING**:
   - Write minimal code to make tests pass
   - → TESTING_GREEN

   **TESTING_GREEN**:
   - Run all tests + regression check
   - All pass → REFACTORING
   - Any fail → DEBUGGING

   **REFACTORING**:
   - Improve code quality while tests stay green
   - Keep changes minimal
   - → VERIFYING

   **DEBUGGING** (entered from TESTING_GREEN or REFACTORING):
   - Document obstacle and hypotheses
   - Test fixes
   - After 3 failed attempts → BLOCKED (escalate)
   - Fix found → TESTING_GREEN

   **VERIFYING**:
   - Confirm each Acceptance Criterion is met
   - Collect evidence
   - All met → Complete

4. **Commit**: Stage changes, commit as "T-NNNN: {summary}"

5. **Complete**:
   - Set status to DONE
   - Unblock any successors (tasks with this in their `after` array)
   - Report completion

#### Non-Code Tasks
Skip TESTING_RED, TESTING_GREEN, and REFACTORING phases.

For detailed phase execution, see `docs/task-execution.md`.

## Task Statuses

- **READY**: Can be started (no incomplete predecessors)
- **ONGOING**: In progress (check phase for where)
- **DONE**: Completed
- **BLOCKED**: Waiting on predecessors or external dependency

## Dependency Mechanism

Tasks can depend on other tasks via the `after` array in backlog.json.

**When user describes dependencies naturally:**
- "after the auth work" → find auth task, add its ID to new task's `after` array
- "before we can deploy" → find deploy task, add new task's ID to deploy's `after` array
- "depends on T-0003" → add `"T-0003"` to new task's `after` array

**Status implications:**
- Task with `after` containing non-DONE tasks → BLOCKED
- Task with empty `after` or all DONE → READY

**On task completion:**
- Find all tasks that have this task in their `after` array (successors)
- For each successor, check if ALL its `after` tasks are now DONE
- If yes: set successor to READY, report "T-XXXX is now unblocked"

**Cycle detection:**
- Before adding a dependency, check it won't create a cycle
- Build graph from `after` relationships, DFS from target
- If path leads back to source → reject: "Cannot add: would create cycle"

## backlog.json Format

```json
{
  "version": "2.0.0",
  "next_id": 3,
  "tasks": [
    {"id": "T-0001", "status": "DONE", "summary": "Initial setup"},
    {"id": "T-0002", "status": "ONGOING", "summary": "Add feature", "phase": "IMPLEMENTING"}
  ]
}
```

Optional fields: `phase` (ONGOING only), `after` (array of predecessor IDs), `blocker` (external reason).

## Task File Template

```markdown
# T-NNNN: Summary

**Status**: READY
**Phase**: —

## Context
Why this task exists.

## Acceptance Criteria
- [ ] Criterion 1

## Verification Plan

### Tests to Create
| Test | Proves | Status |
|------|--------|--------|

### Regression Check
```
pytest tests/ -q
```

## Approach
(Filled during PLANNING)

## Implementation Notes

## Obstacles

## Evidence

## Outcome
```

## Execute Now

Parse $ARGUMENTS and execute the appropriate command.
