# Input Item Interaction Plan

Last updated: 2026-05-14

## Goal
Improve the Excel workbench interaction for input cells and simplify cell color meaning.

## Scope
- Work on branch `codex/input-item-interaction`.
- Change only the current Excel workbench interaction and styling.
- Do not reintroduce the reverted Excel import / multi-Sheet changes.

## Requested Product Changes
- Input area cells must support two forms:
  - recognized metric/function selection via typeahead search;
  - user-defined custom input.
- Interaction should feel closer to Bloomberg-style command/typeahead behavior: type text, see matching functions/metrics, select a system-backed function or keep the raw typed value.
- Color logic should only distinguish:
  - input area: recognized indicators and custom indicators live here;
  - result area: cells that call backend functions for values, plus user custom text/function input.
- Remove the old mental model:
  - green = identified object;
  - blue = fetchable field;
  - red = unknown/custom;
  - yellow = parameter;
  - white = raw data.

## Implementation Checkpoints
- [x] Inspect current input/output grid implementation.
- [x] Update input cell behavior for typeahead + custom input.
- [x] Simplify visual language to input zone vs result zone.
- [x] Build and verify.
- [x] Restore Excel import/template-home features on the current branch.
- [x] Implement dynamic column header typeahead and column-level system/custom modes.
- [x] Remove pending refresh as a visible intermediate state.
- [x] Add field dictionary and column parameter model.
- [x] Build and smoke-check the updated workbench.

## Rollback Notes
- This branch starts from `0c9004b`.
- If the change is bad, delete or reset branch `codex/input-item-interaction`; `master` remains untouched.
- Current branch also contains restored Excel import/template-home work from the reverted commit, because the user expects the import entrance to remain visible while this interaction work continues.
