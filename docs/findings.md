# Input Item Interaction Findings

Last updated: 2026-05-14

## Current Code Shape
- Main workbench is in `src/App.tsx`.
- Input metric cells are detected by `isMetricInputCell(address)`: column A, rows 2-13.
- Year/parameter cells are row 1, columns B-H.
- Result cells are represented by `outputConfigs`, typically `B2:H13`.
- Existing suggestion popover appears for active metric input cells when `enableMetricSuggestions` is true.
- Existing typing in input cells changes only `metricDrafts`; selecting a metric via suggestions writes metric binding and output configs.

## Current Color Shape
- `.input-cell` is yellow and used for both metric input cells and year cells.
- `.table-header-cell` is dark blue for row 1 headers.
- `.output-cell` is light blue, but status classes can override colors:
  - pending refresh becomes yellow;
  - failed becomes red;
  - loading becomes blue.
- This creates multiple meanings for color and makes the product hard to explain.

## Product Interpretation
- Input area should mean "what the user asks for": selected metric/function or custom typed item.
- Result area should mean "where values live": backend-function results or user-entered result text/function.
- Recognition status should be shown with text badges or popover copy, not large cell color changes.

