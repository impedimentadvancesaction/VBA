# Copy Multiple Data Sources to Destination - README

## Purpose
This macro copies data from multiple external Excel workbooks into destination sheets in the current workbook, using a small mapping table.

It is designed to:
- Pull data from several sources in one run
- Copy efficiently using array-based bulk writes
- Support either replacing existing destination data or appending to it
- Restore Excel application settings even when an error occurs

## Main Procedure
`CopyDataFromMultipleWorkbooks`

High-level flow:
1. Turns off screen updates, events, alerts, and automatic calculation for speed and stability.
2. Defines `sourceMap` entries that control what to copy and where to place it.
3. Loops through each map entry:
   - Validates the source file path
   - Opens the source workbook as read-only
   - Locates source and destination worksheets
   - Detects source used range
   - Loads source range to an in-memory array
   - Writes array to destination in one operation
   - Closes the source workbook
4. Restores Excel settings in `CleanExit`.
5. If an error occurs, shows a message and still runs cleanup via `CleanFail` -> `CleanExit`.

## Mapping Table (`sourceMap`)
Each map row looks like:
`Array("source file path", "source tab name", "destination tab name", appendMode)`

Field meanings:
- Source file path: Full path to external workbook
- Source tab name: Worksheet name in source workbook
- Destination tab name: Worksheet name in current workbook
- `appendMode`:
  - `True` = append below existing destination data
  - `False` = clear destination values and paste starting at row 1

## Data Detection and Copy Behavior
- Source emptiness check uses `Find("*")` on the source worksheet.
- Data boundaries are determined by last used row and last used column.
- Data is read with `.Value2` into `dataArr` for faster transfer.
- Destination write is a single bulk assignment via `Resize(rowsCount, colsCount).Value2 = dataArr`.

## Helper Function
`NextAvailableRow(ws As Worksheet) As Long`

Purpose:
- Returns the next empty row in a destination sheet.
- If the sheet is empty, returns 1.
- Otherwise returns last used row + 1.

Used only when `appendMode = True`.

## Error Handling and Cleanup
- `On Error GoTo CleanFail` routes unexpected errors to a single handler.
- `CleanFail` shows: `Copy failed: <error description>`.
- `CleanExit` always attempts to:
  - Close any open source workbook without saving
  - Restore calculation, alerts, events, and screen updating

## Notes and Assumptions
- Destination sheet names must already exist in the current workbook.
- Source workbook sheet names must match exactly.
- In replace mode, only cell contents are cleared (`ClearContents`), so formatting remains.
