# Physical Progress workbook routing

Workbook routing is centralized in `project_sheets.py`. Do not duplicate
spreadsheet IDs in individual scripts.

## Current routes

| Project | RFI Detailed Log | Physical Progress | State |
| --- | --- | --- | --- |
| 65.5 | `1WV58zo-jYgqOTBGOtMAjjlHC3WKdQ83s-XdRFlxZLLU` | Same workbook, exact tab `PHYSICAL PROGRESS` (tab ID `2133406991`) | Migrated |
| 163.22 | `1IjN4W79VEZrAjevG5zZVKyJPuCdGbm_xAbcttTNjfOA` | `1YbDCzePl76CqtO9dLl2YZjIZtVmJKSgVkmU6LKqXKUc` | Not migrated |

The retired standalone 65.5 Physical Progress workbook must not be referenced
by Python code. Every Physical Progress reader requires the exact tab title
`PHYSICAL PROGRESS`; it must never fall back to the first worksheet.

## Later migration route for 163.22

Do not set the override until the copied tab has been checked.

1. Copy the 163.22 `PHYSICAL PROGRESS` tab into its existing RFI Detailed Log
   workbook. Preserve formulas, formatting, protections, and sharing.
2. Confirm the copied tab is named exactly `PHYSICAL PROGRESS`. Check its road
   lookup columns, formula references, totals, and a representative sample of
   rows against the standalone workbook.
3. Set the repository Actions variable `IRIS_PP_ID_163_22` to the 163.22 RFI
   workbook ID: `1IjN4W79VEZrAjevG5zZVKyJPuCdGbm_xAbcttTNjfOA`.
4. Run the workflow manually with project `163.22` and leave
   `physical_progress_live` unchecked. Review the dry-run log and confirm that
   it selected the expected workbook and exact tab.
5. Only after the dry run is accepted, run manually with
   `physical_progress_live` checked. Compare the resulting totals before
   retiring the old standalone workbook.

To roll back routing before the old workbook is retired, unset
`IRIS_PP_ID_163_22`; the scripts will use the current standalone 163.22
Physical Progress workbook again.

## Offline verification

These checks do not authenticate to Google or update a spreadsheet:

```text
python -m unittest discover -s tests -v
python -m compileall -q .
```
