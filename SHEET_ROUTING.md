# IRIS workbook routing

All spreadsheet consumers use `project_sheets.py` as their source of workbook IDs.
65.5 RFIs and physical progress now use the same Detailed Log workbook:
`1WV58zo-jYgqOTBGOtMAjjlHC3WKdQ83s-XdRFlxZLLU`.
Physical progress targets the exact `PHYSICAL PROGRESS` tab (currently gid 2133406991).
The retired standalone 65.5 progress workbook is no longer used by these scripts.
RFI imports and status updates resolve explicit IDs before cached IDs or title searches.
RFI table maintenance excludes tables on `PHYSICAL PROGRESS`.

## Verify 65.5 first

With the normal Google credentials configured:

```sh
python update_physical_progress.py --project 65.5 --diagnose
python update_physical_progress.py --project 65.5 --dry-run
```

Review the proposed cells before a live run. The Actions workflow defaults to
65.5 physical progress and dry run; select `physical_progress_live` to write.
The RFI import step still processes both existing project logs.
No spreadsheet data is moved or deleted by this configuration change.

## Move 163.22 when ready

1. Copy its physical progress tab into its own RFI Detailed Log workbook,
   `1IjN4W79VEZrAjevG5zZVKyJPuCdGbm_xAbcttTNjfOA`.
2. Keep the exact tab title `PHYSICAL PROGRESS`; verify copied formulas and references.
3. Set repository **Actions variable** `IRIS_PP_ID_163_22` to that RFI workbook ID.
   For local runs, set the environment variable of the same name.
4. Run diagnostics and dry run with `--project 163.22`, then select that project
   (or `both`) in a manual workflow run when ready to write.

Until that variable is set, 163.22 retains its existing standalone progress workbook.
If the required tab is missing, the progress updater fails instead of choosing an RFI tab.

## Offline checks

```sh
python -m unittest discover -s tests -v
python -m compileall -q .
```
