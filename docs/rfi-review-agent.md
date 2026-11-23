# IRIS RFI Review Agent

`rfi_review_agent.py` reviews every new site and laboratory RFI for the three
projects with the Claude Message Batches API and writes one row per RFI to the
**AGENT REVIEW** tab of that project's Google Sheet. On Mondays it adds a
**WEEKLY SUMMARY** row per project.

| Project | Workbook | How RFIs are recognised |
| --- | --- | --- |
| 65.5 KM | `1WV58zo-jYgqOTBGOtMAjjlHC3WKdQ83s-XdRFlxZLLU` (RFI Detailed Log) | road codes XX-1##, `IR-QAQC`, "65.5 KM" |
| 163.22 KM Feeder and Rural Roads | `1IjN4W79VEZrAjevG5zZVKyJPuCdGbm_xAbcttTNjfOA` (RFI Detailed Log) | road codes XX-2##, `IRIS-QAQC`, "163.22 KM" |
| Kizimkazi Road Network | `15tSQgfQ0tT45nnD-Z1E6YUAJzyg_IFpV_Kqpt_ofF7A` (IRIS Kizimkazi Road Network - RFI Review) | "KIZIMKAZI" in the ZIP name or RFI, `KZP-` RFI numbers |

The agent only adds rows to its own two tabs. It does not touch the RFI tables,
Physical Progress, or any existing tab.

## One-time setup

1. **API key.** In the Claude Console (platform.claude.com) go to *API keys*,
   create a key named `iris-rfi-review`, and copy it.
2. **GitHub secret.** In the repository, *Settings > Secrets and variables >
   Actions > New repository secret*: name `ANTHROPIC_API_KEY`, paste the key.
   The Gmail and Google service-account secrets are the ones the IRIS RFI Daily
   Run already uses.
3. **Share the Kizimkazi sheet** with the service-account email (the
   `client_email` in the service-account JSON) as **Editor**. The 65.5 and
   163.22 workbooks are already shared with it.
4. **First run.** *Actions > IRIS RFI Review Agent > Run workflow*, tick
   *Dry run*. The log lists every RFI found and its project. Then run again
   without *Dry run*.

## Schedule and cost

- Runs daily at 18:30 EAT, after the 17:00 RFI Daily Run. Weekly summary on Mondays.
- Only RFIs dated in the last 14 days are sent on a run (`since_days` input).
  Each RFI is reviewed once; the state is kept in the Actions cache.
- Batches usually finish within minutes. If one takes longer than 40 minutes,
  the next run collects the results.
- Batch pricing is half the normal API price. A one-page RFI is a few thousand
  tokens, so 100 RFIs should cost roughly USD 1 to 3 (an estimate; check *Cost* in the Console after the first run).

## Local commands

```sh
python rfi_review_agent.py --list       # documents found and their project
python rfi_review_agent.py --dry-run    # what would be sent; no API, no sheets
python rfi_review_agent.py              # review
python rfi_review_agent.py --weekly     # review + weekly summary
python -m unittest tests.test_review_agent -v
```
