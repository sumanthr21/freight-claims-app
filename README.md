# Cash Application - EOD loader

Streamlit app that adds the Cash Application agents' EOD files to the
"Test Equity - Cash Application Specialist" report and saves the report to a private GitHub repo.

## Files (the public code repo holds only these four)
- `app.py` - the web page
- `cash_report.py` - the engine that edits the report
- `requirements.txt` - Python packages
- `README.md` - this file

The report itself (`report.xlsx`) lives in a SEPARATE PRIVATE repo. Never put it in the public repo.

## Streamlit Secrets (Manage app > Settings > Secrets)

    APP_PASSWORD = "choose-a-strong-password"

    [GITHUB]
    token = "github_pat_..."
    repo = "yourname/cash-application-data"
    path = "report.xlsx"
    branch = "main"

`repo` is the PRIVATE data repo. Use straight quotes, no spaces inside them.

## What the app does with each EOD sheet

| EOD sheet               | Goes to report sheet | Task written                          |
|-------------------------|----------------------|---------------------------------------|
| Billtrust               | Billtrust            | Billtrust Review                      |
| TE_ MATCH REVIEW        | Billtrust            | Billtrust Review (Task Type: Match Review) |
| Key offs                | Keyoff               | Keyoff                                |
| Unapplied cash          | Unapplied Cash       | Unapplied Cash                        |
| CCP_TE                  | Credit Card          | Credit Card Processing (Test Equity)  |
| CCP_HISCO               | Credit Card          | Credit Card Processing (Hisco)        |
| COA / Unposted Report / Correlation HISCO | same-named sheet | by matching column names (not yet tested on a real EOD) |

It also fills Month, Date, Week and Name, turns "$ (8,763.86)" style text into real numbers,
corrects spellings (e.g. "Corelation Review"), and sets the Daily date cell on the current month sheet.
Excel recalculates Daily, Weekly, Monthly Comparison and the month sheets when the file is opened.

## Safety checks
- A date + agent + task that is already in the report is refused ("Already loaded") and nothing is changed.
- Unknown agent names, unknown status or task-type values, unreadable amounts and unrecognised sheets give yellow warnings.
- Rows in TE_ MATCH REVIEW without a TRUE tick are not loaded (and you are told).

## New month
There must be a sheet for the month (e.g. `Dec'26`). On the 1st, prepare the new month as you do today,
then put the file in the private repo as `report.xlsx` before loading that day's EOD files.
