# Hisco Freight Claims report automation
Repo (private) contains: app.py, freight_report.py, requirements.txt, report.xlsx (latest report).
Streamlit Secrets:
APP_PASSWORD = "choose-a-strong-password"
[GITHUB]
token = "github_pat_..."   # fine-grained token, ONLY this repo, Contents: Read and write
repo = "yourname/freight-claims-report"
path = "report.xlsx"
branch = "main"
Daily: MIS selects both agents' EOD files together -> Update report -> download.
