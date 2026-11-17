# Order Maintenance report automation
Public code repo: app.py, om_report.py, requirements.txt, README.md  (NO report, NO token in this repo)
Private data repo: report.xlsx (latest report, about 25 MB)
Streamlit Secrets:
APP_PASSWORD = "choose-a-strong-password"
[GITHUB]
token = "github_pat_..."   # fine-grained, ONLY the private data repo, Contents: Read and write
repo = "yourname/order-maintenance-data"
path = "report.xlsx"
branch = "main"
Daily: MIS selects all agents' EOD files together, checks the date and each agent, clicks Update report.
