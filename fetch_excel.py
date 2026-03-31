"""
fetch_excel.py
──────────────
Downloads the Jobs Status Excel from SharePoint and saves it locally.
Run by GitHub Actions on a schedule to keep the repo copy up to date.

Required environment variables:
    M365_USERNAME         — your M365 email
    M365_PASSWORD         — your M365 password
    SHAREPOINT_SITE_URL   — e.g. https://sap.sharepoint.com/teams/YourTeam
    SHAREPOINT_FILE_URL   — server-relative path to the xlsx file
"""

import os
import sys
from sharepoint_fetcher import fetch_excel

site_url = os.environ["SHAREPOINT_SITE_URL"]
file_url = os.environ["SHAREPOINT_FILE_URL"]
username = os.environ["M365_USERNAME"]
password = os.environ["M365_PASSWORD"]

print(f"Fetching from: {file_url}")

try:
    buf = fetch_excel(site_url=site_url, file_url=file_url, username=username, password=password)
except Exception as exc:
    print(f"ERROR: {exc}", file=sys.stderr)
    sys.exit(1)

out_path = "Jobs_Status_Report_2026.xlsx"
with open(out_path, "wb") as f:
    f.write(buf.read())

size_kb = os.path.getsize(out_path) / 1024
print(f"Saved {out_path} ({size_kb:.1f} KB)")
