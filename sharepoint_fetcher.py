"""
sharepoint_fetcher.py
─────────────────────
Fetches the Excel file from SharePoint using your Microsoft 365
email + password.  Works on the corporate network (no App Registration needed).

Library: Office365-REST-Python-Client
    pip install Office365-REST-Python-Client

Required secrets (in .streamlit/secrets.toml):
    M365_USERNAME         — your SAP / M365 email
    M365_PASSWORD         — your M365 password
    SHAREPOINT_SITE_URL   — e.g. https://sap.sharepoint.com/teams/YourTeam
    SHAREPOINT_FILE_URL   — server-relative path, e.g.
                            /teams/YourTeam/Shared Documents/General/Report.xlsx
"""

import io
import logging

log = logging.getLogger(__name__)


def fetch_excel(site_url: str, file_url: str, username: str, password: str) -> io.BytesIO:
    """
    Authenticate to SharePoint with email + password and download the Excel file.

    Returns
    -------
    io.BytesIO  Ready to pass directly to pandas.read_excel()
    """
    try:
        from office365.runtime.auth.user_credential import UserCredential
        from office365.sharepoint.client_context import ClientContext
    except ImportError:
        raise ImportError(
            "Missing library. Add to requirements.txt:\n"
            "    Office365-REST-Python-Client"
        )

    log.info("Connecting to SharePoint: %s", site_url)
    try:
        ctx = ClientContext(site_url).with_credentials(
            UserCredential(username, password)
        )
        web = ctx.web
        ctx.load(web)
        ctx.execute_query()
        log.info("Authenticated — site: %s", web.properties.get("Title", "unknown"))
    except Exception as exc:
        raise ConnectionError(
            f"SharePoint login failed. Check M365_USERNAME / M365_PASSWORD in secrets.\n"
            f"Details: {exc}"
        ) from exc

    log.info("Downloading: %s", file_url)
    try:
        buf = io.BytesIO()
        (
            ctx.web
               .get_file_by_server_relative_url(file_url)
               .download(buf)
               .execute_query()
        )
        buf.seek(0)
    except Exception as exc:
        raise FileNotFoundError(
            f"Could not download the file.\n"
            f"SHAREPOINT_FILE_URL must be the server-relative path, e.g.\n"
            f"  /teams/TeamName/Shared Documents/General/Report.xlsx\n"
            f"Details: {exc}"
        ) from exc

    size_kb = buf.getbuffer().nbytes / 1024
    log.info("Downloaded %.1f KB", size_kb)
    return buf
