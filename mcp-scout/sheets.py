# © 2026 Iryna Subbotina. All Rights Reserved. Proprietary — not open source. See LICENSE/NOTICE. Unauthorized copying, redistribution, or claim of authorship is prohibited.
"""Excel Online (Microsoft Graph) — candidate tracker and research articles."""
from __future__ import annotations

import logging
import os
import string
from datetime import date

import httpx

import gauth

log = logging.getLogger(__name__)

CANDIDATE_WORKBOOK_PATH = os.environ.get("CANDIDATE_WORKBOOK_PATH", "").strip() or None
RESEARCH_WORKBOOK_PATH = os.environ.get("RESEARCH_WORKBOOK_PATH", "").strip() or None

CANDIDATES_SHEET = "Candidates"
CANDIDATES_TABLE = "Candidates"
ARTICLES_SHEET = "Articles"
ARTICLES_TABLE = "Articles"
JD_SHEET = "Job Descriptions"
JD_TABLE = "JobDescriptions"

_FORMULA_PREFIXES = ("=", "+", "-", "@")


def _sanitize(value) -> object:
    """Neutralize leading =/+/-/@ so Excel doesn't interpret written values as formulas."""
    if isinstance(value, str) and value.startswith(_FORMULA_PREFIXES):
        return "'" + value
    return value


def _col_letter(n: int) -> str:
    """1-indexed column number -> Excel column letter (1 -> A, 27 -> AA)."""
    letters = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        letters = string.ascii_uppercase[rem] + letters
    return letters


def _item(workbook_path: str) -> str:
    return f"/me/drive/root:{workbook_path}:"


def _workbook_web_url(workbook_path: str) -> str:
    meta = gauth.request("GET", f"{_item(workbook_path)}?$select=webUrl")
    return meta.get("webUrl", "")


def _ensure_table(workbook_path: str, sheet: str, table: str, headers: list[str]) -> None:
    try:
        gauth.request("GET", f"{_item(workbook_path)}/workbook/tables/{table}")
        return
    except httpx.HTTPStatusError as e:
        if e.response.status_code != 404:
            raise

    last_col = _col_letter(len(headers))
    range_address = f"A1:{last_col}1"
    gauth.request(
        "PATCH",
        f"{_item(workbook_path)}/workbook/worksheets/{sheet}/range(address='{range_address}')",
        json={"values": [headers]},
    )
    created = gauth.request(
        "POST",
        f"{_item(workbook_path)}/workbook/worksheets/{sheet}/tables/add",
        json={"address": f"{sheet}!{range_address}", "hasHeaders": True},
    )
    table_id = created.get("id")
    if table_id:
        gauth.request(
            "PATCH",
            f"{_item(workbook_path)}/workbook/tables/{table_id}",
            json={"name": table},
        )


def _append_row(workbook_path: str, sheet: str, table: str, headers: list[str], row: list) -> None:
    _ensure_table(workbook_path, sheet, table, headers)
    sanitized = [_sanitize(v) for v in row]
    gauth.request(
        "POST",
        f"{_item(workbook_path)}/workbook/tables/{table}/rows/add",
        json={"values": [sanitized]},
    )


def track_candidate(
    name: str,
    profile_url: str,
    role: str,
    status: str,
    source: str = "",
    comment: str = "",
) -> dict:
    """Append a candidate row to the candidate tracker Excel workbook.

    Columns: Дата, Ім'я, Профіль URL, Роль, Статус, Джерело, Коментар
    """
    if not CANDIDATE_WORKBOOK_PATH:
        raise RuntimeError(
            "CANDIDATE_WORKBOOK_PATH env var is not set. "
            "Create an Excel workbook in OneDrive and add its path to .env."
        )
    headers = ["Дата", "Ім'я", "Профіль URL", "Роль", "Статус", "Джерело", "Коментар"]
    row = [date.today().isoformat(), name, profile_url, role, status, source, comment]
    _append_row(CANDIDATE_WORKBOOK_PATH, CANDIDATES_SHEET, CANDIDATES_TABLE, headers, row)
    log.info("sheets: tracked candidate '%s' for role '%s'", name, role)
    return {
        "sheet_url": _workbook_web_url(CANDIDATE_WORKBOOK_PATH),
        "candidate": name,
        "role": role,
        "status": status,
    }


def save_research_article(
    title: str,
    url: str,
    summary: str,
    topic: str = "",
    source: str = "",
) -> dict:
    """Append a research article to the research Excel workbook."""
    if not RESEARCH_WORKBOOK_PATH:
        raise RuntimeError(
            "RESEARCH_WORKBOOK_PATH env var is not set. "
            "Create an Excel workbook in OneDrive and add its path to .env."
        )
    headers = ["Дата", "Заголовок", "URL", "Самарі", "Тема", "Джерело"]
    row = [date.today().isoformat(), title, url, summary, topic, source]
    _append_row(RESEARCH_WORKBOOK_PATH, ARTICLES_SHEET, ARTICLES_TABLE, headers, row)
    log.info("sheets: saved article '%s'", title)
    return {
        "sheet_url": _workbook_web_url(RESEARCH_WORKBOOK_PATH),
        "title": title,
    }


def save_job_description(
    title: str,
    content: str,
    company: str = "",
    department: str = "",
    status: str = "Draft",
) -> dict:
    """Save a Job Description to the JD sheet of the candidate tracker Excel workbook.

    Columns: Дата, Роль, Компанія, Відділ, Статус, Текст JD
    """
    workbook_path = CANDIDATE_WORKBOOK_PATH
    if not workbook_path:
        raise RuntimeError(
            "CANDIDATE_WORKBOOK_PATH env var is not set. "
            "Create an Excel workbook in OneDrive and add its path to .env."
        )
    headers = ["Дата", "Роль", "Компанія", "Відділ", "Статус", "Текст JD"]
    row = [date.today().isoformat(), title, company, department, status, content]
    _append_row(workbook_path, JD_SHEET, JD_TABLE, headers, row)
    log.info("sheets: saved JD '%s'", title)
    return {
        "sheet_url": _workbook_web_url(workbook_path),
        "title": title,
        "status": status,
    }


def list_job_descriptions(status: str = "") -> list[dict]:
    """List saved JDs from the Job Descriptions table."""
    workbook_path = CANDIDATE_WORKBOOK_PATH
    if not workbook_path:
        raise RuntimeError("CANDIDATE_WORKBOOK_PATH env var is not set.")
    headers = ["Дата", "Роль", "Компанія", "Відділ", "Статус", "Текст JD"]
    try:
        result = gauth.request("GET", f"{_item(workbook_path)}/workbook/tables/{JD_TABLE}/rows")
    except httpx.HTTPStatusError as e:
        if e.response.status_code == 404:
            return []
        raise
    out = []
    for entry in result.get("value", []):
        row = (entry.get("values") or [[]])[0]
        padded = list(row) + [""] * (len(headers) - len(row))
        item = dict(zip(headers, padded))
        if not status or item.get("Статус", "") == status:
            out.append(item)
    return out
