"""
Google Docs & Google Sheets Live Snapshot Utilities
===================================================
Extracts File Identifiers (IDs) and Tab IDs (`?tab=t.xxxxxx`) from Google Docs
and Google Sheets URLs, downloads live `.docx` / `.xlsx` snapshots without
requiring OAuth when link-shared, and detects multi-tab Google Docs so tabs are
never accidentally flattened during in-place updates.
"""

import os
import re
import tempfile
from typing import Any, Dict, Optional, Tuple
import requests

from .docx_engine import inspect_docx_comments_and_changes


def extract_google_id_and_tab(url_or_id: str) -> Tuple[str, Optional[str]]:
    """
    Extracts `(file_id, tab_id)` from a Google Docs/Sheets URL or raw File ID.
    """
    trimmed = url_or_id.strip()
    file_id = trimmed
    m_doc = re.search(r"/d/([a-zA-Z0-9_-]{20,})", trimmed)
    if m_doc:
        file_id = m_doc.group(1)
    else:
        m_param = re.search(r"[?&]id=([a-zA-Z0-9_-]{20,})", trimmed)
        if m_param:
            file_id = m_param.group(1)

    m_tab = re.search(r"[?&]tab=([a-zA-Z0-9._-]+)", trimmed)
    tab_id = m_tab.group(1) if m_tab else None
    return file_id, tab_id


def download_google_snapshot(
    url_or_id: str,
    output_path: str,
    doc_type: str = "auto",
    single_tab_only: bool = False,
    timeout: int = 60,
) -> Dict[str, Any]:
    """
    Downloads a live `.docx` (Google Docs) or `.xlsx` (Google Sheets) export
    from a link-shared Google Workspace URL.

    Parameters
    ----------
    url_or_id : str
        Full Google Docs / Sheets URL or raw File ID.
    output_path : str
        Destination `.docx` or `.xlsx` path.
    doc_type : str
        `"doc"`, `"sheet"`, or `"auto"` (inferred from URL or output extension).
    single_tab_only : bool
        If True and `url_or_id` contains `?tab=<tab_id>`, exports only that tab
        (`export?format=docx&tab=<tab_id>`). If False, exports all tabs so the
        caller can inspect the full multi-tab structure.
    """
    file_id, tab_id = extract_google_id_and_tab(url_or_id)

    if doc_type == "auto":
        if "spreadsheets" in url_or_id or output_path.lower().endswith(".xlsx"):
            doc_type = "sheet"
        else:
            doc_type = "doc"

    if doc_type == "sheet":
        export_url = f"https://docs.google.com/spreadsheets/d/{file_id}/export?format=xlsx"
    else:
        export_url = f"https://docs.google.com/document/d/{file_id}/export?format=docx"
        if single_tab_only and tab_id:
            export_url += f"&tab={tab_id}"

    resp = requests.get(export_url, timeout=timeout)
    resp.raise_for_status()

    parent = os.path.dirname(os.path.abspath(output_path))
    if parent:
        os.makedirs(parent, exist_ok=True)
    with open(output_path, "wb") as f:
        f.write(resp.content)

    return {
        "file_id": file_id,
        "tab_id": tab_id,
        "doc_type": doc_type,
        "single_tab_only": single_tab_only,
        "export_url": export_url,
        "output_path": output_path,
        "bytes_written": len(resp.content),
    }


def detect_google_doc_tabs(url_or_id: str, timeout: int = 60) -> Dict[str, Any]:
    """
    Downloads the full Google Doc `.docx` export (and single-tab export if a
    `?tab=` parameter is present) to determine whether the live Google Doc
    contains multiple tabs that must be preserved via
    `google_apps_script/update_google_docs_preserve_tabs.gs`.
    """
    file_id, tab_id = extract_google_id_and_tab(url_or_id)
    tmp_dir = tempfile.mkdtemp(prefix="gdoc_tab_check_")
    try:
        full_path = os.path.join(tmp_dir, "full_export.docx")
        download_google_snapshot(url_or_id, full_path, doc_type="doc", single_tab_only=False, timeout=timeout)
        full_info = inspect_docx_comments_and_changes(full_path)

        tab_info = None
        if tab_id:
            tab_path = os.path.join(tmp_dir, "single_tab.docx")
            download_google_snapshot(url_or_id, tab_path, doc_type="doc", single_tab_only=True, timeout=timeout)
            tab_info = inspect_docx_comments_and_changes(tab_path)

        is_multi_tab = full_info["estimated_tab_or_section_count"] > 1 or (
            tab_info is not None and full_info["body_element_count"] > tab_info["body_element_count"]
        )
        return {
            "file_id": file_id,
            "target_tab_id": tab_id,
            "is_multi_tab": is_multi_tab,
            "estimated_tab_count": full_info["estimated_tab_or_section_count"],
            "full_doc_elements": full_info["body_element_count"],
            "full_doc_comments": full_info["existing_comment_count"],
            "target_tab_elements": tab_info["body_element_count"] if tab_info else full_info["body_element_count"],
            "target_tab_comments": tab_info["existing_comment_count"] if tab_info else full_info["existing_comment_count"],
        }
    finally:
        for fname in os.listdir(tmp_dir):
            try:
                os.remove(os.path.join(tmp_dir, fname))
            except OSError:
                pass
        try:
            os.rmdir(tmp_dir)
        except OSError:
            pass
