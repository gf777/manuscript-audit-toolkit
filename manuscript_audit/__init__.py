"""
Manuscript Audit Toolkit
========================
End-to-end scientific manuscript auditing, AI-prose de-slopping, spreadsheet &
figure forensics, non-destructive OpenXML Track Changes (`<w:del>`/`<w:ins>`) +
phrase-anchored comment injection, and zero-tab-merge in-place Google Docs upload.
"""

__version__ = "2.0.0"

from .auditor import ManuscriptAuditor
from .aligner import ManuscriptAligner
from .commenter import OpenXMLCommentInjector
from .docx_engine import (
    W_NS,
    R_NS,
    DocxAuditEngine,
    OpenXMLAuditSession,
    inspect_docx_comments_and_changes,
    format_audit_comment_text,
)
from .figures import audit_and_repair_docx_figures, composite_rgba_pngs_in_dir
from .spreadsheets import SpreadsheetAuditor
from .google_docs import (
    extract_google_id_and_tab,
    download_google_snapshot,
    detect_google_doc_tabs,
)

__all__ = [
    "W_NS",
    "R_NS",
    "ManuscriptAuditor",
    "ManuscriptAligner",
    "OpenXMLCommentInjector",
    "DocxAuditEngine",
    "OpenXMLAuditSession",
    "inspect_docx_comments_and_changes",
    "format_audit_comment_text",
    "audit_and_repair_docx_figures",
    "composite_rgba_pngs_in_dir",
    "SpreadsheetAuditor",
    "extract_google_id_and_tab",
    "download_google_snapshot",
    "detect_google_doc_tabs",
]
