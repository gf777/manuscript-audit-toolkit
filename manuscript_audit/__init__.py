"""
Manuscript Audit Toolkit
========================
Automated line-by-line academic manuscript auditing, pre/post revision alignment,
and native OpenXML (.docx / Google Docs) comment injection while preserving 100%
of live reference manager links (Paperpile, EndNote, Zotero).
"""

__version__ = "1.0.0"

from .auditor import ManuscriptAuditor
from .aligner import ManuscriptAligner
from .commenter import OpenXMLCommentInjector

__all__ = ["ManuscriptAuditor", "ManuscriptAligner", "OpenXMLCommentInjector"]
