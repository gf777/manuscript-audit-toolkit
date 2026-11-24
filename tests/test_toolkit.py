"""
End-to-end unit tests for Manuscript Audit Toolkit v2.0:
- Preserving 100% of existing coauthor comments & hyperlinks while injecting
  native <w:del>/<w:ins> Track Changes and [AI AUDIT | ...] third-person comments
- Preserving <w:sectPr> multi-tab boundaries when moving figures inline
- Repairing transparent RGBA PNGs onto solid white RGB backgrounds
- Detecting & repairing Excel auto-date gene corruption and URL-encoded strings
- Parsing Google Docs URLs with `?tab=` IDs
"""

import datetime
import os
import tempfile
import unittest
import zipfile
import xml.etree.ElementTree as ET

import docx
import numpy as np
import openpyxl
from PIL import Image

from manuscript_audit import (
    OpenXMLAuditSession,
    OpenXMLCommentInjector,
    SpreadsheetAuditor,
    W_NS,
    audit_and_repair_docx_figures,
    extract_google_id_and_tab,
    inspect_docx_comments_and_changes,
)
from manuscript_audit.docx_engine import R_NS, _ensure_comments_packaging


def _create_mock_docx_with_comment_and_hyperlink(path: str, include_rgba: bool = False) -> None:
    d = docx.Document()
    d.add_paragraph("Draft Title: Telomere architecture across species")
    d.add_paragraph("This finding underscores the intricate landscape of evolution in 100 samples.")
    d.add_paragraph("Figure 1 | Caption at end of manuscript.")
    d.save(path)

    tmp = tempfile.mkdtemp()
    with zipfile.ZipFile(path, "r") as z:
        z.extractall(tmp)

    doc_xml = os.path.join(tmp, "word", "document.xml")
    tree = ET.parse(doc_xml)
    root = tree.getroot()
    body = root.find(f"{{{W_NS}}}body")
    p0, p1, p2 = body.findall(f"{{{W_NS}}}p")[:3]

    # Attach an existing coauthor comment (id=0, author="Coauthor A") on p0
    p0.insert(0, ET.Element(f"{{{W_NS}}}commentRangeStart", {f"{{{W_NS}}}id": "0"}))
    p0.append(ET.Element(f"{{{W_NS}}}commentRangeEnd", {f"{{{W_NS}}}id": "0"}))

    # Add a mock Zotero/Paperpile <w:hyperlink> inside p1
    hl = ET.SubElement(p1, f"{{{W_NS}}}hyperlink", {f"{{{R_NS}}}id": "rIdMockCite1"})
    hl_r = ET.SubElement(hl, f"{{{W_NS}}}r")
    hl_t = ET.SubElement(hl_r, f"{{{W_NS}}}t")
    hl_t.text = " [1]"

    # Put a <w:sectPr> on p2 (simulating end of Tab 1 in a multi-tab Google Doc export)
    p2_pr = p2.find(f"{{{W_NS}}}pPr")
    if p2_pr is None:
        p2_pr = ET.Element(f"{{{W_NS}}}pPr")
        p2.insert(0, p2_pr)
    ET.SubElement(p2_pr, f"{{{W_NS}}}sectPr")

    tree.write(doc_xml, xml_declaration=True, encoding="utf-8")

    comments_root = ET.Element(f"{{{W_NS}}}comments")
    c0 = ET.SubElement(comments_root, f"{{{W_NS}}}comment", {
        f"{{{W_NS}}}id": "0",
        f"{{{W_NS}}}author": "Senior Coauthor",
        f"{{{W_NS}}}date": "2026-11-01T10:00:00Z",
        f"{{{W_NS}}}initials": "SC",
    })
    cp = ET.SubElement(c0, f"{{{W_NS}}}p")
    cr = ET.SubElement(cp, f"{{{W_NS}}}r")
    ct = ET.SubElement(cr, f"{{{W_NS}}}t")
    ct.text = "Please sharpen the title."
    ET.ElementTree(comments_root).write(
        os.path.join(tmp, "word", "comments.xml"), xml_declaration=True, encoding="utf-8"
    )
    _ensure_comments_packaging(tmp)

    if include_rgba:
        media_dir = os.path.join(tmp, "word", "media")
        os.makedirs(media_dir, exist_ok=True)
        arr = np.zeros((40, 40, 4), dtype=np.uint8)
        arr[10:30, 10:30] = [0, 0, 0, 255]  # Black ink on transparent alpha=0 background
        Image.fromarray(arr, mode="RGBA").save(os.path.join(media_dir, "image1.png"))

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
        for folder, _, files in os.walk(tmp):
            for fn in files:
                fp = os.path.join(folder, fn)
                zout.write(fp, os.path.relpath(fp, tmp))


class TestManuscriptAuditToolkit(unittest.TestCase):
    def test_google_url_tab_extraction(self):
        fid, tid = extract_google_id_and_tab(
            "https://docs.google.com/document/d/1AbCdEfGhIjKlMnOpQrStUvWxYz0123456789/edit?tab=t.abc123xyz"
        )
        self.assertEqual(fid, "1AbCdEfGhIjKlMnOpQrStUvWxYz0123456789")
        self.assertEqual(tid, "t.abc123xyz")

    def test_session_preserves_coauthor_comments_hyperlinks_sectpr_and_repairs_rgba(self):
        with tempfile.TemporaryDirectory() as td:
            in_docx = os.path.join(td, "in.docx")
            out_docx = os.path.join(td, "out.docx")
            _create_mock_docx_with_comment_and_hyperlink(in_docx, include_rgba=True)

            with OpenXMLAuditSession(in_docx, out_docx, comment_prefix="AI AUDIT") as sess:
                paras = sess.engine.body.findall(f"{{{W_NS}}}p")
                sess.engine.apply_phrase_edit(
                    paras[1],
                    target="underscores the intricate landscape of evolution",
                    replacement="demonstrates lineage-specific divergence",
                    category="AI-PROSE DE-SLOPPING",
                    summary="Replaced LLM cliché in Results P12 (per Senior Coauthor Comment 0).",
                )
                # Move Figure 1 inline after paras[0] while preserving <w:sectPr> on paras[1]
                sess.engine.move_elements_inline_preserving_sectpr(
                    [(paras[0], [paras[2]])],
                    fallback_sectpr_target_p=paras[1],
                )
                stats = sess.save()

            self.assertEqual(stats["preserved_existing_comments"], 1)
            self.assertEqual(stats["applied_comments"], 1)
            self.assertEqual(stats["total_comments"], 2)
            self.assertEqual(stats["preserved_hyperlinks"], 1)
            self.assertEqual(stats["repaired_rgba_figures"], ["image1.png"])

            info = inspect_docx_comments_and_changes(out_docx)
            self.assertEqual(info["existing_comment_count"], 2)
            self.assertEqual(info["comments"][0]["id"], 0)
            self.assertEqual(info["comments"][0]["author"], "Senior Coauthor")
            self.assertEqual(info["comments"][1]["id"], 1)
            self.assertTrue(info["comments"][1]["comment_text"].startswith("[AI AUDIT | AI-PROSE DE-SLOPPING]"))
            self.assertNotIn("P12", info["comments"][1]["comment_text"])

            fig_rep = audit_and_repair_docx_figures(out_docx)
            self.assertEqual(fig_rep["figures"][0]["mode"], "RGB")
            self.assertFalse(fig_rep["figures"][0]["has_transparent_pixels"])

    def test_comment_injector_preserves_prior_comments(self):
        with tempfile.TemporaryDirectory() as td:
            in_docx = os.path.join(td, "in.docx")
            out_docx = os.path.join(td, "out.docx")
            _create_mock_docx_with_comment_and_hyperlink(in_docx)

            injector = OpenXMLCommentInjector(in_docx)
            count = injector.inject_audit_comments(
                audit_results=[{
                    "line_id": 2,
                    "section": "RESULTS",
                    "original_line": "underscores the intricate landscape of evolution",
                    "has_issue": True,
                    "error_type": "AI_PROSE",
                    "error_explanation": "Replace formulaic phrasing.",
                    "corrected_line": "shows rapid divergence",
                }],
                output_docx_path=out_docx,
            )
            self.assertEqual(count, 1)
            info = inspect_docx_comments_and_changes(out_docx)
            self.assertEqual(info["existing_comment_count"], 2)
            self.assertEqual(info["track_change_count"], 2)  # 1 <w:del> + 1 <w:ins>

    def test_spreadsheet_auditor_repairs_gene_dates_and_url_encoding(self):
        with tempfile.TemporaryDirectory() as td:
            in_xlsx = os.path.join(td, "supp.xlsx")
            out_xlsx = os.path.join(td, "supp_fixed.xlsx")

            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Table_S1"
            ws.append(["gene_symbol", "start", "end", "description", "status"])
            ws.append([datetime.datetime(2026, 3, 1), 106000000, 106000000, "receptor%2C%20alpha", "IN PROGRESS"])
            wb.save(in_xlsx)

            aud = SpreadsheetAuditor(in_xlsx)
            rep = aud.audit_and_repair(output_xlsx_path=out_xlsx)
            self.assertEqual(rep["issue_count"], 4)
            self.assertEqual(rep["fixes_applied"], 2)

            wb_fixed = openpyxl.load_workbook(out_xlsx)
            ws_fixed = wb_fixed["Table_S1"]
            self.assertEqual(ws_fixed["A2"].value, "MARCH1")
            self.assertEqual(ws_fixed["D2"].value, "receptor, alpha")


if __name__ == "__main__":
    unittest.main()
