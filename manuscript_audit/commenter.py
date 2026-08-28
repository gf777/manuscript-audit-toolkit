"""
OpenXML Comment Injector
========================
Injects native, rich-text sidebar comments into Microsoft Word (.docx) documents
anchored precisely to specific sentence ranges while preserving 100% of live
reference manager links (Paperpile, EndNote, Zotero).

Why this is needed:
- Google Docs APIs and standard scripts DO NOT support anchoring comments to text selections.
- Naive docx modification scripts strip <w:hyperlink> or Word fields, breaking live citations.
- This injector operates directly at the OpenXML package level, preserving all relationships
  while inserting valid OpenXML comment range tags.
"""

import os
import shutil
import zipfile
import re
from typing import List, Dict, Any, Optional
from xml.etree import ElementTree as ET

from .aligner import normalize_text, ManuscriptAligner

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_R_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

ET.register_namespace("w", W_NS)
ET.register_namespace("r", R_NS)


class OpenXMLCommentInjector:
    def __init__(self, docx_path: str):
        self.docx_path = docx_path

    def _find_best_anchor(self, p_elem: ET.Element, target_sentence: str):
        """
        Finds child elements (runs/hyperlinks) spanning the target sentence.
        Returns (start_idx, end_idx) or None.
        """
        children = list(p_elem)
        if not children:
            return None

        clean_target = normalize_text(target_sentence[:40])
        if not clean_target:
            return None

        for i, c in enumerate(children):
            c_txt = normalize_text("".join(c.itertext()))
            if clean_target in c_txt or (len(c_txt) > 10 and c_txt in clean_target):
                end_idx = i
                full_span = ""
                for j in range(i, len(children)):
                    child_clean = normalize_text("".join(children[j].itertext()))
                    full_span += child_clean
                    if normalize_text(target_sentence) in full_span or j - i > 6:
                        end_idx = j
                        break
                return (i, end_idx)
        return None

    def inject_audit_comments(
        self,
        audit_results: List[Dict[str, Any]],
        output_docx_path: str,
        pre_revision_docx_path: Optional[str] = None
    ) -> int:
        """
        Injects audit results as native sidebar comments into output_docx_path.
        If pre_revision_docx_path is provided, automatically includes pre-revision original text.
        """
        aligner = None
        if pre_revision_docx_path and os.path.exists(pre_revision_docx_path):
            print(f"Initializing aligner against pre-revision document: {pre_revision_docx_path}")
            aligner = ManuscriptAligner(pre_revision_docx_path, self.docx_path)

        extract_dir = f"_tmp_unpack_{os.path.basename(self.docx_path)}"
        if os.path.exists(extract_dir):
            shutil.rmtree(extract_dir)

        with zipfile.ZipFile(self.docx_path, 'r') as z:
            z.extractall(extract_dir)

        doc_xml_p = os.path.join(extract_dir, "word", "document.xml")
        tree = ET.parse(doc_xml_p)
        root = tree.getroot()
        body = root.find(f"{{{W_NS}}}body")

        comments_root = ET.Element(f"{{{W_NS}}}comments")
        paragraphs = body.findall(f"{{{W_NS}}}p")

        flagged_items = [d for d in audit_results if d.get("has_issue") and d.get("section") != "REFERENCES"]
        embedded_count = 0

        for item_idx, item in enumerate(flagged_items):
            orig_draft = item.get("original_line", "").strip()
            corr = item.get("corrected_line", "").strip()
            expl = item.get("error_explanation", "").strip()
            line_id = item.get("line_id", item_idx + 1)
            sec = item.get("section", "MANUSCRIPT")
            comment_id = str(item_idx + 1)

            search_key = normalize_text(orig_draft[:30])
            if not search_key:
                continue

            matched_p = None
            matched_p_idx = -1
            for p_idx, p in enumerate(paragraphs):
                p_text = normalize_text("".join(p.itertext()))
                if search_key in p_text:
                    matched_p = p
                    matched_p_idx = p_idx
                    break

            if matched_p is not None:
                pre_sentence = aligner.find_pre_revision_sentence(matched_p_idx, orig_draft) if aligner else None

                c_elem = ET.SubElement(comments_root, f"{{{W_NS}}}comment", {
                    f"{{{W_NS}}}id": comment_id,
                    f"{{{W_NS}}}author": "Manuscript Auditor",
                    f"{{{W_NS}}}date": "2026-08-28T12:00:00Z",
                    f"{{{W_NS}}}initials": "AUD"
                })

                # Header
                cp0 = ET.SubElement(c_elem, f"{{{W_NS}}}p")
                cr0 = ET.SubElement(cp0, f"{{{W_NS}}}r")
                cr0_pr = ET.SubElement(cr0, f"{{{W_NS}}}rPr")
                ET.SubElement(cr0_pr, f"{{{W_NS}}}b")
                cr0_t = ET.SubElement(cr0, f"{{{W_NS}}}t")
                cr0_t.text = f"[Audit Line {line_id} ({sec})]"

                # Pre-Revision Original Text
                if aligner:
                    cp_pre = ET.SubElement(c_elem, f"{{{W_NS}}}p")
                    cr_pre_lbl = ET.SubElement(cp_pre, f"{{{W_NS}}}r")
                    cr_pre_lbl_pr = ET.SubElement(cr_pre_lbl, f"{{{W_NS}}}rPr")
                    ET.SubElement(cr_pre_lbl_pr, f"{{{W_NS}}}b")
                    cr_pre_lbl_t = ET.SubElement(cr_pre_lbl, f"{{{W_NS}}}t")
                    cr_pre_lbl_t.text = "PRE-REVISION ORIGINAL: "

                    cr_pre_val = ET.SubElement(cp_pre, f"{{{W_NS}}}r")
                    cr_pre_val_pr = ET.SubElement(cr_pre_val, f"{{{W_NS}}}rPr")
                    ET.SubElement(cr_pre_val_pr, f"{{{W_NS}}}i")
                    cr_pre_val_t = ET.SubElement(cr_pre_val, f"{{{W_NS}}}t")
                    cr_pre_val_t.text = f'"{pre_sentence}"' if pre_sentence else "(New text introduced in revision)"

                # Current Draft Text
                cp1 = ET.SubElement(c_elem, f"{{{W_NS}}}p")
                cr_ai_lbl = ET.SubElement(cp1, f"{{{W_NS}}}r")
                cr_ai_lbl_pr = ET.SubElement(cr_ai_lbl, f"{{{W_NS}}}rPr")
                ET.SubElement(cr_ai_lbl_pr, f"{{{W_NS}}}b")
                cr_ai_lbl_t = ET.SubElement(cr_ai_lbl, f"{{{W_NS}}}t")
                cr_ai_lbl_t.text = "CURRENT DRAFT TEXT: "

                cr_ai_val = ET.SubElement(cp1, f"{{{W_NS}}}r")
                cr_ai_val_t = ET.SubElement(cr_ai_val, f"{{{W_NS}}}t")
                cr_ai_val_t.text = f'"{orig_draft}"'

                # Suggested Correction
                cp2 = ET.SubElement(c_elem, f"{{{W_NS}}}p")
                cr_corr_lbl = ET.SubElement(cp2, f"{{{W_NS}}}r")
                cr_corr_lbl_pr = ET.SubElement(cr_corr_lbl, f"{{{W_NS}}}rPr")
                ET.SubElement(cr_corr_lbl_pr, f"{{{W_NS}}}b")
                cr_corr_lbl_t = ET.SubElement(cr_corr_lbl, f"{{{W_NS}}}t")
                cr_corr_lbl_t.text = "SUGGESTED CORRECTION: "

                cr_corr_val = ET.SubElement(cp2, f"{{{W_NS}}}r")
                cr_corr_val_t = ET.SubElement(cr_corr_val, f"{{{W_NS}}}t")
                cr_corr_val_t.text = f'"{corr}"'

                # Explanation
                cp3 = ET.SubElement(c_elem, f"{{{W_NS}}}p")
                cr_expl_lbl = ET.SubElement(cp3, f"{{{W_NS}}}r")
                cr_expl_lbl_pr = ET.SubElement(cr_expl_lbl, f"{{{W_NS}}}rPr")
                ET.SubElement(cr_expl_lbl_pr, f"{{{W_NS}}}b")
                cr_expl_lbl_t = ET.SubElement(cr_expl_lbl, f"{{{W_NS}}}t")
                cr_expl_lbl_t.text = "EXPLANATION: "

                cr_expl_val = ET.SubElement(cp3, f"{{{W_NS}}}r")
                cr_expl_val_t = ET.SubElement(cr_expl_val, f"{{{W_NS}}}t")
                cr_expl_val_t.text = expl

                # Sentence-Level Anchoring
                anchor_span = self._find_best_anchor(matched_p, orig_draft)
                range_start = ET.Element(f"{{{W_NS}}}commentRangeStart", {f"{{{W_NS}}}id": comment_id})
                range_end = ET.Element(f"{{{W_NS}}}commentRangeEnd", {f"{{{W_NS}}}id": comment_id})
                ref_r = ET.Element(f"{{{W_NS}}}r")
                ref_r_pr = ET.SubElement(ref_r, f"{{{W_NS}}}rPr")
                ET.SubElement(ref_r_pr, f"{{{W_NS}}}rStyle", {f"{{{W_NS}}}val": "CommentReference"})
                ET.SubElement(ref_r, f"{{{W_NS}}}commentReference", {f"{{{W_NS}}}id": comment_id})

                if anchor_span is not None:
                    start_i, end_i = anchor_span
                    matched_p.insert(start_i, range_start)
                    matched_p.insert(end_i + 2, range_end)
                    matched_p.insert(end_i + 3, ref_r)
                else:
                    pPr = matched_p.find(f"{{{W_NS}}}pPr")
                    insert_idx = 1 if pPr is not None else 0
                    matched_p.insert(insert_idx, range_start)
                    matched_p.append(range_end)
                    matched_p.append(ref_r)

                embedded_count += 1

        # Save document.xml
        tree.write(doc_xml_p, xml_declaration=True, encoding="utf-8")

        # Save comments.xml
        comments_xml_path = os.path.join(extract_dir, "word", "comments.xml")
        ET.ElementTree(comments_root).write(comments_xml_path, xml_declaration=True, encoding="utf-8")

        # Update relationships
        doc_rels_path = os.path.join(extract_dir, "word", "_rels", "document.xml.rels")
        rels_tree = ET.parse(doc_rels_path)
        rels_root = rels_tree.getroot()

        has_comment_rel = any(
            rel.get("Type") == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
            for rel in rels_root
        )

        if not has_comment_rel:
            ET.SubElement(rels_root, f"{{{PKG_R_NS}}}Relationship", {
                "Id": "rIdCommentsAudit99",
                "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
                "Target": "comments.xml"
            })
            rels_tree.write(doc_rels_path, xml_declaration=True, encoding="utf-8")

        # Update Content Types
        ct_path = os.path.join(extract_dir, "[Content_Types].xml")
        ct_tree = ET.parse(ct_path)
        ct_root = ct_tree.getroot()

        has_comment_ct = any(
            override.get("PartName") == "/word/comments.xml"
            for override in ct_root.findall(f"{{{CT_NS}}}Override")
        )

        if not has_comment_ct:
            ET.SubElement(ct_root, f"{{{CT_NS}}}Override", {
                "PartName": "/word/comments.xml",
                "ContentType": "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
            })
            ct_tree.write(ct_path, xml_declaration=True, encoding="utf-8")

        # Repack into clean zip
        with zipfile.ZipFile(output_docx_path, 'w', zipfile.ZIP_DEFLATED) as zip_out:
            for foldername, subfolders, filenames in os.walk(extract_dir):
                for filename in filenames:
                    filepath = os.path.join(foldername, filename)
                    arcname = os.path.relpath(filepath, extract_dir)
                    zip_out.write(filepath, arcname)

        if os.path.exists(extract_dir):
            shutil.rmtree(extract_dir)

        print(f"Successfully generated {output_docx_path} with {embedded_count} native OpenXML comments.")
        return embedded_count
