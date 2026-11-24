"""
OpenXML Comment & Track-Changes Injector
========================================
High-level injector that embeds structured audit results into `.docx` packages
as BOTH native OpenXML Track Changes (`<w:del>` / `<w:ins>`) AND phrase-anchored
sidebar comments (`<w:commentRangeStart>` / `<w:commentRangeEnd>`), while:

1. Preserving 100% of existing coauthor comments and their anchors (starting new
   comment IDs strictly above `max(existing_comment_ids)`).
2. Preserving 100% of live reference manager hyperlinks (Paperpile, Zotero, EndNote).
3. Compositing transparent RGBA PNG figures onto solid white RGB backgrounds.
4. Optionally pairing each comment with the aligned pre-revision sentence.
"""

import datetime
import os
import shutil
import tempfile
import zipfile
from typing import Any, Dict, List, Optional
from xml.etree import ElementTree as ET

from .aligner import normalize_text, ManuscriptAligner
from .docx_engine import (
    DEFAULT_AUDIT_AUTHOR,
    DEFAULT_AUDIT_INITIALS,
    DocxAuditEngine,
    W_NS,
    XML_NS,
    _ensure_comments_packaging,
    format_audit_comment_text,
)
from .figures import composite_rgba_pngs_in_dir


class OpenXMLCommentInjector:
    def __init__(
        self,
        docx_path: str,
        author: str = DEFAULT_AUDIT_AUTHOR,
        initials: str = DEFAULT_AUDIT_INITIALS,
        comment_prefix: Optional[str] = "AI AUDIT",
    ):
        self.docx_path = docx_path
        self.author = author
        self.initials = initials
        self.comment_prefix = comment_prefix

    def _find_best_anchor(self, p_elem: ET.Element, target_sentence: str):
        """Fallback child-span matcher when a sentence crosses `<w:hyperlink>` barriers."""
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
        pre_revision_docx_path: Optional[str] = None,
        apply_track_changes: bool = True,
        repair_rgba_figures: bool = True,
    ) -> int:
        """
        Injects audit results into `output_docx_path` while preserving 100% of
        existing coauthor comments, Track Changes, and hyperlinks.
        """
        aligner = None
        if pre_revision_docx_path and os.path.exists(pre_revision_docx_path):
            aligner = ManuscriptAligner(pre_revision_docx_path, self.docx_path)

        extract_dir = tempfile.mkdtemp(prefix="manuscript_audit_inject_")
        try:
            with zipfile.ZipFile(self.docx_path, "r") as z:
                z.extractall(extract_dir)

            if repair_rgba_figures:
                composite_rgba_pngs_in_dir(os.path.join(extract_dir, "word", "media"))

            doc_xml_p = os.path.join(extract_dir, "word", "document.xml")
            comments_xml_p = os.path.join(extract_dir, "word", "comments.xml")

            tree = ET.parse(doc_xml_p)
            root = tree.getroot()
            body = root.find(f"{{{W_NS}}}body")

            if os.path.exists(comments_xml_p):
                comments_tree = ET.parse(comments_xml_p)
                comments_root = comments_tree.getroot()
            else:
                comments_root = ET.Element(f"{{{W_NS}}}comments")
                comments_tree = ET.ElementTree(comments_root)

            engine = DocxAuditEngine(
                doc_root=root,
                comments_root=comments_root,
                author=self.author,
                initials=self.initials,
                comment_prefix=self.comment_prefix,
            )

            paragraphs = body.findall(f"{{{W_NS}}}p")
            flagged_items = [
                d for d in audit_results
                if d.get("has_issue") and d.get("section") != "REFERENCES"
            ]
            embedded_count = 0
            date_iso = datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

            for item_idx, item in enumerate(flagged_items):
                orig_draft = item.get("original_line", "").strip()
                corr = item.get("corrected_line", "").strip()
                expl = item.get("error_explanation", "").strip()
                err_type = item.get("error_type", "AUDIT").strip() or "AUDIT"
                line_id = item.get("line_id", item_idx + 1)
                sec = item.get("section", "MANUSCRIPT")

                search_key = normalize_text(orig_draft[:30])
                if not search_key:
                    continue

                matched_p = None
                matched_p_idx = -1
                for p_idx, p in enumerate(paragraphs):
                    if search_key in normalize_text("".join(p.itertext())):
                        matched_p = p
                        matched_p_idx = p_idx
                        break

                if matched_p is None:
                    continue

                pre_sentence = (
                    aligner.find_pre_revision_sentence(matched_p_idx, orig_draft)
                    if aligner else None
                )

                # Attempt native Track Change + phrase-anchored comment when `orig_draft` sits cleanly in direct runs
                if apply_track_changes and corr and corr != orig_draft and not aligner:
                    try:
                        engine.apply_phrase_edit(
                            matched_p,
                            target=orig_draft,
                            replacement=corr,
                            category=f"{err_type} — {sec}",
                            summary=expl or "Suggested revision.",
                            details="",
                        )
                        embedded_count += 1
                        continue
                    except ValueError:
                        # Crosses hyperlink/equation or fuzzy whitespace; fall through to rich comment card
                        pass

                cid = engine._alloc_cid()
                engine.applied_comments += 1
                header_cat = format_audit_comment_text(
                    f"{err_type} — Line {line_id} ({sec})",
                    is_category=True,
                    comment_prefix=self.comment_prefix,
                )

                c_elem = ET.SubElement(comments_root, f"{{{W_NS}}}comment", {
                    f"{{{W_NS}}}id": str(cid),
                    f"{{{W_NS}}}author": self.author,
                    f"{{{W_NS}}}date": date_iso,
                    f"{{{W_NS}}}initials": self.initials,
                })

                def _add_labeled_p(label: str, val: str, italic_val: bool = False):
                    cp = ET.SubElement(c_elem, f"{{{W_NS}}}p")
                    r_lbl = ET.SubElement(cp, f"{{{W_NS}}}r")
                    r_lbl_pr = ET.SubElement(r_lbl, f"{{{W_NS}}}rPr")
                    ET.SubElement(r_lbl_pr, f"{{{W_NS}}}b")
                    t_lbl = ET.SubElement(r_lbl, f"{{{W_NS}}}t")
                    t_lbl.set(f"{{{XML_NS}}}space", "preserve")
                    t_lbl.text = label
                    if val:
                        r_val = ET.SubElement(cp, f"{{{W_NS}}}r")
                        if italic_val:
                            r_val_pr = ET.SubElement(r_val, f"{{{W_NS}}}rPr")
                            ET.SubElement(r_val_pr, f"{{{W_NS}}}i")
                        t_val = ET.SubElement(r_val, f"{{{W_NS}}}t")
                        t_val.set(f"{{{XML_NS}}}space", "preserve")
                        t_val.text = val

                _add_labeled_p(f"[{header_cat}]", "")
                if aligner:
                    _add_labeled_p(
                        "PRE-REVISION ORIGINAL: ",
                        f'"{pre_sentence}"' if pre_sentence else "(New text introduced in revision)",
                        italic_val=True,
                    )
                _add_labeled_p("CURRENT DRAFT TEXT: ", f'"{orig_draft}"')
                if corr and corr != orig_draft:
                    _add_labeled_p("SUGGESTED CORRECTION: ", f'"{corr}"')
                _add_labeled_p("EXPLANATION: ", expl)

                anchor_span = self._find_best_anchor(matched_p, orig_draft)
                range_start = ET.Element(f"{{{W_NS}}}commentRangeStart", {f"{{{W_NS}}}id": str(cid)})
                range_end = ET.Element(f"{{{W_NS}}}commentRangeEnd", {f"{{{W_NS}}}id": str(cid)})
                ref_r = ET.Element(f"{{{W_NS}}}r")
                ref_r_pr = ET.SubElement(ref_r, f"{{{W_NS}}}rPr")
                ET.SubElement(ref_r_pr, f"{{{W_NS}}}rStyle", {f"{{{W_NS}}}val": "CommentReference"})
                ET.SubElement(ref_r, f"{{{W_NS}}}commentReference", {f"{{{W_NS}}}id": str(cid)})

                if anchor_span is not None:
                    start_i, end_i = anchor_span
                    matched_p.insert(start_i, range_start)
                    matched_p.insert(end_i + 2, range_end)
                    matched_p.insert(end_i + 3, ref_r)
                else:
                    p_pr = matched_p.find(f"{{{W_NS}}}pPr")
                    insert_idx = 1 if p_pr is not None else 0
                    matched_p.insert(insert_idx, range_start)
                    matched_p.append(range_end)
                    matched_p.append(ref_r)

                embedded_count += 1

            tree.write(doc_xml_p, xml_declaration=True, encoding="utf-8")
            comments_tree.write(comments_xml_p, xml_declaration=True, encoding="utf-8")
            _ensure_comments_packaging(extract_dir)

            with zipfile.ZipFile(output_docx_path, "w", zipfile.ZIP_DEFLATED) as zip_out:
                for foldername, _, filenames in os.walk(extract_dir):
                    for filename in filenames:
                        filepath = os.path.join(foldername, filename)
                        arcname = os.path.relpath(filepath, extract_dir)
                        zip_out.write(filepath, arcname)

            print(
                f"Generated {output_docx_path}: {embedded_count} new audit items "
                f"({len(engine.existing_cids)} prior coauthor comments preserved)."
            )
            return embedded_count
        finally:
            shutil.rmtree(extract_dir, ignore_errors=True)
