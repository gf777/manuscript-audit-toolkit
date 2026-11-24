"""
OpenXML Track Changes & Phrase-Anchored Comment Engine
======================================================
Low-level, non-destructive Word OpenXML (.docx) manipulation engine that injects
BOTH native Track Changes (`<w:del>` / `<w:ins>`, rendered as green Suggesting-mode
cards in Google Docs) AND phrase-anchored sidebar comments (`<w:commentRangeStart>` /
`<w:commentRangeEnd>`) while strictly preserving:

1. 100% of existing coauthor comments and their exact anchor spans
   (assigning new comment IDs strictly starting at `max(existing_ids) + 1`).
2. 100% of existing coauthor Track Changes (`<w:ins>`, `<w:del>`).
3. 100% of live reference manager hyperlinks (`<w:hyperlink>` for Paperpile, Zotero,
   EndNote, DOI/GitHub URLs) and run-level formatting (`<w:rPr>` bold, italics,
   superscripts, subscripts, font styles).
4. Multi-tab Google Docs section boundaries (`<w:sectPr>` inside `<w:pPr>`) when
   relocating figures inline into the Results section.
"""

import copy
import datetime
import os
import re
import shutil
import tempfile
import zipfile
from typing import Any, Dict, List, Optional, Sequence, Tuple
import xml.etree.ElementTree as ET

from .figures import composite_rgba_pngs_in_dir

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_R_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
XML_NS = "http://www.w3.org/XML/1998/namespace"
W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"
W15_NS = "http://schemas.microsoft.com/office/word/2012/wordml"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"

ET.register_namespace("w", W_NS)
ET.register_namespace("r", R_NS)
ET.register_namespace("w14", W14_NS)
ET.register_namespace("w15", W15_NS)
ET.register_namespace("wp", WP_NS)
ET.register_namespace("a", A_NS)
ET.register_namespace("pic", PIC_NS)
ET.register_namespace("m", M_NS)

DEFAULT_AUDIT_AUTHOR = "Manuscript Audit (AI)"
DEFAULT_AUDIT_INITIALS = "AI"


def format_audit_comment_text(
    text: str,
    is_category: bool = False,
    comment_prefix: Optional[str] = "AI AUDIT",
    strip_internal_p_indices: bool = True,
) -> str:
    """
    Formats comment headers and bodies so AI audit comments are unmistakably
    distinguished from prior coauthor comments (even after Google Docs `.docx`
    conversion, which attributes card headers to the uploader's Google account).
    """
    if not text:
        return text
    s = text

    if is_category:
        if strip_internal_p_indices:
            s = re.sub(r"\s+IN\s+P\d+(?:[–\-/]\s*P?\d+)*\b", " IN METHODS", s)
            s = re.sub(r"\bP\d+(?:[–\-/]\s*P?\d+)*\b", "", s)
        s = re.sub(r"\s{2,}", " ", s).strip(" -–&/")
        if comment_prefix:
            if not s.upper().startswith(comment_prefix.upper()):
                return f"{comment_prefix} | {s}" if s else comment_prefix
        return s

    if strip_internal_p_indices:
        s = re.sub(r"\b(Abstract|Introduction|Intro|Results|Discussion|Methods)\s+P\d+(?:[–\-/]\s*P?\d+)*", r"\1", s)
        s = re.sub(r"\(formerly P\d+[–\-]P?\d+ after Methods\)", "(formerly placed as a standalone section after Methods)", s)
        s = re.sub(r"\(P\d+[–\-]P?\d+\)", "", s)
        s = re.sub(r"\(P\d+\)", "", s)
        s = re.sub(r"\bin P\d+(?:[–\-]P?\d+)?\b", "in that section", s)
        s = re.sub(r"\bfrom P\d+(?:[–\-]P?\d+)?\b", "from below", s)
        s = re.sub(r"\bP\d+(?:[–\-/]P?\d+)+\b", "those paragraphs", s)
        s = re.sub(r"\bP\d+\b", "this section", s)
        s = s.replace("Intro ", "Introduction ")

    s = re.sub(r"  +", " ", s)
    s = re.sub(r" \.", ".", s)
    s = re.sub(r" ,", ",", s)
    return s.strip()


def normalize_paragraph_runs(p_elem: ET.Element) -> None:
    """
    Splits any `<w:r>` that contains multiple `<w:t>` or `<w:br>` children into
    separate `<w:r>` elements (each with cloned `<w:rPr>`), so every text run has
    at most one `<w:t>` for exact character-offset boundary splitting.
    """
    new_children = []
    modified = False
    for child in list(p_elem):
        if child.tag == f"{{{W_NS}}}r":
            r_pr = child.find(f"{{{W_NS}}}rPr")
            other_children = [c for c in child if c.tag != f"{{{W_NS}}}rPr"]
            t_or_br = [c for c in other_children if c.tag in (f"{{{W_NS}}}t", f"{{{W_NS}}}br")]
            if len(other_children) > 1 and len(t_or_br) == len(other_children):
                modified = True
                for sub in other_children:
                    new_r = ET.Element(f"{{{W_NS}}}r", dict(child.attrib))
                    if r_pr is not None:
                        new_r.append(copy.deepcopy(r_pr))
                    new_r.append(copy.deepcopy(sub))
                    new_children.append(new_r)
                continue
        new_children.append(child)
    if modified:
        p_elem[:] = new_children


def make_run_with_text(
    text: str,
    r_pr_template: Optional[ET.Element] = None,
    is_del: bool = False,
) -> ET.Element:
    """Creates a `<w:r>` containing `<w:t>` (or `<w:delText>`) with cloned formatting."""
    r = ET.Element(f"{{{W_NS}}}r")
    if r_pr_template is not None:
        r.append(copy.deepcopy(r_pr_template))
    tag = f"{{{W_NS}}}delText" if is_del else f"{{{W_NS}}}t"
    t = ET.SubElement(r, tag)
    t.set(f"{{{XML_NS}}}space", "preserve")
    t.text = text
    return r


def append_comment_xml(
    comments_root: ET.Element,
    cid: int,
    category: str,
    summary: str,
    details: str = "",
    author: str = DEFAULT_AUDIT_AUTHOR,
    initials: str = DEFAULT_AUDIT_INITIALS,
    date_iso: Optional[str] = None,
    comment_prefix: Optional[str] = "AI AUDIT",
    strip_internal_p_indices: bool = True,
) -> ET.Element:
    """Appends a formatted `<w:comment>` entry to `word/comments.xml`."""
    date_str = date_iso or datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    cat_fmt = format_audit_comment_text(
        category,
        is_category=True,
        comment_prefix=comment_prefix,
        strip_internal_p_indices=strip_internal_p_indices,
    )
    sum_fmt = format_audit_comment_text(
        summary,
        is_category=False,
        comment_prefix=comment_prefix,
        strip_internal_p_indices=strip_internal_p_indices,
    )
    det_fmt = format_audit_comment_text(
        details,
        is_category=False,
        comment_prefix=comment_prefix,
        strip_internal_p_indices=strip_internal_p_indices,
    )

    c_elem = ET.SubElement(comments_root, f"{{{W_NS}}}comment", {
        f"{{{W_NS}}}id": str(cid),
        f"{{{W_NS}}}author": author,
        f"{{{W_NS}}}date": date_str,
        f"{{{W_NS}}}initials": initials,
    })

    p1 = ET.SubElement(c_elem, f"{{{W_NS}}}p")
    r1_b = ET.SubElement(p1, f"{{{W_NS}}}r")
    r1_b_pr = ET.SubElement(r1_b, f"{{{W_NS}}}rPr")
    ET.SubElement(r1_b_pr, f"{{{W_NS}}}b")
    t1_b = ET.SubElement(r1_b, f"{{{W_NS}}}t")
    t1_b.set(f"{{{XML_NS}}}space", "preserve")
    t1_b.text = f"[{cat_fmt}] {sum_fmt}" if cat_fmt else sum_fmt

    if det_fmt:
        for line in det_fmt.split("\n"):
            if not line.strip():
                continue
            p2 = ET.SubElement(c_elem, f"{{{W_NS}}}p")
            r2 = ET.SubElement(p2, f"{{{W_NS}}}r")
            t2 = ET.SubElement(r2, f"{{{W_NS}}}t")
            t2.set(f"{{{XML_NS}}}space", "preserve")
            t2.text = line

    return c_elem


def inspect_docx_comments_and_changes(docx_path: str) -> Dict[str, Any]:
    """
    Inventories all existing coauthor comments (`w:id`, `w:author`, `w:date`,
    anchored text span, comment body), existing Track Changes (`<w:ins>`, `<w:del>`),
    hyperlinks, and section breaks (`<w:sectPr>`, corresponding to Google Docs tabs).
    """
    with zipfile.ZipFile(docx_path, "r") as z:
        doc_root = ET.fromstring(z.read("word/document.xml"))
        comments_root = (
            ET.fromstring(z.read("word/comments.xml"))
            if "word/comments.xml" in z.namelist()
            else ET.Element(f"{{{W_NS}}}comments")
        )

    # Extract anchored text spans per comment ID
    active_ids: set = set()
    anchor_texts: Dict[str, List[str]] = {}
    for el in doc_root.iter():
        if el.tag == f"{{{W_NS}}}commentRangeStart":
            cid = el.attrib.get(f"{{{W_NS}}}id", "")
            active_ids.add(cid)
            anchor_texts.setdefault(cid, [])
        elif el.tag == f"{{{W_NS}}}commentRangeEnd":
            cid = el.attrib.get(f"{{{W_NS}}}id", "")
            active_ids.discard(cid)
        elif el.tag in (f"{{{W_NS}}}t", f"{{{W_NS}}}delText") and el.text:
            for cid in active_ids:
                anchor_texts.setdefault(cid, []).append(el.text)

    comments_list = []
    for c in comments_root.findall(f"{{{W_NS}}}comment"):
        cid = c.attrib.get(f"{{{W_NS}}}id", "")
        comments_list.append({
            "id": int(cid) if cid.isdigit() else cid,
            "author": c.attrib.get(f"{{{W_NS}}}author", ""),
            "initials": c.attrib.get(f"{{{W_NS}}}initials", ""),
            "date": c.attrib.get(f"{{{W_NS}}}date", ""),
            "anchored_text": "".join(anchor_texts.get(cid, [])),
            "comment_text": "\n".join("".join(p.itertext()) for p in c.findall(f"{{{W_NS}}}p")),
        })

    track_changes = []
    for el in doc_root.iter():
        if el.tag in (f"{{{W_NS}}}ins", f"{{{W_NS}}}del"):
            wid = el.attrib.get(f"{{{W_NS}}}id", "")
            track_changes.append({
                "id": int(wid) if wid.isdigit() else wid,
                "type": "ins" if el.tag.endswith("}ins") else "del",
                "author": el.attrib.get(f"{{{W_NS}}}author", ""),
                "date": el.attrib.get(f"{{{W_NS}}}date", ""),
                "text": "".join(el.itertext()),
            })

    body = doc_root.find(f"{{{W_NS}}}body")
    body_elems = list(body) if body is not None else []
    sect_breaks = [
        idx for idx, el in enumerate(body_elems)
        if el.tag == f"{{{W_NS}}}sectPr" or el.find(f".//{{{W_NS}}}sectPr") is not None
    ]
    hyperlinks = doc_root.findall(f".//{{{W_NS}}}hyperlink")

    return {
        "docx_path": docx_path,
        "body_element_count": len(body_elems),
        "estimated_tab_or_section_count": max(1, len(sect_breaks)),
        "section_break_indices": sect_breaks,
        "hyperlink_count": len(hyperlinks),
        "existing_comment_count": len(comments_list),
        "max_comment_id": max((c["id"] for c in comments_list if isinstance(c["id"], int)), default=-1),
        "comments": comments_list,
        "track_change_count": len(track_changes),
        "track_changes": track_changes,
    }


class DocxAuditEngine:
    """
    Phrase-level OpenXML Track Changes (`<w:del>`/`<w:ins>`) and sidebar comment
    injection engine.
    """

    def __init__(
        self,
        doc_root: ET.Element,
        comments_root: ET.Element,
        start_comment_id: Optional[int] = None,
        start_change_id: Optional[int] = None,
        author: str = DEFAULT_AUDIT_AUTHOR,
        initials: str = DEFAULT_AUDIT_INITIALS,
        date_iso: Optional[str] = None,
        comment_prefix: Optional[str] = "AI AUDIT",
        normalize_body_limit: Optional[int] = None,
    ):
        self.doc_root = doc_root
        self.body = doc_root.find(f"{{{W_NS}}}body")
        self.comments_root = comments_root
        self.author = author
        self.initials = initials
        self.date_iso = date_iso or datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        self.comment_prefix = comment_prefix

        existing_cids = [
            int(c.attrib[f"{{{W_NS}}}id"])
            for c in comments_root.findall(f"{{{W_NS}}}comment")
            if c.attrib.get(f"{{{W_NS}}}id", "").isdigit()
        ]
        existing_wids = [
            int(el.attrib[f"{{{W_NS}}}id"])
            for el in doc_root.iter()
            if el.tag in (f"{{{W_NS}}}ins", f"{{{W_NS}}}del") and el.attrib.get(f"{{{W_NS}}}id", "").isdigit()
        ]

        self.existing_cids = sorted(existing_cids)
        self.next_cid = start_comment_id if start_comment_id is not None else (max(existing_cids, default=-1) + 1)
        self.next_wid = start_change_id if start_change_id is not None else (max(existing_wids + [99]) + 1)
        self.applied_edits = 0
        self.applied_comments = 0

        if self.body is not None:
            elems = list(self.body)
            target_elems = elems[:normalize_body_limit] if normalize_body_limit is not None else elems
            for elem in target_elems:
                if elem.tag == f"{{{W_NS}}}p":
                    normalize_paragraph_runs(elem)

    def _alloc_cid(self) -> int:
        cid = self.next_cid
        self.next_cid += 1
        return cid

    def _alloc_wid(self) -> int:
        wid = self.next_wid
        self.next_wid += 1
        return wid

    def apply_phrase_edit(
        self,
        p_elem: ET.Element,
        target: str,
        replacement: str,
        category: Optional[str] = None,
        summary: Optional[str] = None,
        details: Optional[str] = None,
    ) -> bool:
        """
        Matches `target` across contiguous direct `<w:r>` children of `p_elem`
        (treating `<w:hyperlink>`, `<w:ins>`, `<w:del>`, `<m:oMath>`, `<w:drawing>`,
        and `<w:br>` as safe barriers so live links and equations are never flattened),
        splits boundary runs at exact character offsets while cloning `<w:rPr>`,
        and injects `<w:del>`, `<w:ins>`, and an optional phrase-anchored comment.
        """
        normalize_paragraph_runs(p_elem)
        children = list(p_elem)

        char_map: List[Tuple[int, int]] = []
        full_text: List[str] = []

        for idx, ch in enumerate(children):
            tag = ch.tag.split("}")[-1]
            if tag == "r":
                t_el = ch.find(f"{{{W_NS}}}t")
                drawing = ch.find(f"{{{W_NS}}}drawing")
                br = ch.find(f"{{{W_NS}}}br")
                if t_el is not None and t_el.text and drawing is None and br is None:
                    for ci, char in enumerate(t_el.text):
                        char_map.append((idx, ci))
                        full_text.append(char)
                else:
                    char_map.append((idx, -1))
                    full_text.append("\x00")
            elif tag in ("commentRangeStart", "commentRangeEnd", "bookmarkStart", "bookmarkEnd", "pPr"):
                continue
            else:
                char_map.append((idx, -1))
                full_text.append("\x00")

        joined = "".join(full_text)
        pos = joined.find(target)
        if pos == -1:
            raise ValueError(f"Target phrase not found in direct runs!\n  Target: {target!r}\n  Joined: {joined!r}")

        end_pos = pos + len(target) - 1
        start_child_idx, start_char_idx = char_map[pos]
        end_child_idx, end_char_idx = char_map[end_pos]

        first_r = children[start_child_idx]
        last_r = children[end_child_idx]
        first_r_pr = first_r.find(f"{{{W_NS}}}rPr")
        last_r_pr = last_r.find(f"{{{W_NS}}}rPr")

        first_txt = first_r.find(f"{{{W_NS}}}t").text or ""
        last_txt = last_r.find(f"{{{W_NS}}}t").text or ""

        prefix_txt = first_txt[:start_char_idx]
        suffix_txt = last_txt[end_char_idx + 1:]

        replacement_nodes: List[ET.Element] = []
        if prefix_txt:
            replacement_nodes.append(make_run_with_text(prefix_txt, first_r_pr, is_del=False))

        cid = None
        if summary:
            cid = self._alloc_cid()
            append_comment_xml(
                self.comments_root,
                cid=cid,
                category=category or "AUDIT",
                summary=summary,
                details=details or "",
                author=self.author,
                initials=self.initials,
                date_iso=self.date_iso,
                comment_prefix=self.comment_prefix,
            )
            self.applied_comments += 1
            replacement_nodes.append(ET.Element(f"{{{W_NS}}}commentRangeStart", {f"{{{W_NS}}}id": str(cid)}))

        preserved_markers: List[ET.Element] = []
        del_runs: List[Tuple[str, Optional[ET.Element]]] = []
        for c_i in range(start_child_idx, end_child_idx + 1):
            ch = children[c_i]
            tag = ch.tag.split("}")[-1]
            if tag != "r":
                preserved_markers.append(ch)
            else:
                r_pr = ch.find(f"{{{W_NS}}}rPr")
                t_txt = ch.find(f"{{{W_NS}}}t").text or ""
                if c_i == start_child_idx and c_i == end_child_idx:
                    sub_txt = t_txt[start_char_idx : end_char_idx + 1]
                elif c_i == start_child_idx:
                    sub_txt = t_txt[start_char_idx:]
                elif c_i == end_child_idx:
                    sub_txt = t_txt[: end_char_idx + 1]
                else:
                    sub_txt = t_txt
                if sub_txt:
                    del_runs.append((sub_txt, r_pr))

        for pm in preserved_markers:
            replacement_nodes.append(pm)

        if target != replacement:
            self.applied_edits += 1
            if target:
                del_el = ET.Element(f"{{{W_NS}}}del", {
                    f"{{{W_NS}}}id": str(self._alloc_wid()),
                    f"{{{W_NS}}}author": self.author,
                    f"{{{W_NS}}}date": self.date_iso,
                })
                for sub_txt, r_pr in del_runs:
                    del_el.append(make_run_with_text(sub_txt, r_pr, is_del=True))
                replacement_nodes.append(del_el)

            if replacement:
                ins_el = ET.Element(f"{{{W_NS}}}ins", {
                    f"{{{W_NS}}}id": str(self._alloc_wid()),
                    f"{{{W_NS}}}author": self.author,
                    f"{{{W_NS}}}date": self.date_iso,
                })
                ins_el.append(make_run_with_text(replacement, first_r_pr, is_del=False))
                replacement_nodes.append(ins_el)
        else:
            for sub_txt, r_pr in del_runs:
                replacement_nodes.append(make_run_with_text(sub_txt, r_pr, is_del=False))

        if cid is not None:
            replacement_nodes.append(ET.Element(f"{{{W_NS}}}commentRangeEnd", {f"{{{W_NS}}}id": str(cid)}))
            ref_r = ET.Element(f"{{{W_NS}}}r")
            ref_r_pr = ET.SubElement(ref_r, f"{{{W_NS}}}rPr")
            ET.SubElement(ref_r_pr, f"{{{W_NS}}}rStyle", {f"{{{W_NS}}}val": "CommentReference"})
            ET.SubElement(ref_r, f"{{{W_NS}}}commentReference", {f"{{{W_NS}}}id": str(cid)})
            replacement_nodes.append(ref_r)

        if suffix_txt:
            replacement_nodes.append(make_run_with_text(suffix_txt, last_r_pr, is_del=False))

        p_elem[start_child_idx : end_child_idx + 1] = replacement_nodes
        return True

    def attach_comment_to_whole_paragraph(
        self,
        p_elem: ET.Element,
        category: str,
        summary: str,
        details: str = "",
    ) -> int:
        """Anchors a sidebar comment across the entire paragraph without modifying its runs."""
        cid = self._alloc_cid()
        append_comment_xml(
            self.comments_root,
            cid=cid,
            category=category,
            summary=summary,
            details=details,
            author=self.author,
            initials=self.initials,
            date_iso=self.date_iso,
            comment_prefix=self.comment_prefix,
        )
        self.applied_comments += 1

        p_pr = p_elem.find(f"{{{W_NS}}}pPr")
        insert_idx = 1 if p_pr is not None else 0
        p_elem.insert(insert_idx, ET.Element(f"{{{W_NS}}}commentRangeStart", {f"{{{W_NS}}}id": str(cid)}))
        p_elem.append(ET.Element(f"{{{W_NS}}}commentRangeEnd", {f"{{{W_NS}}}id": str(cid)}))
        ref_r = ET.Element(f"{{{W_NS}}}r")
        ref_r_pr = ET.SubElement(ref_r, f"{{{W_NS}}}rPr")
        ET.SubElement(ref_r_pr, f"{{{W_NS}}}rStyle", {f"{{{W_NS}}}val": "CommentReference"})
        ET.SubElement(ref_r, f"{{{W_NS}}}commentReference", {f"{{{W_NS}}}id": str(cid)})
        p_elem.append(ref_r)
        return cid

    def mark_existing_ins_runs_deleted(self, p_elem: ET.Element, target_substr: str) -> None:
        """
        Preserves a prior coauthor's `<w:ins>` wrapper intact (matched by text substring so it is
        invariant to Google Docs export `w:id` renumbering) while wrapping its `<w:r><w:t>` runs
        inside `<w:del><w:r><w:delText>`.
        """
        found = False
        for ch in list(p_elem):
            if ch.tag == f"{{{W_NS}}}ins" and target_substr in "".join(ch.itertext()):
                new_ins_children = []
                for sub in list(ch):
                    if sub.tag == f"{{{W_NS}}}r":
                        t_el = sub.find(f"{{{W_NS}}}t")
                        r_pr = sub.find(f"{{{W_NS}}}rPr")
                        if t_el is not None and t_el.text:
                            del_wrap = ET.Element(f"{{{W_NS}}}del", {
                                f"{{{W_NS}}}id": str(self._alloc_wid()),
                                f"{{{W_NS}}}author": self.author,
                                f"{{{W_NS}}}date": self.date_iso,
                            })
                            del_wrap.append(make_run_with_text(t_el.text, r_pr, is_del=True))
                            new_ins_children.append(del_wrap)
                            self.applied_edits += 1
                            found = True
                            continue
                    new_ins_children.append(sub)
                ch[:] = new_ins_children
                break
        if not found:
            raise ValueError(f"Could not find <w:ins> containing {target_substr!r} in paragraph!")

    def split_soft_break_subheading(self, p_elem: ET.Element, heading_p_pr_template: ET.Element) -> ET.Element:
        """Splits a bold subheading merged with body text via `<w:br/>` into a standalone heading paragraph."""
        normalize_paragraph_runs(p_elem)
        children = list(p_elem)
        br_idx = None
        for idx, ch in enumerate(children):
            if ch.tag == f"{{{W_NS}}}r" and ch.find(f"{{{W_NS}}}br") is not None:
                br_idx = idx
                break
        if br_idx is None:
            raise ValueError("No <w:br/> found to split subheading!")

        head_p = ET.Element(f"{{{W_NS}}}p")
        head_p.append(copy.deepcopy(heading_p_pr_template))
        for ch in children[:br_idx]:
            if ch.tag != f"{{{W_NS}}}pPr":
                head_p.append(copy.deepcopy(ch))

        p_pr = p_elem.find(f"{{{W_NS}}}pPr")
        remaining = [c for c in children[br_idx + 1:] if c.tag != f"{{{W_NS}}}pPr"]
        p_elem[:] = ([p_pr] if p_pr is not None else []) + remaining

        body_idx = list(self.body).index(p_elem)
        self.body.insert(body_idx, head_p)
        return head_p

    def move_elements_inline_preserving_sectpr(
        self,
        inline_moves: Sequence[Tuple[ET.Element, Sequence[ET.Element]]],
        fallback_sectpr_target_p: Optional[ET.Element] = None,
    ) -> None:
        """
        Relocates figure/caption `<w:p>` elements inline right after target anchor paragraphs
        while ensuring any `<w:sectPr>` (Google Docs tab boundary) on a moved paragraph is
        transferred to `fallback_sectpr_target_p` rather than carried into the middle of Results.
        """
        for anchor_p, fig_elems in inline_moves:
            for fe in fig_elems:
                fe_p_pr = fe.find(f"{{{W_NS}}}pPr")
                if fe_p_pr is not None:
                    sect_pr = fe_p_pr.find(f"{{{W_NS}}}sectPr")
                    if sect_pr is not None and fallback_sectpr_target_p is not None:
                        fe_p_pr.remove(sect_pr)
                        dst_p_pr = fallback_sectpr_target_p.find(f"{{{W_NS}}}pPr")
                        if dst_p_pr is None:
                            dst_p_pr = ET.Element(f"{{{W_NS}}}pPr")
                            fallback_sectpr_target_p.insert(0, dst_p_pr)
                        dst_p_pr.append(sect_pr)
                self.body.remove(fe)

            anchor_idx = list(self.body).index(anchor_p)
            for offset, fe in enumerate(fig_elems):
                self.body.insert(anchor_idx + 1 + offset, fe)


class OpenXMLAuditSession:
    """
    Context manager that unpacks an input `.docx`, repairs transparent RGBA PNGs,
    exposes `DocxAuditEngine` + `body` elements, verifies 100% preservation of
    existing coauthor comments and hyperlinks, and writes a validated output `.docx`.
    """

    def __init__(
        self,
        input_docx: str,
        output_docx: str,
        author: str = DEFAULT_AUDIT_AUTHOR,
        initials: str = DEFAULT_AUDIT_INITIALS,
        comment_prefix: Optional[str] = "AI AUDIT",
        repair_rgba_figures: bool = True,
        normalize_body_limit: Optional[int] = None,
    ):
        self.input_docx = input_docx
        self.output_docx = output_docx
        self.author = author
        self.initials = initials
        self.comment_prefix = comment_prefix
        self.repair_rgba_figures = repair_rgba_figures
        self.normalize_body_limit = normalize_body_limit

        self.extract_dir: Optional[str] = None
        self.doc_tree: Optional[ET.ElementTree] = None
        self.comments_tree: Optional[ET.ElementTree] = None
        self.engine: Optional[DocxAuditEngine] = None
        self.repaired_figures: List[str] = []
        self._orig_hyperlink_count = 0

    def __enter__(self) -> "OpenXMLAuditSession":
        self.extract_dir = tempfile.mkdtemp(prefix="manuscript_audit_docx_")
        with zipfile.ZipFile(self.input_docx, "r") as z:
            z.extractall(self.extract_dir)

        if self.repair_rgba_figures:
            media_dir = os.path.join(self.extract_dir, "word", "media")
            self.repaired_figures = composite_rgba_pngs_in_dir(media_dir)

        doc_xml_path = os.path.join(self.extract_dir, "word", "document.xml")
        comments_xml_path = os.path.join(self.extract_dir, "word", "comments.xml")

        self.doc_tree = ET.parse(doc_xml_path)
        doc_root = self.doc_tree.getroot()
        self._orig_hyperlink_count = len(doc_root.findall(f".//{{{W_NS}}}hyperlink"))

        if os.path.exists(comments_xml_path):
            self.comments_tree = ET.parse(comments_xml_path)
            comments_root = self.comments_tree.getroot()
        else:
            comments_root = ET.Element(f"{{{W_NS}}}comments")
            self.comments_tree = ET.ElementTree(comments_root)

        self.engine = DocxAuditEngine(
            doc_root=doc_root,
            comments_root=comments_root,
            author=self.author,
            initials=self.initials,
            comment_prefix=self.comment_prefix,
            normalize_body_limit=self.normalize_body_limit,
        )
        return self

    def save(self) -> Dict[str, Any]:
        assert self.extract_dir and self.doc_tree and self.comments_tree and self.engine
        doc_xml_path = os.path.join(self.extract_dir, "word", "document.xml")
        comments_xml_path = os.path.join(self.extract_dir, "word", "comments.xml")

        self.doc_tree.write(doc_xml_path, xml_declaration=True, encoding="utf-8")

        has_any_comments = len(self.engine.comments_root.findall(f"{{{W_NS}}}comment")) > 0
        if has_any_comments:
            self.comments_tree.write(comments_xml_path, xml_declaration=True, encoding="utf-8")
            _ensure_comments_packaging(self.extract_dir)

        # Validate XML well-formedness, existing comment preservation, & hyperlink count
        val_doc = ET.fromstring(open(doc_xml_path, "rb").read())
        final_hyperlinks = len(val_doc.findall(f".//{{{W_NS}}}hyperlink"))
        if final_hyperlinks != self._orig_hyperlink_count:
            raise AssertionError(
                f"Hyperlink count changed ({self._orig_hyperlink_count} -> {final_hyperlinks})!"
            )

        final_cids = set()
        if os.path.exists(comments_xml_path):
            val_com = ET.fromstring(open(comments_xml_path, "rb").read())
            final_cids = {
                int(c.attrib[f"{{{W_NS}}}id"])
                for c in val_com.findall(f"{{{W_NS}}}comment")
                if c.attrib.get(f"{{{W_NS}}}id", "").isdigit()
            }
            start_anchors = {
                int(el.attrib[f"{{{W_NS}}}id"])
                for el in val_doc.findall(f".//{{{W_NS}}}commentRangeStart")
                if el.attrib.get(f"{{{W_NS}}}id", "").isdigit()
            }
            end_anchors = {
                int(el.attrib[f"{{{W_NS}}}id"])
                for el in val_doc.findall(f".//{{{W_NS}}}commentRangeEnd")
                if el.attrib.get(f"{{{W_NS}}}id", "").isdigit()
            }
            for orig_id in self.engine.existing_cids:
                if orig_id not in final_cids:
                    raise AssertionError(f"Missing original comment ID {orig_id} in comments.xml!")
                if orig_id not in start_anchors or orig_id not in end_anchors:
                    raise AssertionError(f"Missing anchor span for original comment ID {orig_id}!")

        if os.path.exists(self.output_docx):
            os.remove(self.output_docx)
        with zipfile.ZipFile(self.output_docx, "w", zipfile.ZIP_DEFLATED) as zout:
            for foldername, _, filenames in os.walk(self.extract_dir):
                for filename in filenames:
                    filepath = os.path.join(foldername, filename)
                    arcname = os.path.relpath(filepath, self.extract_dir)
                    zout.write(filepath, arcname)

        return {
            "output_docx": self.output_docx,
            "applied_track_changes": self.engine.applied_edits,
            "applied_comments": self.engine.applied_comments,
            "preserved_existing_comments": len(self.engine.existing_cids),
            "total_comments": len(final_cids),
            "preserved_hyperlinks": final_hyperlinks,
            "repaired_rgba_figures": self.repaired_figures,
        }

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        if exc_type is None and self.extract_dir and os.path.exists(self.extract_dir):
            try:
                if not os.path.exists(self.output_docx):
                    self.save()
            finally:
                shutil.rmtree(self.extract_dir, ignore_errors=True)
        elif self.extract_dir and os.path.exists(self.extract_dir):
            shutil.rmtree(self.extract_dir, ignore_errors=True)


def _ensure_comments_packaging(extract_dir: str) -> None:
    """Ensures `document.xml.rels` and `[Content_Types].xml` declare `word/comments.xml`."""
    doc_rels_path = os.path.join(extract_dir, "word", "_rels", "document.xml.rels")
    if os.path.exists(doc_rels_path):
        rels_tree = ET.parse(doc_rels_path)
        rels_root = rels_tree.getroot()
        has_comment_rel = any(
            rel.get("Type") == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
            for rel in rels_root
        )
        if not has_comment_rel:
            existing_rids = {rel.get("Id", "") for rel in rels_root}
            idx = 999
            while f"rIdComments{idx}" in existing_rids:
                idx += 1
            ET.SubElement(rels_root, f"{{{PKG_R_NS}}}Relationship", {
                "Id": f"rIdComments{idx}",
                "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
                "Target": "comments.xml",
            })
            rels_tree.write(doc_rels_path, xml_declaration=True, encoding="utf-8")

    ct_path = os.path.join(extract_dir, "[Content_Types].xml")
    if os.path.exists(ct_path):
        ct_tree = ET.parse(ct_path)
        ct_root = ct_tree.getroot()
        has_comment_ct = any(
            override.get("PartName") == "/word/comments.xml"
            for override in ct_root.findall(f"{{{CT_NS}}}Override")
        )
        if not has_comment_ct:
            ET.SubElement(ct_root, f"{{{CT_NS}}}Override", {
                "PartName": "/word/comments.xml",
                "ContentType": "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
            })
            ct_tree.write(ct_path, xml_declaration=True, encoding="utf-8")
