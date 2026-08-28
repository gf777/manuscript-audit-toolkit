"""
Aligner Module
==============
Performs sentence-level semantic and fuzzy alignment between two versions of a
manuscript (e.g., Pre-AI / Pre-Revision Original vs. Revised Draft).

Features:
- Dual-level hierarchy: Paragraph-block mapping followed by intra-paragraph sentence pairing
- Global fallback search for moved or reorganized sections
- Identification of net-new additions versus modified text
"""

import re
import difflib
from typing import List, Dict, Any, Optional
import docx


def normalize_text(text: str) -> str:
    """Strips punctuation and whitespace for robust phonetic/character comparison."""
    return re.sub(r'[^a-zA-Z0-9]', '', text.lower())


class ManuscriptAligner:
    def __init__(self, pre_revision_docx: str, post_revision_docx: str):
        self.pre_docx_path = pre_revision_docx
        self.post_docx_path = post_revision_docx
        self.pre_paras = self._extract_paragraphs(pre_revision_docx)
        self.post_paras = self._extract_paragraphs(post_revision_docx)
        self.paragraph_map = self._build_paragraph_mapping()

    def _extract_paragraphs(self, docx_path: str) -> List[Dict[str, Any]]:
        doc = docx.Document(docx_path)
        paras = []
        for p_idx, p in enumerate(doc.paragraphs):
            txt = p.text.strip()
            if not txt:
                continue
            sents = [s.strip() for s in re.split(r'(?<=[.!?])\s+', txt) if len(s.strip()) > 5]
            paras.append({
                "p_idx": p_idx,
                "text": txt,
                "clean_prefix": normalize_text(txt[:80]),
                "sents": sents
            })
        return paras

    def _build_paragraph_mapping(self) -> Dict[int, Dict[str, Any]]:
        """Maps each post-revision paragraph to its most likely pre-revision origin."""
        p_map = {}
        for post_p in self.post_paras:
            best_match = None
            best_score = 0.0
            for pre_p in self.pre_paras:
                score = difflib.SequenceMatcher(None, post_p["clean_prefix"], pre_p["clean_prefix"]).ratio()
                if score > best_score:
                    best_score = score
                    best_match = pre_p
            if best_score > 0.35:
                p_map[post_p["p_idx"]] = best_match
        return p_map

    def find_pre_revision_sentence(self, post_p_idx: int, post_sentence: str) -> Optional[str]:
        """
        Finds the exact corresponding pre-revision sentence for a given sentence in the revised draft.
        Returns None if the sentence is a net-new insertion.
        """
        target_clean = normalize_text(post_sentence)
        if not target_clean:
            return None

        # 1. Look inside the mapped paragraph first (fast & context-accurate)
        if post_p_idx in self.paragraph_map:
            pre_p = self.paragraph_map[post_p_idx]
            best_s = None
            best_score = 0.0
            for s in pre_p["sents"]:
                score = difflib.SequenceMatcher(None, target_clean, normalize_text(s)).ratio()
                if score > best_score:
                    best_score = score
                    best_s = s
            if best_score > 0.35:
                return best_s
            if len(pre_p["sents"]) == 1 and len(pre_p["text"]) > 20:
                return pre_p["sents"][0]

        # 2. Global search fallback if paragraphs were relocated
        best_s = None
        best_score = 0.0
        for pre_p in self.pre_paras:
            for s in pre_p["sents"]:
                score = difflib.SequenceMatcher(None, target_clean, normalize_text(s)).ratio()
                if score > best_score:
                    best_score = score
                    best_s = s

        if best_score > 0.45:
            return best_s

        return None
