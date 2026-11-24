"""
Auditor Module
==============
Performs structured, line-by-line / sentence-by-sentence scientific, bioinformatic,
and editorial auditing of academic manuscripts (`.txt` or `.docx`).

Detects:
- Lexical & structural LLM prose artifacts ("AI slop": *delve, underscore, highlight,
  showcase, pivotal, crucial, intricate, multifaceted, tapestry, landscape, interplay,
  leverage, utilize*, throat-clearing starters, `-ing` participial tack-ons, copula
  avoidance, formulaic triplets)
- Ambiguous pronouns (*"This suggests..."*), hidden/shifting sample-size denominators
  (`n`), passive agency blurring present findings vs. cited literature, noun stacks
- Formatting & citation glitches (merged superscript references in software versions
  like `samtools v1.979`, unexpanded first-use acronyms)
- Factual, statistical, and nomenclature errors (Q-scores, P-values/FDR, italicization)
"""

import json
import os
import time
from typing import Any, Dict, List, Optional
import docx

try:
    from google import genai
    from google.genai import types
    HAVE_GEMINI = True
except ImportError:
    HAVE_GEMINI = False


AUDIT_SYSTEM_PROMPT = """You are an elite, rigorous scientific editor and computational genomics reviewer auditing an academic manuscript line-by-line.

Inspect every provided line/paragraph and flag any of the following:
1. LEXICAL & STRUCTURAL AI PROSE ("AI SLOP"):
   - Overused LLM vocabulary (*delve, underscore, highlight, showcase, testament to, pivotal, crucial, intricate, multifaceted, tapestry, realm, landscape, interplay, shed light on, pave the way, bolster, leverage, utilize, foster, overarching, comprehensive, robust, novel*).
   - Formulaic throat-clearing starters (*"Importantly, ...", "Notably, ...", "Interestingly, ...", "Crucially, ...", "It is worth noting that...", "Taken together, these findings underscore..."*).
   - Shallow `-ing` participial tack-ons appended after a factual clause (*"...differed between groups, highlighting the complex interplay of..."*).
   - Copula inflation (*"serves as / stands as / boasts / exhibits / features"* where *"is / are / has"* is cleaner) and artificial rhetorical triplets (*"Not only X, but also Y"*).
2. LACK OF CLARITY, AMBIGUITY, & SHIFTING DENOMINATORS:
   - Vague pronouns or dangling comparatives (*"This suggests..."*, *"was significantly higher"* without explicit antecedent, comparator, effect size, or test statistic).
   - Passive voice where the reader cannot tell whether an observation comes from the present study or prior literature.
   - Missing or shifting sample-size denominators (`n`) on subsets/percentages, or dense 4–6-word noun stacks.
3. BIOINFORMATIC, STATISTICAL, FORMATTING, & NOMENCLATURE ERRORS:
   - Merged superscript citation numbers in software versions or numbers (e.g., `MultiQC v1.993` or `samtools v1.979`).
   - Mathematical/statistical errors (e.g., Q-score formulas, mismatched P-values, collinearity artifacts), unexpanded first-use acronyms, or unitalicized gene/species names.

Return a valid JSON array of objects, one per input item, with exact keys:
- "line_id": integer identifier
- "section": section name (e.g., ABSTRACT, INTRODUCTION, RESULTS, DISCUSSION, METHODS)
- "original_line": exact input string (or exact target phrase to revise)
- "has_issue": boolean
- "error_type": "None" or category ("AI_PROSE", "CLARITY", "DENOMINATOR_CONSISTENCY", "MATH_STAT_ERROR", "CITATION_GLITCH", "NOMENCLATURE", "GRAMMAR")
- "error_explanation": concise, objective third-person rationale explaining what was changed and why
- "corrected_line": exact concrete replacement text
"""


class ManuscriptAuditor:
    def __init__(self, api_key: Optional[str] = None, model: str = "gemini-2.5-pro"):
        self.api_key = api_key or os.environ.get("GEMINI_API_KEY")
        self.model = model
        self.client = None
        if HAVE_GEMINI and self.api_key:
            self.client = genai.Client(api_key=self.api_key)

    def parse_text_lines(self, file_path: str) -> List[Dict[str, Any]]:
        """
        Parses either a plain-text (`.txt`) or Word (`.docx`) manuscript into
        structured paragraphs/lines tagged with section headers.
        """
        if file_path.lower().endswith(".docx"):
            doc = docx.Document(file_path)
            raw_lines = [p.text.strip() for p in doc.paragraphs]
        else:
            with open(file_path, "r", encoding="utf-8") as f:
                raw_lines = [line.strip() for line in f.readlines()]

        structured_lines = []
        current_section = "FRONT_MATTER"
        line_idx = 1

        section_keywords = [
            "ABSTRACT", "INTRODUCTION", "RESULTS", "DISCUSSION",
            "METHODS", "MATERIALS AND METHODS", "LIMITATIONS",
            "DECLARATIONS", "REFERENCES", "FIGURES", "EXTENDED DATA",
        ]

        for line in raw_lines:
            if not line:
                continue
            upper_line = line.strip().upper()
            if upper_line in section_keywords or any(
                upper_line.startswith(k) for k in section_keywords if len(line) < 45
            ):
                current_section = upper_line
                continue

            structured_lines.append({
                "line_id": line_idx,
                "section": current_section,
                "text": line,
            })
            line_idx += 1

        return structured_lines

    def audit_batch(self, lines_batch: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """Submits a batch of lines to the configured Gemini model for structured auditing."""
        if not self.client:
            raise RuntimeError("Gemini Client not initialized. Set GEMINI_API_KEY environment variable.")

        prompt_payload = json.dumps([
            {"line_id": l["line_id"], "section": l["section"], "text": l["text"]}
            for l in lines_batch
        ], indent=2)

        config = types.GenerateContentConfig(
            system_instruction=AUDIT_SYSTEM_PROMPT,
            temperature=0.1,
            response_mime_type="application/json",
        )

        response = self.client.models.generate_content(
            model=self.model,
            contents=f"Audit the following manuscript lines carefully:\n\n{prompt_payload}",
            config=config,
        )

        try:
            return json.loads(response.text)
        except Exception as exc:
            print(f"Error parsing LLM response JSON: {exc}")
            return [
                {
                    "line_id": l["line_id"],
                    "section": l["section"],
                    "original_line": l["text"],
                    "has_issue": False,
                    "error_type": "None",
                    "error_explanation": "",
                    "corrected_line": l["text"],
                }
                for l in lines_batch
            ]

    def run_full_audit(
        self,
        input_file: str,
        batch_size: int = 15,
        output_json: str = "audit_results.json",
    ) -> List[Dict[str, Any]]:
        """Audits an entire `.txt` or `.docx` manuscript in batches."""
        lines = self.parse_text_lines(input_file)
        print(f"Loaded {len(lines)} structured lines from {input_file}")

        all_results: List[Dict[str, Any]] = []
        batches = [lines[i : i + batch_size] for i in range(0, len(lines), batch_size)]

        for i, batch in enumerate(batches):
            print(f"Auditing Batch {i + 1}/{len(batches)} (Lines {batch[0]['line_id']} - {batch[-1]['line_id']})...")
            try:
                all_results.extend(self.audit_batch(batch))
            except Exception as exc:
                print(f"Failed batch {i + 1}: {exc}")
                time.sleep(2)
            time.sleep(0.5)

        with open(output_json, "w", encoding="utf-8") as f:
            json.dump(all_results, f, indent=2)

        flagged_count = sum(1 for r in all_results if r.get("has_issue"))
        print(f"Audit completed! Total lines: {len(all_results)}, Flagged issues: {flagged_count}")
        return all_results
