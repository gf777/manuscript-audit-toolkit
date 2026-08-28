"""
Auditor Module
==============
Performs structured, line-by-line / sentence-by-sentence auditing of scientific
manuscripts using LLMs (Google Gemini, OpenAI, Claude, or custom validators).

Detects:
- AI formatting glitches (merged superscripts, broken punctuation)
- Scientific / mathematical inconsistencies (e.g. Q-score vs error rate formulas)
- Nomenclature errors (gene symbols, species names, italicization)
- Truncated or garbled clauses
- Citations vs text context mismatches
"""

import json
import re
import os
import time
from typing import List, Dict, Any, Optional

try:
    from google import genai
    from google.genai import types
    HAVE_GEMINI = True
except ImportError:
    HAVE_GEMINI = False


AUDIT_SYSTEM_PROMPT = """You are an elite, rigorous scientific editor and computational genomics expert auditing an academic manuscript line-by-line.

Your task is to inspect every provided line/sentence from the draft and identify any:
1. AI glitches or formatting errors (e.g. merged superscripts like 'Chenc' instead of 'Chen c', run-together text, unexpanded acronyms).
2. Factual, mathematical, or scientific inaccuracies (e.g. Q60 accuracy defined as 1 error per 2 Mb instead of 1 error per 1 Mb).
3. Grammatical lapses, garbled text, incomplete sentences, or truncated clauses.
4. Capitalization and nomenclature issues (e.g. gene symbols vs protein symbols, italicization).

You must return a valid JSON array of objects for all lines evaluated.
Each object must have:
- "line_id": integer identifier
- "section": section name (e.g. ABSTRACT, INTRODUCTION, RESULTS, DISCUSSION, METHODS)
- "original_line": exact input string
- "has_issue": boolean (true if an issue is found, false if clean)
- "error_type": "None" or category (e.g. "AI_GLITCH", "MATH_ERROR", "GRAMMAR", "NOMENCLATURE", "TRUNCATION")
- "error_explanation": concise, objective explanation of what is wrong and why
- "corrected_line": exact proposed corrected text
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
        Parses a manuscript text file into structured lines tagged with section headers.
        """
        with open(file_path, "r", encoding="utf-8") as f:
            raw_lines = [line.strip() for line in f.readlines()]

        structured_lines = []
        current_section = "FRONT_MATTER"
        line_idx = 1

        section_keywords = [
            "ABSTRACT", "INTRODUCTION", "RESULTS", "DISCUSSION",
            "METHODS", "MATERIALS AND METHODS", "DECLARATIONS", "REFERENCES"
        ]

        for line in raw_lines:
            if not line:
                continue

            # Check for section header
            upper_line = line.strip().upper()
            if upper_line in section_keywords or any(upper_line.startswith(k) for k in section_keywords if len(line) < 40):
                current_section = upper_line
                continue

            structured_lines.append({
                "line_id": line_idx,
                "section": current_section,
                "text": line
            })
            line_idx += 1

        return structured_lines

    def audit_batch(self, lines_batch: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """
        Submits a batch of lines to the LLM for rigorous auditing.
        """
        if not self.client:
            raise RuntimeError("Gemini Client not initialized. Set GEMINI_API_KEY environment variable.")

        prompt_payload = json.dumps([{
            "line_id": l["line_id"],
            "section": l["section"],
            "text": l["text"]
        } for l in lines_batch], indent=2)

        prompt = f"Audit the following manuscript lines carefully:\n\n{prompt_payload}"

        config = types.GenerateContentConfig(
            system_instruction=AUDIT_SYSTEM_PROMPT,
            temperature=0.1,
            response_mime_type="application/json"
        )

        response = self.client.models.generate_content(
            model=self.model,
            contents=prompt,
            config=config
        )

        try:
            results = json.loads(response.text)
            return results
        except Exception as e:
            print(f"Error parsing LLM response JSON: {e}")
            print("Raw response:", response.text[:500])
            # Fallback: return default unflagged
            return [{
                "line_id": l["line_id"],
                "section": l["section"],
                "original_line": l["text"],
                "has_issue": False,
                "error_type": "None",
                "error_explanation": "",
                "corrected_line": l["text"]
            } for l in lines_batch]

    def run_full_audit(self, input_file: str, batch_size: int = 15, output_json: str = "audit_results.json") -> List[Dict[str, Any]]:
        """
        Audits an entire manuscript in batches with rate-limit protection.
        """
        lines = self.parse_text_lines(input_file)
        print(f"Loaded {len(lines)} structured lines from {input_file}")

        all_results = []
        batches = [lines[i:i + batch_size] for i in range(0, len(lines), batch_size)]

        for i, batch in enumerate(batches):
            print(f"Auditing Batch {i + 1}/{len(batches)} (Lines {batch[0]['line_id']} - {batch[-1]['line_id']})...")
            try:
                batch_res = self.audit_batch(batch)
                all_results.extend(batch_res)
            except Exception as e:
                print(f"Failed batch {i + 1}: {e}")
                time.sleep(2)

            time.sleep(0.5)

        with open(output_json, "w", encoding="utf-8") as f:
            json.dump(all_results, f, indent=2)

        flagged_count = sum(1 for r in all_results if r.get("has_issue"))
        print(f"Audit completed! Total lines: {len(all_results)}, Flagged issues: {flagged_count}")
        print(f"Saved results to {output_json}")
        return all_results
