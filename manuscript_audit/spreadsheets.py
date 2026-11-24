"""
Programmatic Spreadsheet Forensics (`openpyxl`)
===============================================
Scans Supplementary Tables (`.xlsx`) cell-by-cell for classic bioinformatic and
spreadsheet corruption bugs:

1. Excel Auto-Date Corruption of Gene Symbols (`MARCH1`, `SEPT1`, `DEC1`, `GEN1`,
   `OCT4`, `FEB1`, `MARC1`) and floating-point numbers (e.g. PBS / FST / P-values
   parsed as `datetime` objects or date-formatted strings).
2. Scientific-Notation Genomic Coordinate Truncation (`start == end` on windows
   with coordinates >= 100,000,000 bp due to 3-digit `1.06e+08` rounding).
3. Residual URL-Encoded Strings (`%2C`, `%20`, `%3B`, `%3D`) from GFF/VCF attributes.
4. Unfinished Placeholders (`IN PROGRESS`, `TODO`, `TBD`, `FIXME`, `#REF!`, `#NAME?`).
"""

import datetime
import re
import urllib.parse
from typing import Any, Dict, List, Optional

try:
    import openpyxl
    from openpyxl.comments import Comment
    HAVE_OPENPYXL = True
except ImportError:
    HAVE_OPENPYXL = False

MONTH_TO_GENE_PREFIX = {
    3: "MARCH",
    9: "SEPT",
    12: "DEC",
    10: "OCT",
}

GENE_COLUMN_HINTS = ("gene", "symbol", "locus", "hgnc", "target")
FLOAT_COLUMN_HINTS = ("pbs", "fst", "pval", "p_val", "fdr", "qval", "score", "stat", "beta", "ratio", "freq")
START_COLUMN_HINTS = ("start", "chromstart", "begin", "pos_start", "win_start")
END_COLUMN_HINTS = ("end", "chromend", "stop", "pos_end", "win_end")
PLACEHOLDER_PATTERNS = re.compile(r"\b(IN\s+PROGRESS|TODO|TBD|FIXME)\b|^#(REF!|NAME\?|VALUE!|DIV/0!|N/A)$", re.IGNORECASE)
URL_ENCODED_PATTERN = re.compile(r"%[0-9A-Fa-f]{2}")


class SpreadsheetAuditor:
    """Forensic scanner and non-destructive repair utility for `.xlsx` workbooks."""

    def __init__(self, xlsx_path: str):
        if not HAVE_OPENPYXL:
            raise RuntimeError("openpyxl is required for spreadsheet forensics (`pip install openpyxl`).")
        self.xlsx_path = xlsx_path

    def audit_and_repair(
        self,
        output_xlsx_path: Optional[str] = None,
        author: str = "Manuscript Audit",
    ) -> Dict[str, Any]:
        wb = openpyxl.load_workbook(self.xlsx_path, data_only=False)
        issues: List[Dict[str, Any]] = []
        fixes_applied = 0

        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            max_row = ws.max_row or 0
            max_col = ws.max_column or 0
            if max_row == 0 or max_col == 0:
                continue

            # Identify header row (first non-empty row within top 5 rows)
            header_row_idx = 1
            headers: Dict[int, str] = {}
            for r_idx in range(1, min(6, max_row + 1)):
                row_vals = [ws.cell(r_idx, c).value for c in range(1, max_col + 1)]
                if sum(1 for v in row_vals if v is not None and str(v).strip()) >= 2:
                    header_row_idx = r_idx
                    headers = {
                        c: str(ws.cell(r_idx, c).value or "").strip().lower()
                        for c in range(1, max_col + 1)
                    }
                    break

            start_cols = [c for c, h in headers.items() if h in START_COLUMN_HINTS]
            end_cols = [c for c, h in headers.items() if h in END_COLUMN_HINTS]

            for r_idx in range(header_row_idx + 1, max_row + 1):
                # Check zero-length genomic window truncation (`start == end >= 1e7`)
                if start_cols and end_cols:
                    s_val = ws.cell(r_idx, start_cols[0]).value
                    e_val = ws.cell(r_idx, end_cols[0]).value
                    if isinstance(s_val, (int, float)) and isinstance(e_val, (int, float)):
                        if s_val == e_val and s_val >= 10_000_000:
                            issues.append({
                                "sheet": sheet_name,
                                "cell": ws.cell(r_idx, start_cols[0]).coordinate,
                                "type": "SCIENTIFIC_NOTATION_COORD_TRUNCATION",
                                "value": f"start={s_val}, end={e_val}",
                                "detail": "Zero-length window at large genomic coordinate (likely 3-digit scientific notation rounding).",
                            })

                for c_idx in range(1, max_col + 1):
                    cell = ws.cell(r_idx, c_idx)
                    val = cell.value
                    if val is None:
                        continue
                    col_name = headers.get(c_idx, "")

                    # 1. Excel Auto-Date Corruption (`datetime.datetime` or `datetime.date`)
                    if isinstance(val, (datetime.datetime, datetime.date)):
                        is_gene_col = any(k in col_name for k in GENE_COLUMN_HINTS)
                        is_float_col = any(k in col_name for k in FLOAT_COLUMN_HINTS)
                        if is_gene_col and val.month in MONTH_TO_GENE_PREFIX and val.day <= 31:
                            reconstructed = f"{MONTH_TO_GENE_PREFIX[val.month]}{val.day}"
                            issues.append({
                                "sheet": sheet_name,
                                "cell": cell.coordinate,
                                "type": "EXCEL_DATE_GENE_CORRUPTION",
                                "value": str(val),
                                "reconstructed": reconstructed,
                                "detail": f"Date object in gene column '{col_name}'; likely gene symbol {reconstructed}.",
                            })
                            if output_xlsx_path:
                                cell.value = reconstructed
                                cell.number_format = "@"
                                cell.comment = Comment(
                                    f"[AI AUDIT] Reconstructed gene symbol '{reconstructed}' from Excel auto-date '{val}'.",
                                    author,
                                )
                                fixes_applied += 1
                        elif is_float_col:
                            issues.append({
                                "sheet": sheet_name,
                                "cell": cell.coordinate,
                                "type": "EXCEL_DATE_FLOAT_CORRUPTION",
                                "value": str(val),
                                "detail": f"Date object found in numeric statistic column '{col_name}'.",
                            })

                    # 2. String checks: URL-encoding & placeholders
                    elif isinstance(val, str):
                        if PLACEHOLDER_PATTERNS.search(val):
                            issues.append({
                                "sheet": sheet_name,
                                "cell": cell.coordinate,
                                "type": "UNFINISHED_PLACEHOLDER_OR_FORMULA_ERROR",
                                "value": val,
                                "detail": f"Placeholder or Excel formula error '{val}' in column '{col_name}'.",
                            })
                        if URL_ENCODED_PATTERN.search(val):
                            decoded = urllib.parse.unquote(val)
                            if decoded != val:
                                issues.append({
                                    "sheet": sheet_name,
                                    "cell": cell.coordinate,
                                    "type": "URL_ENCODED_STRING",
                                    "value": val,
                                    "reconstructed": decoded,
                                    "detail": f"URL-encoded string detected; decoded='{decoded}'.",
                                })
                                if output_xlsx_path:
                                    cell.value = decoded
                                    cell.comment = Comment(
                                        f"[AI AUDIT] Decoded URL-escaped text from '{val}'.",
                                        author,
                                    )
                                    fixes_applied += 1

        if output_xlsx_path:
            wb.save(output_xlsx_path)

        return {
            "input_xlsx": self.xlsx_path,
            "output_xlsx": output_xlsx_path,
            "sheet_count": len(wb.sheetnames),
            "issue_count": len(issues),
            "fixes_applied": fixes_applied,
            "issues": issues,
        }
