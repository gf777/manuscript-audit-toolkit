"""
Command Line Interface (CLI)
============================
Unified CLI for Google Docs snapshot downloads, prior-comment inventories,
spreadsheet & RGBA figure forensics, LLM auditing, and non-destructive OpenXML
Track Changes + phrase-anchored comment injection.
"""

import argparse
import json
from .auditor import ManuscriptAuditor
from .commenter import OpenXMLCommentInjector
from .docx_engine import inspect_docx_comments_and_changes
from .figures import audit_and_repair_docx_figures
from .google_docs import detect_google_doc_tabs, download_google_snapshot
from .spreadsheets import SpreadsheetAuditor


def main():
    parser = argparse.ArgumentParser(
        prog="manuscript-audit",
        description="Manuscript Audit Toolkit v2.0: Scientific manuscript forensics, OpenXML Track Changes & comments, and multi-tab Google Docs sync.",
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    # 1. Inspect existing .docx comments, track changes, hyperlinks, & tab sections
    insp_p = subparsers.add_parser("inspect", help="Inventory existing coauthor comments, Track Changes, hyperlinks, and tabs in a .docx.")
    insp_p.add_argument("--doc", "-d", required=True, help="Path to .docx file.")
    insp_p.add_argument("--output", "-o", default=None, help="Optional JSON output path.")

    # 2. Download live Google Doc / Google Sheet snapshot (& detect multiple tabs)
    dl_p = subparsers.add_parser("download", help="Download a live link-shared Google Doc (.docx) or Sheet (.xlsx) snapshot.")
    dl_p.add_argument("--url", "-u", required=True, help="Google Doc/Sheet URL or File ID.")
    dl_p.add_argument("--output", "-o", required=True, help="Output .docx or .xlsx path.")
    dl_p.add_argument("--single-tab", action="store_true", help="Export only the ?tab= target tab instead of all tabs.")
    dl_p.add_argument("--check-tabs", action="store_true", help="Detect and report multi-tab Google Doc structure.")

    # 3. Audit command
    audit_p = subparsers.add_parser("audit", help="Run LLM-powered line-by-line audit on a .txt or .docx manuscript.")
    audit_p.add_argument("--input", "-i", required=True, help="Path to manuscript (.txt or .docx).")
    audit_p.add_argument("--output", "-o", default="audit_results.json", help="Output JSON path.")
    audit_p.add_argument("--batch-size", "-b", type=int, default=15, help="Batch size for LLM evaluation.")
    audit_p.add_argument("--model", "-m", default="gemini-2.5-pro", help="LLM model identifier.")

    # 4. Inject command
    inject_p = subparsers.add_parser("inject", help="Inject native Track Changes & phrase-anchored comments into a .docx.")
    inject_p.add_argument("--doc", "-d", required=True, help="Target .docx file.")
    inject_p.add_argument("--audit-results", "-a", required=True, help="Path to audit_results.json.")
    inject_p.add_argument("--pre-revision", "-p", default=None, help="Optional pre-revision .docx for alignment.")
    inject_p.add_argument("--output", "-o", default="audited_manuscript.docx", help="Output .docx path.")
    inject_p.add_argument("--comments-only", action="store_true", help="Disable <w:del>/<w:ins> Track Changes and emit comments only.")
    inject_p.add_argument("--author", default="Manuscript Audit (AI)", help="XML w:author attribute for new edits/comments.")
    inject_p.add_argument("--prefix", default="AI AUDIT", help="Bold prefix inside comment headers to distinguish from coauthor comments.")

    # 5. Pipeline command (end-to-end)
    pipe_p = subparsers.add_parser("pipeline", help="Run end-to-end audit and OpenXML Track Changes + comment injection.")
    pipe_p.add_argument("--input-txt", "-i", required=True, help="Draft manuscript (.txt or .docx).")
    pipe_p.add_argument("--doc", "-d", required=True, help="Target draft .docx to inject into.")
    pipe_p.add_argument("--pre-revision", "-p", default=None, help="Optional pre-revision .docx.")
    pipe_p.add_argument("--output-docx", "-o", default="audited_final.docx", help="Output .docx path.")
    pipe_p.add_argument("--output-json", "-j", default="audit_results.json", help="Audit results JSON path.")
    pipe_p.add_argument("--model", "-m", default="gemini-2.5-pro", help="LLM model identifier.")

    # 6. Spreadsheet forensics command
    xlsx_p = subparsers.add_parser("audit-xlsx", help="Scan & repair Supplementary Tables (.xlsx) for Excel date/coord/URL corruption.")
    xlsx_p.add_argument("--input", "-i", required=True, help="Input .xlsx path.")
    xlsx_p.add_argument("--output", "-o", default=None, help="Optional repaired output .xlsx path.")
    xlsx_p.add_argument("--report", "-r", default=None, help="Optional JSON report path.")

    # 7. Visual figure RGBA audit & dark-mode repair command
    fig_p = subparsers.add_parser("fix-figures", help="Audit embedded .docx figures and composite transparent RGBA PNGs onto white RGB.")
    fig_p.add_argument("--doc", "-d", required=True, help="Input .docx path.")
    fig_p.add_argument("--output", "-o", default=None, help="Optional repaired output .docx path.")

    args = parser.parse_args()

    if args.command == "inspect":
        info = inspect_docx_comments_and_changes(args.doc)
        if args.output:
            with open(args.output, "w", encoding="utf-8") as f:
                json.dump(info, f, indent=2)
        print(json.dumps({k: v for k, v in info.items() if k not in ("comments", "track_changes")}, indent=2))

    elif args.command == "download":
        res = download_google_snapshot(args.url, args.output, single_tab_only=args.single_tab)
        print(json.dumps(res, indent=2))
        if args.check_tabs and args.output.lower().endswith(".docx"):
            tab_report = detect_google_doc_tabs(args.url)
            print("Tab Structure Check:", json.dumps(tab_report, indent=2))

    elif args.command == "audit":
        auditor = ManuscriptAuditor(model=args.model)
        auditor.run_full_audit(args.input, batch_size=args.batch_size, output_json=args.output)

    elif args.command == "inject":
        with open(args.audit_results, "r", encoding="utf-8") as f:
            results = json.load(f)
        injector = OpenXMLCommentInjector(args.doc, author=args.author, comment_prefix=args.prefix or None)
        injector.inject_audit_comments(
            audit_results=results,
            output_docx_path=args.output,
            pre_revision_docx_path=args.pre_revision,
            apply_track_changes=not args.comments_only,
        )

    elif args.command == "pipeline":
        auditor = ManuscriptAuditor(model=args.model)
        results = auditor.run_full_audit(args.input_txt, output_json=args.output_json)
        injector = OpenXMLCommentInjector(args.doc)
        injector.inject_audit_comments(
            audit_results=results,
            output_docx_path=args.output_docx,
            pre_revision_docx_path=args.pre_revision,
        )

    elif args.command == "audit-xlsx":
        sa = SpreadsheetAuditor(args.input)
        rep = sa.audit_and_repair(output_xlsx_path=args.output)
        if args.report:
            with open(args.report, "w", encoding="utf-8") as f:
                json.dump(rep, f, indent=2)
        print(json.dumps({k: v for k, v in rep.items() if k != "issues"}, indent=2))

    elif args.command == "fix-figures":
        rep = audit_and_repair_docx_figures(args.doc, output_docx=args.output)
        print(json.dumps(rep, indent=2))


if __name__ == "__main__":
    main()
