"""
Command Line Interface (CLI)
============================
Unified command line tool for running audits, aligning revisions, and injecting comments.
"""

import argparse
import sys
import json
from .auditor import ManuscriptAuditor
from .commenter import OpenXMLCommentInjector


def main():
    parser = argparse.ArgumentParser(
        prog="manuscript-audit",
        description="Manuscript Audit Toolkit: Automated scientific line audit & OpenXML comment injection."
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    # 1. Audit command
    audit_p = subparsers.add_parser("audit", help="Run LLM-powered line-by-line audit on manuscript text.")
    audit_p.add_argument("--input", "-i", required=True, help="Path to manuscript text file (.txt).")
    audit_p.add_argument("--output", "-o", default="audit_results.json", help="Output JSON path.")
    audit_p.add_argument("--batch-size", "-b", type=int, default=15, help="Batch size for LLM evaluation.")
    audit_p.add_argument("--model", "-m", default="gemini-2.5-pro", help="LLM model identifier.")

    # 2. Inject command
    inject_p = subparsers.add_parser("inject", help="Inject audit comments into .docx file.")
    inject_p.add_argument("--doc", "-d", required=True, help="Target draft .docx to inject comments into.")
    inject_p.add_argument("--audit-results", "-a", required=True, help="Path to audit_results.json.")
    inject_p.add_argument("--pre-revision", "-p", default=None, help="Optional pre-revision .docx for comparison.")
    inject_p.add_argument("--output", "-o", default="audited_manuscript.docx", help="Output .docx path.")

    # 3. Pipeline command (all-in-one)
    pipe_p = subparsers.add_parser("pipeline", help="Run end-to-end audit and comment injection.")
    pipe_p.add_argument("--input-txt", "-i", required=True, help="Draft manuscript text file (.txt).")
    pipe_p.add_argument("--doc", "-d", required=True, help="Target draft .docx to inject comments into.")
    pipe_p.add_argument("--pre-revision", "-p", default=None, help="Pre-revision .docx file.")
    pipe_p.add_argument("--output-docx", "-o", default="audited_final.docx", help="Output .docx file.")
    pipe_p.add_argument("--output-json", "-j", default="audit_results.json", help="Audit results JSON path.")

    args = parser.parse_args()

    if args.command == "audit":
        auditor = ManuscriptAuditor(model=args.model)
        auditor.run_full_audit(args.input, batch_size=args.batch_size, output_json=args.output)

    elif args.command == "inject":
        with open(args.audit_results, "r", encoding="utf-8") as f:
            results = json.load(f)
        injector = OpenXMLCommentInjector(args.doc)
        injector.inject_audit_comments(
            audit_results=results,
            output_docx_path=args.output,
            pre_revision_docx_path=args.pre_revision
        )

    elif args.command == "pipeline":
        auditor = ManuscriptAuditor()
        results = auditor.run_full_audit(args.input_txt, output_json=args.output_json)
        injector = OpenXMLCommentInjector(args.doc)
        injector.inject_audit_comments(
            audit_results=results,
            output_docx_path=args.output_docx,
            pre_revision_docx_path=args.pre_revision
        )


if __name__ == "__main__":
    main()
