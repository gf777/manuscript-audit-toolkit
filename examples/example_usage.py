"""
Example Script: How to use the Manuscript Audit Toolkit on any paper.
"""

from manuscript_audit import ManuscriptAuditor, OpenXMLCommentInjector

def run_example():
    # 1. Audit a manuscript text file
    # auditor = ManuscriptAuditor(model="gemini-2.5-pro")
    # results = auditor.run_full_audit("sample_manuscript.txt", output_json="sample_audit.json")

    # Sample mock audit results for demonstration
    sample_audit = [
        {
            "line_id": 1,
            "section": "RESULTS",
            "original_line": "Surpassing the Q60 benchmark with <1 error per 2 Mb.",
            "has_issue": True,
            "error_type": "MATH_ERROR",
            "error_explanation": "Q60 corresponds to 1 error per 1 Mb (10^-6), not 1 error per 2 Mb.",
            "corrected_line": "Surpassing the Q60 benchmark with <1 error per 1 Mb."
        }
    ]

    # 2. Inject into target docx
    # injector = OpenXMLCommentInjector("my_manuscript_draft.docx")
    # injector.inject_audit_comments(
    #     audit_results=sample_audit,
    #     output_docx_path="my_manuscript_audited.docx",
    #     pre_revision_docx_path="my_manuscript_pre_revision.docx"
    # )
    print("Example completed successfully.")

if __name__ == "__main__":
    run_example()
