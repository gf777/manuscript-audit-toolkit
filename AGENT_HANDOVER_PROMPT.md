# End-to-End Scientific Manuscript Audit, AI-Prose De-Slopping, Track Changes, & In-Place Multi-Tab Google Docs Upload — Agent Handover Protocol

Execute a publication-grade scientific, bioinformatic, visual, statistical, and editorial audit of a manuscript package hosted on Google Docs and Google Sheets, and upload revisions directly back into the original live Google Docs and Google Sheets links as native Suggesting-mode edits (`<w:del>`/`<w:ins>`) and phrase-anchored sidebar comments **without merging or losing any Google Docs tabs**.

---

## Step 1: Download Live Snapshots, Detect Multi-Tab Structure, & Inventory Prior Comments

1. Extract the File Identifiers (IDs) and Tab IDs (`?tab=t.xxxxxx`) from the user-provided URLs using `manuscript_audit.extract_google_id_and_tab`.
2. Download BOTH the full multi-tab `.docx` export (`export?format=docx`) and, when `?tab=` is present, the single-tab export (`export?format=docx&tab=<TAB_ID>`) via `manuscript_audit.download_google_snapshot` / `detect_google_doc_tabs`:
   - **CRITICAL MULTI-TAB CHECK**: Compare `body_element_count` and `<w:sectPr>` section breaks between the full export and single-tab export. If the live Google Doc has multiple tabs (e.g., Main Manuscript + Notes + Extended Data + Outlines), build the single-tab `*_AUDITED.docx` from the target Main Manuscript tab and deploy it using `google_apps_script/update_google_docs_preserve_tabs.gs` (`updateGoogleDocPreservingAllTabs()`), which updates the Main Manuscript tab in place AND recreates/populates all other tabs unmerged via Google Docs REST API v1 (`AddDocumentTabRequest` + `UpdateDocumentTabPropertiesRequest`).
3. Run `manuscript_audit.inspect_docx_comments_and_changes` on the downloaded `.docx`:
   - Catalog 100% of existing coauthor comments (`w:id`, `w:author`, `w:date`, anchored text, comment body) and existing Track Changes (`<w:ins>`, `<w:del>`).
   - **CRITICAL PRESERVATION RULE**: Preserve 100% of existing coauthor comments and their exact `<w:commentRangeStart>` / `<w:commentRangeEnd>` anchors. Start new comment IDs strictly at `max(existing_comment_ids) + 1`.
   - When marking temporary inline coauthor question-prompts inside `<w:ins>` as deleted, match `<w:ins>` by text substring (`engine.mark_existing_ins_runs_deleted(p_elem, substr)`) rather than volatile export `w:id` integers.

---

## Step 2: Deep Scientific, Bioinformatic, Visual, & Spreadsheet Forensics

1. **Cross-Document Consistency Audit**:
   - Cross-check every sample size (`n`), filtering step, statistical test, P-value, False Discovery Rate (FDR) threshold, regression slope, and gene/species symbol across Title, Abstract, Introduction, Results, Discussion, Methods, Figure captions, Extended Data, and Supplementary Tables.
   - Verify that exclusions or thresholds stated in Methods match the actual figure panels and open-source repository code.
2. **Programmatic Spreadsheet Forensics (`manuscript_audit.SpreadsheetAuditor`)**:
   - Scan every cell of every `.xlsx` sheet for Excel auto-date corruption of gene symbols (`MARCH1`, `SEPT1`, `DEC1`, `GEN1`) and floating-point statistics, 3-digit scientific-notation genomic coordinate truncation (`start == end` at `>= 1e8` bp), URL-encoded strings (`%2C`), and `"IN PROGRESS"` / `"TODO"` placeholders. Save `*_AUDITED_FIXED.xlsx`.
3. **Visual Figure Audit & Dark-Mode RGBA Repair (`manuscript_audit.audit_and_repair_docx_figures`)**:
   - Visually inspect every image in `word/media/` against its caption and Results claims (checking panel letters `a, b, c...`, raw variable names on axes, label collisions, and desynchronized captions from earlier figure drafts).
   - Composite any RGBA PNG with transparent pixels (`alpha < 255`) onto a solid white RGB `(255, 255, 255)` background so black axes and labels remain legible in dark mode.
4. **Methods & Bibliography Verification**:
   - Check software versions/flags for merged superscript citation numbers (`MultiQC v1.993`, `samtools v1.979`).
   - Verify all GitHub/external URLs, duplicate preprint-vs-published references, consortium author parsing glitches (`Consortium, H. P. R.`), and missing references.

---

## Step 3: Mandatory Audit for AI-Generated Prose ("LLM Slop") & Lack of Clarity

Apply every prose and clarity fix directly via `DocxAuditEngine.apply_phrase_edit` paired with a concise third-person rationale:

1. **Eliminate Lexical & Structural AI Tells**:
   - Replace *delve, underscore, highlight, showcase, testament to, pivotal, crucial, intricate, multifaceted, tapestry, realm, landscape, interplay, shed light on, pave the way, leverage, utilize, foster, overarching, comprehensive, robust, novel*, throat-clearing starters (*"Importantly, ..."*, *"Notably, ..."*), shallow `-ing` participial tack-ons, copula inflation (*"serves as / boasts / exhibits"*), and artificial triplets.
2. **Fix Ambiguity, Hidden Denominators, & Passive Agency**:
   - Replace ambiguous *"This suggests..."* / *"Both..."* starters with explicit subjects, distinguish the authors' present findings from cited literature, state explicit sample-size denominators (`n`) for every subset, and unpack dense noun stacks.
3. **Distinguish AI Audit Comments Clearly from Prior Coauthor Comments**:
   - Because Google Docs `.docx` conversion attributes imported comment cards to the Google account performing the upload, use **third-person voice** when referencing coauthors (*"resolving Giulio Formenti's Comment 14"*) AND prefix every new comment header in bold with **`[AI AUDIT | <CATEGORY>]`** (`comment_prefix="AI AUDIT"`), while stripping any internal zero-based XML array indices (`P113`).

---

## Step 4: Comparative Target Journal Assessment & Editorial Gatekeeper Strategy

1. Build a Comparative Journal Matrix (Impact Factor, Open Access APC vs. $0 hybrid option, scope fit, review speed, consortium precedents).
2. Identify **The Under-Sold Asset** (upgrading Title, Abstract, and Discussion synthesis in Track Changes) vs. **The Over-Exposed Vulnerability** (reframing/diagnosing vulnerable statistical sub-analyses in the text and providing a concrete 1–2-day bioinformatic bridge analysis in the comments).
3. Recommend a 3-step top-down submission cascade.

---

## Step 5: Non-Destructive OpenXML Injection (`OpenXMLAuditSession` / `DocxAuditEngine`)

1. Use `OpenXMLAuditSession` so all edits split `<w:r>` runs cleanly at exact character offsets while cloning `<w:rPr>` formatting and treating `<w:hyperlink>`, `<w:ins>`, `<w:del>`, and `<m:oMath>` as inviolable barriers (`preserved_hyperlinks == original_hyperlinks`).
2. When moving end-of-manuscript figures inline into `Results`, use `engine.move_elements_inline_preserving_sectpr` so any `<w:sectPr>` (Google Docs tab boundary) on the final figure caption stays anchored at the end of Tab 1 rather than introducing a section break inside `Results`.

---

## Step 6: In-Place Multi-Tab Google Docs & Sheets Upload (`update_google_docs_preserve_tabs.gs`)

Provide the user with `google_apps_script/update_google_docs_preserve_tabs.gs` pre-configured with their target File ID, Tab ID, and staged `.docx` filename:
1. Upload `<manuscript>_AUDITED.docx` to Google Drive (kept as `.docx`).
2. In the live Google Doc (**Extensions → Apps Script**), paste `update_google_docs_preserve_tabs.gs`, enable **Drive API** and **Google Docs API** under **Services (+)**, and run **`updateGoogleDocPreservingAllTabs()`**.
3. Verify that Tab 1 displays all native green Suggesting-mode edits and `[AI AUDIT | ...]` anchored comments alongside all preserved coauthor comments, while Tabs 2..N remain separate, unmerged, and intact in the left sidebar.
