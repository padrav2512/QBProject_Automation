from __future__ import annotations

import importlib
import json
import os
from pathlib import Path

import pandas as pd
import streamlit as st

import cms_processor as _cms_processor


EXPECTED_PROCESSOR_BUILD_ID = "2026.09.25-authoring-rules-v12.1"
if getattr(_cms_processor, "PROCESSOR_BUILD_ID", None) != EXPECTED_PROCESSOR_BUILD_ID:
    _cms_processor = importlib.reload(_cms_processor)
if getattr(_cms_processor, "PROCESSOR_BUILD_ID", None) != EXPECTED_PROCESSOR_BUILD_ID:
    st.error("The app interface and document processor are temporarily out of sync. Reboot the Streamlit app before processing a document.")
    st.stop()

ProcessorOptions = _cms_processor.ProcessorOptions
process_docx_bytes = _cms_processor.process_docx_bytes


st.set_page_config(page_title="CMS DOCX Verification Processor", page_icon="✅", layout="wide")

st.title("CMS DOCX Verification Processor")
st.caption("Convert quality-checked Word question banks into tagged CMS verification documents, then download the document, images and audit report.")
st.caption(f"Processor build: {EXPECTED_PROCESSOR_BUILD_ID}")
st.caption("Clearly grouped mathematical expressions become Word equations. Version 12 also converts spaced slash fractions, inline radicals and rupee amounts, and applies conservative variable and geometry-label italics while protecting recognised units. Ambiguous cases remain for review.")

with st.sidebar:
    st.header("Processing mode")
    processing_mode_label = st.radio(
        "Choose what the app should change",
        ["Full CMS preparation", "Images and Alt Text only", "Images, Alt Text and safe Math formatting"],
        help="Use either image mode when the document already contains CMS tags and IDs. The safe Math option converts unambiguous mathematical forms and corrects a clearly mismatched FIB/MCQ Type tag; uncertain structures are reported.",
    )
    processing_mode = {
        "Full CMS preparation": "full",
        "Images and Alt Text only": "images_only",
        "Images, Alt Text and safe Math formatting": "images_math",
    }[processing_mode_label]

    st.header("CMS numbering")
    project_id = st.text_input(
        "Project ID",
        value="project",
        help="In Full CMS preparation this is used for generated question IDs. In Images and Alt Text only it is only a fallback when no project ID can be read from existing Question id tags.",
    )
    if processing_mode == "full":
        start_question = st.number_input("First question number", min_value=1, value=1, step=1)
        start_snippet_text = st.text_input("First snippet ID", value="", placeholder="Enter the first snippet ID")
        replace_ids = st.checkbox("Replace existing question and snippet IDs", value=True)
        st.header("Defaults for missing metadata")
        default_type = st.selectbox("Question type", ["FIB", "MCQ"], index=0)
        default_difficulty = st.selectbox("Difficulty", ["Easy", "Average", "Challenging"], index=1)
        default_objective = st.selectbox("Objective", ["Knowledge", "Comprehension", "Application", "Analysis"], index=2)
    else:
        start_question = 1
        start_snippet_text = ""
        replace_ids = False
        default_type = "FIB"
        default_difficulty = "Average"
        default_objective = "Application"

    if processing_mode == "images_only":
        st.info("Existing CMS text, tags, question IDs and snippet IDs will be preserved. The project ID in existing Question id tags takes priority over the fallback Project ID above.")
    elif processing_mode == "images_math":
        st.info("Existing question IDs and snippet IDs will be preserved. Images and Alt Text will be prepared, unambiguous mathematical forms will be converted to native Word equations, and a clearly mismatched FIB/MCQ Type tag will be corrected. Uncertain expressions or question structures will be reported for review.")

storage_root = Path(os.environ.get("CMS_STORAGE_DIR", Path(__file__).parent / "data")).resolve()
st.info("Upload a quality-checked DOCX, process it, and download all three outputs before closing the page.")

uploaded = st.file_uploader("Upload a Word document", type=["docx"], accept_multiple_files=False)

if processing_mode == "full":
    with st.expander("Input format — no CMS tags needed"):
        st.write("Use ordinary text in Word. Put each heading, option and answer on its own paragraph (Enter). The app adds CMS tags to the output automatically.")
        st.code("""Question 26: Average, Comprehension
Which congruence criterion applies?
a) ASA
b) SAS
c) SSS
d) RHS
Answer: c
Solution:
OA = OB, AM = BM, and OM is common.

Question 27: Easy, Knowledge
Type: FIB
What is 6 times 7?
Answer: 42
Solution:
6 times 7 equals 42.""", language=None)
        st.write("Question: and Choices: (or Options:) headings are optional for MCQs. Type: and Question type: are accepted. Use four options a) to d) and an Answer: or Correct answer: line before Solution:. MCQ keys may be b, b), (b), 2 or b) choice text. For multiple FIB answers, use an Answers: heading followed by one labelled answer per paragraph. A space after a slash fraction ends the fraction; use brackets for a multi-term denominator. Review the audit report before using the output.")

if uploaded is not None:
    size_mb = uploaded.size / (1024 * 1024)
    st.write(f"**Selected:** {uploaded.name} ({size_mb:.2f} MB)")
    max_mb = int(os.environ.get("CMS_MAX_UPLOAD_MB", "50"))
    if size_mb > max_mb:
        st.error(f"The file exceeds the configured {max_mb} MB upload limit.")
    elif st.button("Process document", type="primary", use_container_width=True):
        start_snippet = int(start_snippet_text.strip()) if start_snippet_text.strip().isdigit() and int(start_snippet_text.strip()) > 0 else None
        if processing_mode == "full" and start_snippet is None:
            st.error("Enter a valid positive First snippet ID before processing the document.")
            st.stop()
        options = ProcessorOptions(
            project_question_prefix=f"{project_id.strip().removesuffix('_q')}_q",
            start_question_number=int(start_question),
            start_snippet_id=start_snippet or 1,
            default_type=default_type,
            default_difficulty=default_difficulty,
            default_objective=default_objective,
            replace_existing_ids=replace_ids,
            processing_mode=processing_mode,
        )
        try:
            with st.spinner("Processing the Word document…"):
                result = process_docx_bytes(uploaded.getvalue(), uploaded.name, options, storage_root)
        except Exception as exc:
            st.exception(exc)
        else:
            if result.question_count:
                if processing_mode == "images_only":
                    st.success("Image filenames and Alt Text are ready. Download the updated DOCX, images ZIP and audit report below.")
                elif processing_mode == "images_math":
                    st.success("Image filenames, Alt Text and safe mathematical formatting are ready. Download the updated DOCX, images ZIP and audit report below.")
                else:
                    st.success("The document is ready. Download the processed DOCX, images ZIP and audit report below.")
            else:
                st.error("No question headings were detected, so CMS records and tags were not created. Review the verification report and correct the question-heading format before downloading a CMS-ready document.")
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Questions detected", result.question_count)
            col2.metric("Automatic fixes", result.fixed_count)
            col3.metric("Manual checks", result.manual_review_count)
            col4.metric("Image occurrences", result.image_count)

            if result.manual_review_count:
                st.warning(f"Action required: complete {result.manual_review_count} manual check(s) before final CMS upload. The app does not guess when a change could alter mathematical meaning.")
                st.caption("Verification queue status: PENDING_MANUAL_REVIEW")
            else:
                st.success("No unresolved manual-review findings were detected.")
                st.caption("Verification queue status: READY_FOR_VERIFICATION")

            report_data = json.loads(result.report_path.read_text(encoding="utf-8"))
            effective_project_id = report_data.get("effective_project_id") or project_id.strip() or "project"
            output_suffix = {
                "images_only": "Images_Alt_Text_Ready",
                "images_math": "Images_Alt_Text_Math_Ready",
                "full": "CMS_Verification_Ready",
            }[processing_mode]
            output_download_name = f"{Path(uploaded.name).stem}_{output_suffix}.docx"
            col_doc, col_images, col_report = st.columns(3)
            if result.question_count:
                col_doc.download_button(
                    "Download processed DOCX",
                    data=result.output_path.read_bytes(),
                    file_name=output_download_name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                    on_click="ignore",
                )
            else:
                col_doc.caption("Processed DOCX unavailable because no questions were detected.")
            col_images.download_button(
                "Download images ZIP",
                data=result.images_zip_path.read_bytes(),
                file_name=f"{effective_project_id}_images.zip",
                mime="application/zip",
                use_container_width=True,
                on_click="ignore",
            )
            col_report.download_button(
                "Download JSON audit report",
                data=json.dumps(report_data, ensure_ascii=False, indent=2).encode("utf-8"),
                file_name=f"{Path(uploaded.name).stem}_CMS_audit.json",
                mime="application/json",
                use_container_width=True,
                on_click="ignore",
            )

            manual_rows = [
                {
                    "Question": f.question or "Document",
                    "Paragraph": f.paragraph or "—",
                    "Action needed": f.message,
                }
                for f in result.findings
                if f.status == "manual_review"
            ]
            if manual_rows:
                st.subheader("Action required before CMS upload")
                st.caption("Work through this shorter list first. Each row is an unresolved check that needs a person to review the source or processed document.")
                st.dataframe(
                    pd.DataFrame(manual_rows),
                    use_container_width=True,
                    hide_index=True,
                    column_config={
                        "Question": st.column_config.TextColumn(width="small"),
                        "Paragraph": st.column_config.TextColumn(width="small"),
                        "Action needed": st.column_config.TextColumn(width="large"),
                    },
                )

            rows = [
                {
                    "Rule": f.rule,
                    "Status": f.status.replace("_", " ").title(),
                    "Question": f.question or "—",
                    "Paragraph": f.paragraph or "—",
                    "Finding": f.message,
                }
                for f in result.findings
            ]
            st.subheader("Verification report")
            st.caption("Complete record of passed checks, automatic fixes and manual-review findings.")
            st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

            st.caption(f"Processing reference: {result.run_id}")

st.divider()
with st.expander("What the app checks"):
    st.markdown(
        """
1. Removes bold only when an entire Question, Choices or Solution block is bold; selective bold emphasis is preserved.
2. Preserves native Word equations, converts unambiguous radical or simple-fraction assignments, and flags possible equation images or expressions with uncertain mathematical scope.
3. Preserves author-supplied italics in text and native equations. Authors decide the formatting of variables, geometry labels and units.
4. Removes completely blank equation objects, blank exponent/subscript templates and repeated spaces.
5. Normalises spaces around `=`, `+`, `−`, `×` and `÷`.
6. Adds a space after commas.
7. Converts slash-style fractions into native stacked Word equations and flags only cases that cannot be transformed safely.
8. Places labelled answers on separate lines when they are combined in one paragraph.
9. Normalises top-level subpart labels where the structure is unambiguous.
10. Checks and normalises Assertion–Reason wording and choice structure without guessing the correct answer.
11. Preserves table formatting. Authors must identify and bold any header rows or header columns in the source document.
12. Converts `₹50`, `₹ 50` and similar numeric amounts to `Rs 50`.
13. Italicises high-confidence variables and geometry labels while protecting articles, option labels and recognised measurement units; uncertain labels are reported.

The input does not need CMS tags. The app first identifies the question and section structure, then creates or repairs the confirmed CMS markers, sequential question IDs, sequential snippet IDs and required `@e@` delimiters. Metadata and section tag paragraphs are bold, matching the approved CMS document convention; choice markers, `@correct answer@` and `@e@` remain ordinary text. It validates the generated structure afterwards.
"""
    )
