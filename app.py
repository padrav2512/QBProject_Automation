from __future__ import annotations

import importlib
import json
import os
from pathlib import Path

import pandas as pd
import streamlit as st

import cms_processor as _cms_processor


EXPECTED_PROCESSOR_BUILD_ID = "2026.09.25-processing-sections-v13"
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
st.caption("Choose the processing sections you need. Mathematical changes remain conservative: clear expressions are converted and ambiguous cases are preserved for review.")

with st.sidebar:
    st.header("Processing sections")
    with st.container(border=True):
        add_cms_tags = st.checkbox("1. Add or repair CMS tags", value=True)
        st.caption("Creates or repairs question records, metadata, IDs, snippet IDs, Choices/Answers/Solution sections and required `@e@` markers.")
    with st.container(border=True):
        process_images = st.checkbox("2. Handle images and Alt Text", value=True)
        st.caption("Extracts images, assigns CMS filenames, reuses filenames for exact duplicates, writes Alt Text and creates the images ZIP.")
    with st.container(border=True):
        format_math_structure = st.checkbox("3. Check Math format and document structure", value=True)
        st.caption("Applies safe Word equations, fractions, radicals, variable italics, spacing, rupee conversion and other deterministic formatting checks. Ambiguous cases are only reported.")

    if not any((add_cms_tags, process_images, format_math_structure)):
        st.warning("Select at least one processing section.")

    if add_cms_tags or process_images:
        st.header("CMS and image naming")
        project_id = st.text_input(
            "Project ID",
            value="project",
            help="Used for generated question IDs and image filenames. When tag insertion is off, a project ID already present in the document takes priority.",
        )
    else:
        project_id = "project"

    if add_cms_tags:
        st.subheader("Question and snippet numbering")
        start_question = st.number_input("First question number", min_value=1, value=1, step=1)
        start_snippet_text = st.text_input("First snippet ID", value="", placeholder="Enter the first snippet ID")
        replace_ids = st.checkbox("Replace existing question and snippet IDs", value=True)
        st.subheader("Defaults for missing metadata")
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

    if not add_cms_tags and process_images:
        st.info("Existing CMS tags and IDs will be preserved. A project ID found in existing Question id tags takes priority over the Project ID entered above.")
    if not format_math_structure:
        st.info("Math and document-formatting changes are off. Existing equations, mathematical text, italics, bold text and spacing will be preserved.")

storage_root = Path(os.environ.get("CMS_STORAGE_DIR", Path(__file__).parent / "data")).resolve()
st.info("Upload a quality-checked DOCX, process it, and download all three outputs before closing the page.")

uploaded = st.file_uploader("Upload a Word document", type=["docx"], accept_multiple_files=False)

if add_cms_tags:
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
        if not any((add_cms_tags, process_images, format_math_structure)):
            st.error("Select at least one processing section before processing the document.")
            st.stop()
        start_snippet = int(start_snippet_text.strip()) if start_snippet_text.strip().isdigit() and int(start_snippet_text.strip()) > 0 else None
        if add_cms_tags and start_snippet is None:
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
            processing_mode="custom",
            add_cms_tags=add_cms_tags,
            process_images=process_images,
            format_math_structure=format_math_structure,
        )
        try:
            with st.spinner("Processing the Word document…"):
                result = process_docx_bytes(uploaded.getvalue(), uploaded.name, options, storage_root)
        except Exception as exc:
            st.exception(exc)
        else:
            if result.question_count:
                completed_sections = []
                if add_cms_tags:
                    completed_sections.append("CMS tags")
                if process_images:
                    completed_sections.append("images and Alt Text")
                if format_math_structure:
                    completed_sections.append("Math and document formatting")
                st.success(f"Processing is complete for: {', '.join(completed_sections)}. Download the available outputs below.")
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
                (True, True, True): "CMS_Verification_Ready",
                (True, False, False): "CMS_Tags_Ready",
                (False, True, False): "Images_Alt_Text_Ready",
                (False, False, True): "Math_Structure_Ready",
                (True, True, False): "CMS_Tags_Images_Ready",
                (True, False, True): "CMS_Tags_Math_Ready",
                (False, True, True): "Images_Alt_Text_Math_Ready",
            }[(add_cms_tags, process_images, format_math_structure)]
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
            if process_images:
                col_images.download_button(
                    "Download images ZIP",
                    data=result.images_zip_path.read_bytes(),
                    file_name=f"{effective_project_id}_images.zip",
                    mime="application/zip",
                    use_container_width=True,
                    on_click="ignore",
                )
            else:
                col_images.caption("Images ZIP not created because image handling was not selected.")
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
**1. Add or repair CMS tags**

- Identifies plain or already-tagged question records.
- Creates or repairs CMS metadata, sequential question IDs, sequential snippet IDs, section markers and required `@e@` delimiters.
- Builds Choices or Answers blocks from unambiguous author input and corrects a clearly mismatched MCQ/FIB Type tag.
- Applies bold formatting to CMS metadata and section tags only.

**2. Handle images and Alt Text**

- Extracts question, choice and solution images into GIF files.
- Applies the agreed CMS filename and Word Alt Text conventions.
- Reuses the first filename when the same image occurs more than once in the same CMS image area.
- Flags unassigned images and images in an FIB Answers block for manual review.

**3. Check Math format and document structure**

- Preserves existing native Word equations and converts unambiguous fractions, radicals, exponents and equations into native Word equations.
- Places visible spacing correctly at prose-to-equation boundaries and preserves surrounding text.
- Italicises high-confidence variables and geometry labels while protecting articles, option labels and recognised units.
- Removes blank equation or exponent templates, repairs deterministic spacing, and converts numeric rupee amounts to `Rs`.
- Removes bold only when an entire Question, Choices or Solution block is bold; selective emphasis is preserved.
- Normalises unambiguous answer lines, subpart labels and Assertion–Reason structure, while preserving table formatting.
- Leaves expressions whose mathematical meaning is uncertain unchanged and lists them under Action required.

Only the selected sections change the document. The verification report records every applied fix and every unresolved manual check.
"""
    )
