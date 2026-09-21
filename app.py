from __future__ import annotations

import json
import os
from pathlib import Path

import pandas as pd
import streamlit as st

from cms_processor import ProcessorOptions, process_docx_bytes


st.set_page_config(page_title="CMS DOCX Verification Processor", page_icon="✅", layout="wide")

st.title("CMS DOCX Verification Processor")
st.caption("Convert quality-checked Word question banks into tagged CMS verification documents, then download the document, images and audit report.")

with st.sidebar:
    st.header("Processing mode")
    processing_mode_label = st.radio(
        "Choose what the app should change",
        ["Full CMS preparation", "Images and Alt Text only"],
        help="Use Images and Alt Text only when the document already contains correct CMS tags and you want its text, tags, IDs and formatting preserved.",
    )
    processing_mode = "images_only" if processing_mode_label == "Images and Alt Text only" else "full"

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
    else:
        start_question = 1
        start_snippet_text = ""
        replace_ids = False

    st.header("Defaults for missing metadata")
    default_type = st.selectbox("Question type", ["FIB", "MCQ"], index=0)
    default_difficulty = st.selectbox("Difficulty", ["Easy", "Average", "Challenging"], index=1)
    default_objective = st.selectbox("Objective", ["Knowledge", "Comprehension", "Application", "Analysis"], index=2)

    st.header("Document structure")
    bold_first_column = st.checkbox("Treat first table column as a header", value=False)
    convert_romans = st.checkbox("Convert top-level (i), (ii)… to (a), (b)…", value=True)

    if processing_mode == "images_only":
        st.info("Existing CMS text, tags, question IDs and snippet IDs will be preserved. The project ID in existing Question id tags takes priority over the fallback Project ID above.")

storage_root = Path(os.environ.get("CMS_STORAGE_DIR", Path(__file__).parent / "data")).resolve()
st.info("Upload a quality-checked DOCX, process it, and download all three outputs before closing the page.")

uploaded = st.file_uploader("Upload a Word document", type=["docx"], accept_multiple_files=False)

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
            bold_first_table_column=bold_first_column,
            convert_top_level_roman_subparts=convert_romans,
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
                st.warning("Complete the listed manual checks before final CMS upload. The app does not guess when a change could alter mathematical meaning.")
                st.caption("Verification queue status: PENDING_MANUAL_REVIEW")
            else:
                st.success("No unresolved manual-review findings were detected.")
                st.caption("Verification queue status: READY_FOR_VERIFICATION")

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
            st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

            report_data = json.loads(result.report_path.read_text(encoding="utf-8"))
            effective_project_id = report_data.get("effective_project_id") or project_id.strip() or "project"
            output_suffix = "Images_Alt_Text_Ready" if processing_mode == "images_only" else "CMS_Verification_Ready"
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

            st.caption(f"Processing reference: {result.run_id}")

st.divider()
with st.expander("What the app checks"):
    st.markdown(
        """
1. Removes bold only when an entire Question or Solution block is bold; selective bold emphasis is preserved.
2. Detects native Word equations and flags possible equation images or ordinary-text equations.
3. Keeps only mathematical variables in italics across Questions, Answers, Choices and Solutions, including ordinary text and native equations.
4. Removes completely blank equation objects, blank exponent/subscript templates and repeated spaces.
5. Normalises spaces around `=`, `+`, `−`, `×` and `÷`.
6. Adds a space after commas.
7. Converts slash-style fractions into native stacked Word equations and flags only cases that cannot be transformed safely.
8. Places labelled answers on separate lines when they are combined in one paragraph.
9. Normalises top-level subpart labels where the structure is unambiguous.
10. Checks and normalises Assertion–Reason wording and choice structure without guessing the correct answer.
11. Bolds table header rows and, optionally, the first column.

The input does not need CMS tags. The app first identifies the question and section structure, then creates or repairs the confirmed CMS markers, sequential question IDs, sequential snippet IDs and required `@e@` delimiters. Metadata and section tag paragraphs are bold, matching the approved CMS document convention; choice markers, `@correct answer@` and `@e@` remain ordinary text. It validates the generated structure afterwards.
"""
    )
