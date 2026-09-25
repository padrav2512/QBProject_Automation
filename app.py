from __future__ import annotations

import importlib
import json
import os
from pathlib import Path

import pandas as pd
import streamlit as st

import cms_processor as _cms_processor
from heymath_cms_mapper import HeyMathCmsMapper


EXPECTED_PROCESSOR_BUILD_ID = "2026.09.25-mapping-workflow-v14"
if getattr(_cms_processor, "PROCESSOR_BUILD_ID", None) != EXPECTED_PROCESSOR_BUILD_ID:
    _cms_processor = importlib.reload(_cms_processor)
if getattr(_cms_processor, "PROCESSOR_BUILD_ID", None) != EXPECTED_PROCESSOR_BUILD_ID:
    st.error("The app interface and document processor are temporarily out of sync. Reboot the Streamlit app before processing a document.")
    st.stop()

ProcessorOptions = _cms_processor.ProcessorOptions
process_docx_bytes = _cms_processor.process_docx_bytes


@st.cache_resource
def get_cms_mapper(base_url: str, login: str, password: str) -> HeyMathCmsMapper:
    return HeyMathCmsMapper(base_url, login, password)


st.set_page_config(page_title="CMS DOCX Verification Processor", page_icon="✅", layout="wide")

st.title("CMS DOCX Verification Processor")
st.caption("Convert quality-checked Word question banks into tagged CMS verification documents, then download the document, images and audit report.")
st.caption(f"Processor build: {EXPECTED_PROCESSOR_BUILD_ID}")
st.caption("Choose the processing sections you need. Mathematical changes remain conservative: clear expressions are converted and ambiguous cases are preserved for review.")

with st.sidebar:
    st.header("Processing sections")
    with st.container(border=True):
        add_cms_tags = st.checkbox("1. CMS tags and question structure", value=True)
        with st.expander("What this covers · example"):
            st.write("Creates or repairs question records, metadata, IDs, snippet IDs, Choices/Answers/Solution sections and required `@e@` markers.")
            st.code("Question 1: Easy, Knowledge\n...\nAnswer: b", language=None)
    with st.container(border=True):
        process_images = st.checkbox("2. Handle images and Alt Text", value=True)
        with st.expander("What this covers · example"):
            st.write("Extracts images, assigns CMS filenames, reuses filenames for exact duplicates, writes Word Alt Text and creates the images ZIP.")
            st.code("project10436_q8_1.gif", language=None)
    with st.container(border=True):
        format_math = st.checkbox("3. Apply safe mathematical formatting", value=True)
        with st.expander("What this covers · example"):
            st.write("Creates native Word equations for unambiguous fractions, radicals, exponents and equations; applies high-confidence variable italics and reports uncertain expressions.")
            st.code("x = √144  →  native Word equation\n1/4  →  stacked fraction", language=None)
    with st.container(border=True):
        normalize_text_structure = st.checkbox("4. Normalize text and document formatting", value=True)
        with st.expander("What this covers · example"):
            st.write("Repairs deterministic spacing, converts numeric rupee amounts, removes whole-block bold, and normalises answer lines, subparts and Assertion–Reason structure.")
            st.code("₹ 1,250  →  Rs 1,250\nword  ?  →  word?", language=None)
    with st.container(border=True):
        prepare_mappings = st.checkbox("5. Prepare curriculum and taxonomy mapping data", value=True)
        with st.expander("What this covers · example"):
            st.write("Collects per-question mapping paths, normalises spacing around `>>`, removes these author-only lines from the CMS-ready DOCX and creates a mapping CSV for use after Anand’s upload.")
            st.code(
                "Curriculum: CBSE NCERT >> Class 4 >> Mathematics >> Measuring Length >> Metres and Centimetres\n"
                "India >> Class 4 >> Mathematics >> Measurement\n"
                "Taxonomy: Mathematics >> Measurement >> Length",
                language=None,
            )

    if not any((add_cms_tags, process_images, format_math, normalize_text_structure, prepare_mappings)):
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
    if not format_math:
        st.info("Mathematical-formatting changes are off. Existing equations, mathematical text and mathematical italics will be preserved.")
    if not normalize_text_structure:
        st.info("Text and document-formatting changes are off. Existing prose spacing, bold text, subparts and answer layout will be preserved.")

storage_root = Path(os.environ.get("CMS_STORAGE_DIR", Path(__file__).parent / "data")).resolve()
st.info("Upload a quality-checked DOCX, process it, and download the available outputs before closing the page.")

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

if prepare_mappings:
    with st.expander("Input format for curriculum and taxonomy mappings"):
        st.write("Place mapping instructions inside each question. Start a block with `Curriculum:` or `Taxonomy:`. Additional paths may follow on separate paragraphs and must contain `>>`.")
        st.code(
            "Curriculum: CBSE NCERT >> Class 4 >> Mathematics >> Measuring Length >> Metres and Centimetres\n"
            "India >> Class 4 >> Mathematics >> Measurement\n"
            "India >> Class 4 >> Mathematics >> Case Study Based Questions\n\n"
            "Taxonomy: Mathematics >> Measurement >> Length\n"
            "Mathematics >> Word Problems >> Measurement",
            language=None,
        )
        st.caption("These lines are authoring metadata. After they are captured in the mapping CSV, they are removed from the CMS-ready Word document.")

if uploaded is not None:
    size_mb = uploaded.size / (1024 * 1024)
    st.write(f"**Selected:** {uploaded.name} ({size_mb:.2f} MB)")
    max_mb = int(os.environ.get("CMS_MAX_UPLOAD_MB", "50"))
    if size_mb > max_mb:
        st.error(f"The file exceeds the configured {max_mb} MB upload limit.")
    elif st.button("Process document", type="primary", use_container_width=True):
        if not any((add_cms_tags, process_images, format_math, normalize_text_structure, prepare_mappings)):
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
            format_math=format_math,
            normalize_text_structure=normalize_text_structure,
            prepare_mappings=prepare_mappings,
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
                if format_math:
                    completed_sections.append("safe Math formatting")
                if normalize_text_structure:
                    completed_sections.append("text and document formatting")
                if prepare_mappings:
                    completed_sections.append("mapping data")
                st.success(f"Processing is complete for: {', '.join(completed_sections)}. Download the available outputs below.")
            else:
                st.error("No question headings were detected, so CMS records and tags were not created. Review the verification report and correct the question-heading format before downloading a CMS-ready document.")
            col1, col2, col3, col4, col5 = st.columns(5)
            col1.metric("Questions detected", result.question_count)
            col2.metric("Automatic fixes", result.fixed_count)
            col3.metric("Manual checks", result.manual_review_count)
            col4.metric("Image occurrences", result.image_count)
            col5.metric("Mappings prepared", result.mapping_count)

            if result.manual_review_count:
                st.warning(f"Action required: complete {result.manual_review_count} manual check(s) before final CMS upload. The app does not guess when a change could alter mathematical meaning.")
                st.caption("Verification queue status: PENDING_MANUAL_REVIEW")
            else:
                st.success("No unresolved manual-review findings were detected.")
                st.caption("Verification queue status: READY_FOR_VERIFICATION")

            report_data = json.loads(result.report_path.read_text(encoding="utf-8"))
            effective_project_id = report_data.get("effective_project_id") or project_id.strip() or "project"
            output_download_name = report_data.get("download_filename") or f"{Path(uploaded.name).stem}_Processed.docx"
            col_doc, col_images, col_mapping, col_report = st.columns(4)
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
            if prepare_mappings:
                col_mapping.download_button(
                    "Download mapping CSV",
                    data=result.mapping_csv_path.read_bytes(),
                    file_name=f"{Path(uploaded.name).stem}_mapping_data.csv",
                    mime="text/csv",
                    use_container_width=True,
                    on_click="ignore",
                )
            else:
                col_mapping.caption("Mapping CSV not created because mapping preparation was not selected.")
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
with st.expander("Apply mappings to CMS after Anand’s upload", expanded=False):
    st.write("Use this only after Anand’s tool has created the questions in CMS. Upload the mapping CSV generated above, review it, and then apply the mappings.")
    st.info("Existing Curriculum and Taxonomy mappings are retained. This tool only adds missing mappings; it never removes or replaces an existing mapping.")
    mapping_upload = st.file_uploader("Upload mapping CSV", type=["csv"], key="mapping_csv_upload")
    mapping_frame = None
    mapping_problem = None
    if mapping_upload is not None:
        try:
            mapping_frame = pd.read_csv(mapping_upload, dtype=str).fillna("")
            mapping_frame.columns = [str(column).strip() for column in mapping_frame.columns]
        except Exception as exc:
            mapping_problem = f"The mapping CSV could not be read: {exc}"
        else:
            required_columns = {"Question", "Snippet ID", "Mapping Type", "Path"}
            missing_columns = required_columns - set(mapping_frame.columns)
            if missing_columns:
                mapping_problem = "The mapping CSV is missing: " + ", ".join(sorted(missing_columns))
            else:
                mapping_frame = mapping_frame[["Question", "Snippet ID", "Mapping Type", "Path"]].copy()
                mapping_frame["Mapping Type"] = mapping_frame["Mapping Type"].str.strip().str.title()
                mapping_frame["Path"] = mapping_frame["Path"].str.strip()
                invalid_types = mapping_frame[~mapping_frame["Mapping Type"].isin(["Curriculum", "Taxonomy"])]
                invalid_snippets = mapping_frame[~mapping_frame["Snippet ID"].str.fullmatch(r"\d+")]
                empty_paths = mapping_frame[mapping_frame["Path"] == ""]
                if not invalid_types.empty:
                    mapping_problem = "Every Mapping Type must be Curriculum or Taxonomy."
                elif not invalid_snippets.empty:
                    mapping_problem = "Every mapping row needs a numeric Snippet ID."
                elif not empty_paths.empty:
                    mapping_problem = "Every mapping row needs a path."
                else:
                    st.dataframe(mapping_frame, use_container_width=True, hide_index=True)
                    st.caption(f"{len(mapping_frame)} mapping path(s) across {mapping_frame['Snippet ID'].nunique()} snippet(s).")

    if mapping_problem:
        st.error(mapping_problem)

    try:
        cms_config = dict(st.secrets["heymath_cms"])
        missing_secret = [key for key in ("base_url", "login", "password") if not cms_config.get(key)]
    except Exception:
        cms_config = {}
        missing_secret = ["base_url", "login", "password"]

    if missing_secret:
        st.caption("CMS mapping is not yet configured on this server. Add the shared CMS account under `[heymath_cms]` in Streamlit secrets.")

    can_apply = mapping_frame is not None and mapping_problem is None and not missing_secret and not mapping_frame.empty
    if st.button("Validate and apply mappings to CMS", type="primary", disabled=not can_apply, use_container_width=True):
        mapper = get_cms_mapper(cms_config["base_url"], cms_config["login"], cms_config["password"])
        validation_errors: list[str] = []
        with st.spinner("Checking CMS snippets and mapping paths before saving anything…"):
            for snippet_text in mapping_frame["Snippet ID"].drop_duplicates():
                if mapper.get_mappings(int(snippet_text)) is None:
                    validation_errors.append(f"Snippet {snippet_text} does not yet exist in CMS.")
            for _, row in mapping_frame.drop_duplicates(["Mapping Type", "Path"]).iterrows():
                error = mapper.validate_path(row["Path"], taxonomy=row["Mapping Type"] == "Taxonomy")
                if error:
                    validation_errors.append(f"{row['Mapping Type']}: {error}")

        if validation_errors:
            st.error("Nothing was mapped because the pre-check found a problem.")
            st.dataframe(pd.DataFrame({"Problem": validation_errors}), use_container_width=True, hide_index=True)
        else:
            result_rows = []
            grouped = list(mapping_frame.groupby("Snippet ID", sort=False))
            progress = st.progress(0.0)
            for index, (snippet_text, rows_for_snippet) in enumerate(grouped, 1):
                curriculum_paths = rows_for_snippet.loc[rows_for_snippet["Mapping Type"] == "Curriculum", "Path"].tolist()
                taxonomy_paths = rows_for_snippet.loc[rows_for_snippet["Mapping Type"] == "Taxonomy", "Path"].tolist()
                try:
                    mapping_result = mapper.map_question(int(snippet_text), curriculum_paths, taxonomy_paths)
                    status = "Done" if mapping_result.success else "Problem"
                    details = "; ".join(
                        mapping_result.errors
                        if mapping_result.errors
                        else mapping_result.added + mapping_result.already_present
                    )
                    result_rows.append(
                        {
                            "Snippet ID": snippet_text,
                            "Question": mapping_result.question_name or "",
                            "Status": status,
                            "Added": len(mapping_result.added),
                            "Already present": len(mapping_result.already_present),
                            "Details": details or "No change required",
                        }
                    )
                except Exception as exc:
                    result_rows.append(
                        {
                            "Snippet ID": snippet_text,
                            "Question": "",
                            "Status": "Problem",
                            "Added": 0,
                            "Already present": 0,
                            "Details": str(exc),
                        }
                    )
                progress.progress(index / len(grouped))
            st.session_state["cms_mapping_results"] = result_rows

    if st.session_state.get("cms_mapping_results"):
        mapping_results = pd.DataFrame(st.session_state["cms_mapping_results"])
        st.subheader("CMS mapping results")
        st.dataframe(mapping_results, use_container_width=True, hide_index=True)
        if (mapping_results["Status"] == "Problem").any():
            st.warning("Some snippets need attention. Review the Details column before retrying those snippets.")
        else:
            st.success("All requested mappings were applied or were already present. Existing mappings were retained.")
        st.download_button(
            "Download mapping results CSV",
            data=mapping_results.to_csv(index=False).encode("utf-8-sig"),
            file_name="CMS_mapping_results.csv",
            mime="text/csv",
            on_click="ignore",
        )

st.divider()
with st.expander("What the app checks"):
    st.markdown(
        """
**1. CMS tags and question structure**

- Identifies plain or already-tagged question records.
- Creates or repairs CMS metadata, sequential question IDs, sequential snippet IDs, section markers and required `@e@` delimiters.
- Builds Choices or Answers blocks from unambiguous author input and corrects a clearly mismatched MCQ/FIB Type tag.
- Applies bold formatting to CMS metadata and section tags only.

**2. Handle images and Alt Text**

- Extracts question, choice and solution images into GIF files.
- Applies the agreed CMS filename and Word Alt Text conventions.
- Reuses the first filename when the same image occurs more than once in the same CMS image area.
- Flags unassigned images and images in an FIB Answers block for manual review.

**3. Apply safe mathematical formatting**

- Preserves existing native Word equations and converts unambiguous fractions, radicals, exponents and equations into native Word equations.
- Places visible spacing correctly at prose-to-equation boundaries and preserves surrounding text.
- Italicises high-confidence variables and geometry labels while protecting articles, option labels and recognised units.
- Removes blank equation or exponent templates.
- Leaves expressions whose mathematical meaning is uncertain unchanged and lists them under Action required.

**4. Normalize text and document formatting**

- Repairs deterministic prose, punctuation and operator spacing.
- Converts numeric rupee amounts such as `₹ 1,250` to `Rs 1,250`.
- Removes bold only when an entire Question, Choices or Solution block is bold; selective emphasis is preserved.
- Normalises unambiguous answer lines, subpart labels and Assertion–Reason structure, while preserving table formatting.

**5. Prepare curriculum and taxonomy mapping data**

- Reads per-question Curriculum and Taxonomy path blocks and normalises spacing around `>>`.
- Connects every path to its question number and New snippet ID.
- Removes the author-only mapping lines from the processed DOCX and creates a reusable mapping CSV.
- Supports a later CMS mapping step after Anand’s upload has created the snippets.
- Adds missing mappings only. Existing CMS mappings are retained and are never removed or replaced.

Only the selected sections change the document. The verification report records every applied fix and every unresolved manual check.
"""
    )
