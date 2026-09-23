from __future__ import annotations

import io
import json
import os
import re
import shutil
import uuid
from copy import deepcopy
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Iterable, Iterator

from docx import Document
from docx.document import Document as _Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph

from image_pipeline import process_document_images


PROCESSOR_BUILD_ID = "2026.09.23-vml-images-v4"

MATH_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
NS = {"m": MATH_NS, "w": WORD_NS, "wp": DRAWING_NS}

QUESTION_START_RE = re.compile(r"^@Question:\s*(\d+)\s*@$", re.I)
QUESTION_START_WITH_SUFFIX_RE = re.compile(r"^@Question:\s*(\d+)\s*@\s+(.+?)\s*$", re.I)
QUESTION_ID_RE = re.compile(r"^@Question id:\s*(.*?)\s*@$", re.I)
PLAIN_QUESTION_START_RE = re.compile(
    r"^(?:Question|Q)\s*[:.\-]?\s*(\d+)\s*[).:]?\s*(Easy|Medium|Average|Challenging|Hard)?(?:\s*,\s*(Knowledge|Comprehension|Application|Analysis))?\s*$",
    re.I,
)
KNOWN_METADATA_RE = re.compile(
    r"^@(Type|Question id|New snippet id|Difficulty level|Objective):\s*(.*?)\s*@$",
    re.I,
)
PLAIN_METADATA_RE = re.compile(
    r"^(Type|Question id|New snippet id|Difficulty(?: level)?|Objective):\s*(.*?)\s*$",
    re.I,
)
SECTION_ALIASES = {
    "question": "@Question:@",
    "question:": "@Question:@",
    "@question:@": "@Question:@",
    "answers": "@Answers:@",
    "answer": "@Answers:@",
    "answers:": "@Answers:@",
    "answer:": "@Answers:@",
    "@answers:@": "@Answers:@",
    "choices": "@Choices:@",
    "options": "@Choices:@",
    "choices:": "@Choices:@",
    "options:": "@Choices:@",
    "@choices:@": "@Choices:@",
    "solution": "@Solution:@",
    "solution:": "@Solution:@",
    "@solution:@": "@Solution:@",
}
SECTION_MARKERS = {"@Question:@", "@Answers:@", "@Choices:@", "@Solution:@"}
VARIABLE_ONLY_RE = re.compile(r"^[A-Za-z]$")
STANDALONE_VARIABLE_RE = re.compile(r"(?<![A-Za-z])([A-Za-z])(?![A-Za-z])")
SIMPLE_FRACTION_RE = re.compile(r"(?<![\w/])(-?\d+|[A-Za-z])\s*/\s*(-?\d+|[A-Za-z])(?![\w/])")
INLINE_EQUATION_RE = re.compile(r"(?:\b\d+\b|\b[A-Za-z]\b)\s*=\s*(?:\b\d+\b|\b[A-Za-z]\b)")
ROMAN_RE = re.compile(r"^\s*\((i|ii|iii|iv|v|vi|vii|viii|ix|x)\)\s*", re.I)
ALPHA_RE = re.compile(r"^\s*\(([a-z])\)\s*", re.I)
OPTION_RE = re.compile(r"^\s*(?:@([1-4])@|\(?([A-Da-d])\)?[.)])\s*(.*)$")
TAGGED_CHOICE_RE = re.compile(r"^(\s*@[1-4]@\s*)(.*?)(\s+@correct answer@\s*)?$", re.I)
SAFE_ROOT_ASSIGNMENT_RE = re.compile(r"^([A-Za-z])\s*=\s*√\s*(\d+)$")
SAFE_FRACTION_ASSIGNMENT_RE = re.compile(r"^([A-Za-z])\s*=\s*(-?\d+|[A-Za-z])\s*/\s*(-?\d+|[A-Za-z])$")
LETTER_ANSWER_RE = re.compile(r"^\s*Answer\s*:?\s*([A-Da-d])\s*$", re.I)
ANSWER_VALUE_RE = re.compile(r"^\s*Answer\s*:?\s*(.+?)\s*$", re.I)
DIFFICULTY_ALIASES = {
    "easy": "Easy",
    "medium": "Average",
    "average": "Average",
    "hard": "Challenging",
    "challenging": "Challenging",
}


@dataclass
class ProcessorOptions:
    project_question_prefix: str = "project10436_q"
    start_question_number: int = 1
    start_snippet_id: int = 218989
    default_type: str = "FIB"
    default_difficulty: str = "Average"
    default_objective: str = "Application"
    replace_existing_ids: bool = True
    processing_mode: str = "full"


@dataclass
class Finding:
    rule: int
    status: str
    message: str
    question: int | None = None
    paragraph: int | None = None


@dataclass
class ProcessResult:
    output_path: Path
    report_path: Path
    manifest_path: Path
    image_manifest_path: Path
    images_zip_path: Path
    original_path: Path
    run_id: str
    question_count: int
    image_count: int
    findings: list[Finding]

    @property
    def manual_review_count(self) -> int:
        return sum(f.status == "manual_review" for f in self.findings)

    @property
    def fixed_count(self) -> int:
        return sum(f.status == "fixed" for f in self.findings)


def _safe_filename(name: str) -> str:
    base = Path(name).name
    stem = re.sub(r"[^A-Za-z0-9._ -]+", "_", base).strip(" .")
    return stem or "uploaded.docx"


def _storage_root(explicit: str | Path | None = None) -> Path:
    if explicit:
        return Path(explicit)
    return Path(os.environ.get("CMS_STORAGE_DIR", Path(__file__).parent / "data"))


def _existing_project_id(doc: _Document) -> tuple[str | None, set[str]]:
    project_ids: set[str] = set()
    for paragraph in iter_paragraphs(doc):
        match = QUESTION_ID_RE.fullmatch(paragraph.text.strip())
        if not match:
            continue
        question_id = match.group(1).strip()
        project_match = re.fullmatch(r"(.+)_q\d+", question_id, re.I)
        if project_match:
            project_ids.add(project_match.group(1))
    if len(project_ids) == 1:
        return next(iter(project_ids)), project_ids
    return None, project_ids


def iter_paragraphs(parent: _Document | _Cell) -> Iterator[Paragraph]:
    """Yield body and table-cell paragraphs in document order."""
    for child in parent.element.body.iterchildren() if isinstance(parent, _Document) else parent._tc.iterchildren():
        if child.tag == qn("w:p"):
            yield Paragraph(child, parent)
        elif child.tag == qn("w:tbl"):
            table = Table(child, parent)
            for row in table.rows:
                for cell in row.cells:
                    yield from iter_paragraphs(cell)


def body_paragraphs(doc: _Document) -> list[Paragraph]:
    return [Paragraph(p, doc) for p in doc.element.body.findall(qn("w:p"))]


def _set_text(paragraph: Paragraph, text: str) -> None:
    paragraph.clear()
    paragraph.add_run(text)


def _is_bold_cms_tag(text: str) -> bool:
    """Match the standalone metadata and section tags that are bold in the CMS template."""
    stripped = text.strip()
    return bool(
        QUESTION_START_RE.fullmatch(stripped)
        or KNOWN_METADATA_RE.fullmatch(stripped)
        or stripped in SECTION_MARKERS
    )


def _bold_cms_tags(doc: _Document, findings: list[Finding]) -> None:
    changed_paragraphs = 0
    for paragraph in iter_paragraphs(doc):
        if not _is_bold_cms_tag(paragraph.text):
            continue
        changed = False
        for run in paragraph.runs:
            if run.text and run.bold is not True:
                run.bold = True
                changed = True
        if changed:
            changed_paragraphs += 1
    findings.append(
        Finding(
            0,
            "fixed" if changed_paragraphs else "passed",
            (
                f"Applied bold formatting to {changed_paragraphs} CMS metadata or section tag paragraph(s)."
                if changed_paragraphs
                else "All CMS metadata and section tags were already bold."
            ),
        )
    )


def _insert_after(paragraph: Paragraph, text: str = "") -> Paragraph:
    new_p = OxmlElement("w:p")
    paragraph._p.addnext(new_p)
    new_paragraph = Paragraph(new_p, paragraph._parent)
    if text:
        new_paragraph.add_run(text)
    return new_paragraph


def _remove_paragraph(paragraph: Paragraph) -> None:
    parent = paragraph._p.getparent()
    if parent is not None:
        parent.remove(paragraph._p)


def _question_starts(doc: _Document) -> list[tuple[Paragraph, int]]:
    starts: list[tuple[Paragraph, int]] = []
    for paragraph in body_paragraphs(doc):
        text = paragraph.text.strip()
        match = (
            QUESTION_START_RE.fullmatch(text)
            or QUESTION_START_WITH_SUFFIX_RE.fullmatch(text)
            or PLAIN_QUESTION_START_RE.fullmatch(text)
        )
        if match:
            starts.append((paragraph, int(match.group(1))))
    return starts


def _canonicalise_section_labels(doc: _Document, findings: list[Finding]) -> None:
    for index, paragraph in enumerate(iter_paragraphs(doc), 1):
        raw = paragraph.text.strip()
        key = raw.casefold()
        if key in SECTION_ALIASES and raw != SECTION_ALIASES[key]:
            _set_text(paragraph, SECTION_ALIASES[key])
            findings.append(Finding(0, "fixed", f"Converted section label to {SECTION_ALIASES[key]}", paragraph=index))


def _block_paragraphs(start: Paragraph, next_start: Paragraph | None) -> list[Paragraph]:
    result: list[Paragraph] = []
    node = start._p
    stop = next_start._p if next_start else None
    while node is not None and node is not stop:
        if node.tag == qn("w:p"):
            result.append(Paragraph(node, start._parent))
        node = node.getnext()
    return result


def _metadata_value(block: Iterable[Paragraph], name: str) -> str | None:
    for p in block:
        match = KNOWN_METADATA_RE.fullmatch(p.text.strip()) or PLAIN_METADATA_RE.fullmatch(p.text.strip())
        field_name = match.group(1).casefold() if match else ""
        if field_name == "difficulty":
            field_name = "difficulty level"
        if match and field_name == name.casefold():
            return match.group(2).strip()
    return None


def _set_text_preserving_run_format(paragraph: Paragraph, text: str) -> None:
    """Replace a simple tag paragraph without discarding its run formatting."""
    source_properties = None
    for run in paragraph.runs:
        if run._r.rPr is not None:
            source_properties = deepcopy(run._r.rPr)
            break
    paragraph.clear()
    run = paragraph.add_run(text)
    if source_properties is not None:
        if run._r.rPr is not None:
            run._r.remove(run._r.rPr)
        run._r.insert(0, source_properties)


def _audit_question_types(doc: _Document, findings: list[Finding], *, fix_mismatches: bool) -> None:
    """Check a tagged Type value against an unambiguous Answers/Choices structure."""
    starts = _question_starts(doc)
    checked = 0
    issues = 0
    for ordinal, (start, question_number) in enumerate(starts):
        next_start = starts[ordinal + 1][0] if ordinal + 1 < len(starts) else None
        block = _block_paragraphs(start, next_start)
        type_paragraph = None
        declared_type = None
        for paragraph in block:
            match = KNOWN_METADATA_RE.fullmatch(paragraph.text.strip())
            if match and match.group(1).casefold() == "type":
                type_paragraph = paragraph
                declared_type = match.group(2).strip().upper()
                break

        has_answers = any(paragraph.text.strip() == "@Answers:@" for paragraph in block)
        has_choices = any(paragraph.text.strip() == "@Choices:@" for paragraph in block)
        choice_numbers: set[int] = set()
        in_choices = False
        for paragraph in block:
            text = paragraph.text.strip()
            if text == "@Choices:@":
                in_choices = True
                continue
            if in_choices and (text == "@e@" or text in SECTION_MARKERS):
                in_choices = False
            if in_choices:
                marker = re.match(r"^@([1-4])@", text)
                if marker:
                    choice_numbers.add(int(marker.group(1)))

        expected_type = None
        structure_description = None
        if has_choices and not has_answers and choice_numbers == {1, 2, 3, 4}:
            expected_type = "MCQ"
            structure_description = "a @Choices:@ block with choices @1@ to @4@"
        elif has_answers and not has_choices:
            expected_type = "FIB"
            structure_description = "an @Answers:@ block and no @Choices:@ block"
        elif has_choices and has_answers:
            issues += 1
            findings.append(Finding(12, "manual_review", "Both @Choices:@ and @Answers:@ occur in the record, so the question type cannot be corrected safely.", question_number))
            continue
        elif has_choices and choice_numbers != {1, 2, 3, 4}:
            issues += 1
            visible = ", ".join(str(number) for number in sorted(choice_numbers)) or "none"
            findings.append(Finding(12, "manual_review", f"The @Choices:@ block contains choice markers {visible}; four choices @1@ to @4@ are required before the Type tag can be corrected safely.", question_number))
            continue
        else:
            continue

        checked += 1
        if type_paragraph is None or declared_type not in {"FIB", "MCQ"}:
            issues += 1
            findings.append(Finding(12, "manual_review", f"The structure indicates {expected_type}, but a valid @Type:@ tag was not found.", question_number))
            continue
        if declared_type == expected_type:
            continue

        issues += 1
        if fix_mismatches:
            _set_text_preserving_run_format(type_paragraph, f"@Type: {expected_type}@")
            findings.append(Finding(12, "fixed", f"Changed @Type: {declared_type}@ to @Type: {expected_type}@ because the record contains {structure_description}.", question_number))
        else:
            findings.append(Finding(12, "manual_review", f"The Type tag says {declared_type}, but the structure indicates {expected_type} because the record contains {structure_description}. Images and Alt Text only mode preserved the tag unchanged.", question_number))

    if checked and not issues:
        findings.append(Finding(12, "passed", f"Question Type tags matched the detected Answers/Choices structure in {checked} record(s)."))


def _has_visible_content(paragraph: Paragraph) -> bool:
    return bool(paragraph.text.strip() or paragraph._p.xpath(".//w:drawing | .//w:pict | .//m:oMath"))


def _split_compound_choices(paragraph: Paragraph) -> list[Paragraph] | None:
    lines = [line.strip() for line in paragraph.text.splitlines() if line.strip()]
    if len(lines) != 4 or not all(OPTION_RE.fullmatch(line) for line in lines):
        return None
    _set_text(paragraph, lines[0])
    choices = [paragraph]
    cursor = paragraph
    for line in lines[1:]:
        cursor = _insert_after(cursor, line)
        choices.append(cursor)
    return choices


def _write_choice_marker(paragraph: Paragraph, option: int, correct: bool) -> None:
    suffix = " @correct answer@" if correct else ""
    if paragraph._p.xpath(".//w:drawing | .//w:pict"):
        prefix = paragraph.add_run(f"@{option}@")
        paragraph._p.remove(prefix._r)
        paragraph._p.insert(0, prefix._r)
        if correct:
            paragraph.add_run(suffix)
        return
    raw = paragraph.text.strip()
    match = OPTION_RE.fullmatch(raw)
    content = match.group(3).strip() if match else raw
    _set_text(paragraph, f"@{option}@ {content}{suffix}".rstrip())


def _infer_untagged_mcq_sections(doc: _Document, findings: list[Finding]) -> None:
    """Structure common untagged MCQs only when a letter answer and four choices are clear."""
    starts = _question_starts(doc)
    for ordinal, (start, detected_number) in enumerate(starts):
        next_start = starts[ordinal + 1][0] if ordinal + 1 < len(starts) else None
        block = _block_paragraphs(start, next_start)
        if any(p.text.strip() in SECTION_MARKERS for p in block):
            continue
        answer_candidates = [(i, LETTER_ANSWER_RE.fullmatch(p.text.strip())) for i, p in enumerate(block)]
        answer_candidates = [(i, match) for i, match in answer_candidates if match]
        if len(answer_candidates) != 1:
            continue
        answer_index, answer_match = answer_candidates[0]
        correct_option = ord(answer_match.group(1).upper()) - ord("A") + 1

        choice_paragraphs: list[Paragraph] | None = None
        for paragraph in block[1:answer_index]:
            compound = _split_compound_choices(paragraph)
            if compound is not None:
                choice_paragraphs = compound
                break
        if choice_paragraphs is None:
            candidates = [p for p in block[1:answer_index] if _has_visible_content(p)]
            if len(candidates) >= 5:
                choice_paragraphs = candidates[-4:]
        if not choice_paragraphs or len(choice_paragraphs) != 4:
            findings.append(Finding(0, "manual_review", "A letter answer was found, but four MCQ choices could not be identified safely.", detected_number))
            continue

        choice_paragraphs[0].insert_paragraph_before("@Choices:@")
        for option, paragraph in enumerate(choice_paragraphs, 1):
            _write_choice_marker(paragraph, option, option == correct_option)

        # Locate the answer paragraph again because splitting a compound choice can change the block.
        refreshed = _block_paragraphs(start, next_start)
        answer_paragraph = next((p for p in refreshed if LETTER_ANSWER_RE.fullmatch(p.text.strip())), None)
        if answer_paragraph is not None:
            answer_paragraph.insert_paragraph_before("@Solution:@")
            # Do not invent solution prose when the source MCQ has no solution.
            # CMS still requires the Solution block and its closing marker.
            _set_text(answer_paragraph, "@e@")
        findings.append(Finding(0, "fixed", "Identified an untagged MCQ, created its Choices and Solution sections, and marked the keyed answer.", detected_number))


def _infer_untagged_fib_sections(doc: _Document, findings: list[Finding]) -> None:
    """Structure an untagged FIB when a single non-letter Answer value is explicit."""
    starts = _question_starts(doc)
    for ordinal, (start, detected_number) in enumerate(starts):
        next_start = starts[ordinal + 1][0] if ordinal + 1 < len(starts) else None
        block = _block_paragraphs(start, next_start)
        if any(p.text.strip() in SECTION_MARKERS for p in block):
            continue
        candidates: list[tuple[Paragraph, re.Match[str]]] = []
        for paragraph in block:
            match = ANSWER_VALUE_RE.fullmatch(paragraph.text.strip())
            if match and not LETTER_ANSWER_RE.fullmatch(paragraph.text.strip()):
                candidates.append((paragraph, match))
        if len(candidates) != 1:
            continue
        answer_paragraph, answer_match = candidates[0]
        answer_value = answer_match.group(1).strip()
        if not answer_value or len(answer_value) > 100 or "\n" in answer_value:
            findings.append(Finding(0, "manual_review", "An FIB-style answer was found but is too complex to structure automatically.", detected_number))
            continue
        answer_paragraph.insert_paragraph_before("@Answers:@")
        _set_text(answer_paragraph, answer_value)
        solution_marker = _insert_after(answer_paragraph, "@Solution:@")
        _insert_after(solution_marker, f"The answer is {answer_value}.")
        findings.append(Finding(0, "fixed", "Identified an untagged FIB and created its Answers and Solution sections.", detected_number))


def _ensure_cms_records(doc: _Document, options: ProcessorOptions, findings: list[Finding]) -> int:
    _canonicalise_section_labels(doc, findings)
    _infer_untagged_mcq_sections(doc, findings)
    _infer_untagged_fib_sections(doc, findings)
    starts = _question_starts(doc)
    if not starts:
        findings.append(Finding(0, "manual_review", "No question headings were detected. Use @Question: n@ or a standalone heading such as Question 1."))
        return 0

    # Work backwards so insertions do not disturb unprocessed record boundaries.
    for ordinal in range(len(starts) - 1, -1, -1):
        start, detected_number = starts[ordinal]
        next_start = starts[ordinal + 1][0] if ordinal + 1 < len(starts) else None
        block = _block_paragraphs(start, next_start)
        assigned_number = options.start_question_number + ordinal
        suffix_match = QUESTION_START_WITH_SUFFIX_RE.fullmatch(start.text.strip())
        heading_suffix = suffix_match.group(2).strip() if suffix_match else None
        heading_match = PLAIN_QUESTION_START_RE.fullmatch(start.text.strip())
        heading_difficulty = DIFFICULTY_ALIASES.get((heading_match.group(2) or "").casefold()) if heading_match else None
        heading_objective = heading_match.group(3).title() if heading_match and heading_match.group(3) else None
        _set_text(start, f"@Question: {assigned_number}@")

        existing_type = (_metadata_value(block, "Type") or "").upper()
        has_choices = any(p.text.strip() == "@Choices:@" for p in block)
        question_type = existing_type if existing_type in {"FIB", "MCQ"} else ("MCQ" if has_choices else options.default_type)
        difficulty = _metadata_value(block, "Difficulty level") or heading_difficulty or options.default_difficulty
        objective = _metadata_value(block, "Objective") or heading_objective or options.default_objective
        existing_qid = _metadata_value(block, "Question id")
        existing_snippet = _metadata_value(block, "New snippet id")
        qid = f"{options.project_question_prefix}{assigned_number}" if options.replace_existing_ids or not existing_qid else existing_qid
        snippet = str(options.start_snippet_id + ordinal) if options.replace_existing_ids or not existing_snippet else existing_snippet

        for paragraph in list(block[1:]):
            if KNOWN_METADATA_RE.fullmatch(paragraph.text.strip()) or PLAIN_METADATA_RE.fullmatch(paragraph.text.strip()):
                _remove_paragraph(paragraph)

        cursor = start
        metadata = [
            f"@Type: {question_type}@",
            f"@Question id: {qid} @",
            f"@New snippet id: {snippet} @",
            f"@Difficulty level: {difficulty} @",
            f"@Objective: {objective} @",
        ]
        for line in metadata:
            cursor = _insert_after(cursor, line)

        refreshed = _block_paragraphs(start, next_start)
        q_marker = next((p for p in refreshed if p.text.strip() == "@Question:@"), None)
        if q_marker is None:
            q_marker = _insert_after(cursor, "@Question:@")
            findings.append(Finding(0, "fixed", "Inserted the missing @Question:@ marker.", assigned_number))
        if heading_suffix:
            _insert_after(q_marker, heading_suffix)
            findings.append(
                Finding(
                    0,
                    "manual_review",
                    f"The question heading contained text after its closing tag: {heading_suffix!r}. The text was preserved at the start of the Question block; verify its placement.",
                    assigned_number,
                )
            )

        findings.append(Finding(0, "fixed", f"Normalised CMS metadata and assigned question ID {qid} and snippet ID {snippet}.", assigned_number))

    # Recompute and close each content block immediately before the next section.
    starts = _question_starts(doc)
    for ordinal, (start, _) in enumerate(starts):
        next_start = starts[ordinal + 1][0] if ordinal + 1 < len(starts) else None
        assigned_number = options.start_question_number + ordinal
        block = _block_paragraphs(start, next_start)
        section_positions = [i for i, p in enumerate(block) if p.text.strip() in SECTION_MARKERS]
        if not any(block[i].text.strip() in {"@Answers:@", "@Choices:@"} for i in section_positions):
            findings.append(Finding(8, "manual_review", "The record has no @Answers:@ or @Choices:@ block.", assigned_number))
        if not any(block[i].text.strip() == "@Solution:@" for i in section_positions):
            findings.append(Finding(1, "manual_review", "The record has no @Solution:@ block.", assigned_number))

        for pos_index in range(len(section_positions) - 1, -1, -1):
            pos = section_positions[pos_index]
            end = section_positions[pos_index + 1] if pos_index + 1 < len(section_positions) else len(block)
            content = block[pos + 1 : end]
            last_nonblank = next(
                (
                    p
                    for p in reversed(content)
                    if p.text.strip() or p._p.xpath(".//w:drawing | .//w:pict | .//m:oMath")
                ),
                None,
            )
            if last_nonblank is None:
                continue
            if last_nonblank.text.strip() != "@e@":
                _insert_after(last_nonblank, "@e@")
                findings.append(Finding(0, "fixed", f"Inserted a missing @e@ after {block[pos].text.strip()}.", assigned_number))

    return len(starts)


def _section_for_paragraphs(doc: _Document) -> dict[object, tuple[str | None, int | None]]:
    state: dict[object, tuple[str | None, int | None]] = {}
    section: str | None = None
    question: int | None = None
    for paragraph in body_paragraphs(doc):
        text = paragraph.text.strip()
        start = QUESTION_START_RE.fullmatch(text) or QUESTION_START_WITH_SUFFIX_RE.fullmatch(text)
        if start:
            question = int(start.group(1))
            section = None
        elif text in SECTION_MARKERS:
            section = text
        elif text == "@e@":
            section = None
        state[paragraph._p] = (section, question)
    return state


def _remove_bold_from_questions_and_solutions(doc: _Document, findings: list[Finding]) -> None:
    context = _section_for_paragraphs(doc)
    sections: dict[tuple[int | None, str], list[Paragraph]] = {}
    for paragraph in body_paragraphs(doc):
        section, question = context.get(paragraph._p, (None, None))
        if section not in {"@Question:@", "@Solution:@"}:
            continue
        if paragraph.text.strip() in SECTION_MARKERS or paragraph.text.strip() == "@e@":
            continue
        sections.setdefault((question, section), []).append(paragraph)

    changed_sections = 0
    emphasis_sections = 0
    for (question, section), paragraphs in sections.items():
        substantive_runs: list[tuple[Paragraph, object]] = []
        for paragraph in paragraphs:
            for run in paragraph.runs:
                if run.text.strip():
                    substantive_runs.append((paragraph, run))
        if not substantive_runs:
            continue

        def effectively_bold(paragraph: Paragraph, run) -> bool:
            if run.bold is not None:
                return bool(run.bold)
            if run.style is not None and run.style.font.bold is not None:
                return bool(run.style.font.bold)
            if paragraph.style is not None and paragraph.style.font.bold is not None:
                return bool(paragraph.style.font.bold)
            return False

        bold_flags = [effectively_bold(paragraph, run) for paragraph, run in substantive_runs]
        if all(bold_flags):
            for _, run in substantive_runs:
                run.bold = False
            changed_sections += 1
            label = "Question" if section == "@Question:@" else "Solution"
            findings.append(Finding(1, "fixed", f"Removed whole-block bold formatting from the {label.lower()} while preserving the text.", question))
        elif any(bold_flags):
            emphasis_sections += 1
            label = "question" if section == "@Question:@" else "solution"
            findings.append(Finding(1, "passed", f"Preserved selective bold emphasis within the {label}; the entire block was not bold.", question))

    if not changed_sections and not emphasis_sections:
        findings.append(Finding(1, "passed", "No whole-question or whole-solution bold formatting was found."))
    else:
        findings.append(Finding(1, "passed", f"Bold-format review completed: {changed_sections} whole block(s) corrected and {emphasis_sections} selectively emphasised block(s) preserved."))


def _remove_empty_scripts(doc: _Document, findings: list[Finding]) -> None:
    root = doc.element
    removed = 0
    for tag, child_tag in (("sSup", "sup"), ("sSub", "sub")):
        scripts = list(root.xpath(f".//m:{tag}"))
        for script in reversed(scripts):
            slot = script.find(qn(f"m:{child_tag}"))
            if slot is not None and "".join(slot.itertext()).strip():
                continue
            base = script.find(qn("m:e"))
            parent = script.getparent()
            if base is None or parent is None:
                continue
            position = parent.index(script)
            parent.remove(script)
            for child in list(base):
                base.remove(child)
                parent.insert(position, child)
                position += 1
            removed += 1
    blank_equations = 0
    for equation in reversed(list(root.xpath(".//m:oMath"))):
        visible_text = "".join(node.text or "" for node in equation.iter(qn("m:t"))).strip()
        if visible_text:
            continue
        parent = equation.getparent()
        if parent is not None:
            parent.remove(equation)
            blank_equations += 1
    total = removed + blank_equations
    if total:
        findings.append(Finding(4, "fixed", f"Removed {removed} blank exponent/subscript template(s) and {blank_equations} completely blank equation object(s)."))
    else:
        findings.append(Finding(4, "passed", "No blank equation, exponent or subscript templates were found."))


def _normalise_spacing(text: str) -> str:
    if not text:
        return text
    text = text.replace("\u00a0", " ")
    text = re.sub(r"[ \t]{2,}", " ", text)
    text = re.sub(r"\s*,\s*", ", ", text)
    text = re.sub(r"\s*([=+×÷−])\s*", r" \1 ", text)
    text = re.sub(r" +([.;:?!])", r"\1", text)
    return text.strip() if text.strip().startswith("@") and text.strip().endswith("@") else text


def _repair_spacing(doc: _Document, findings: list[Finding]) -> None:
    changes = 0
    for node in doc.element.xpath(".//w:t | .//m:t"):
        original = node.text or ""
        updated = _normalise_spacing(original)
        if updated != original:
            node.text = updated
            changes += 1
    findings.append(Finding(4, "fixed" if changes else "passed", f"Normalised spacing in {changes} text run(s)." if changes else "No repeated or operator-spacing defects were found."))
    findings.append(Finding(5, "passed", "Operator and equals-sign spacing was checked and normalised."))
    findings.append(Finding(6, "passed", "Comma spacing was checked and normalised."))


def _variable_spans(text: str) -> list[tuple[int, int]]:
    """Return conservative variable spans in an ordinary-text run."""
    if not text:
        return []
    tag_ranges = [(m.start(), m.end()) for m in re.finditer(r"@[^@]*@", text)]
    tokens = list(STANDALONE_VARIABLE_RE.finditer(text))
    uppercase_count = sum(1 for token in tokens if token.group(1).isupper())
    has_math_signal = bool(re.search(r"[=+×÷−<>≤≥/]", text) or re.search(r"\d", text))
    cue_positions: set[tuple[int, int]] = set()
    for cue in re.finditer(r"\b(?:let|value\s+of|solve\s+for|find)\s+([A-Za-z])\b", text, re.I):
        cue_positions.add(cue.span(1))

    spans: list[tuple[int, int]] = []
    for token in tokens:
        start, end = token.span(1)
        if any(left <= start < right for left, right in tag_ranges):
            continue
        letter = token.group(1)
        left = text[:start].rstrip()[-1:] or ""
        right = text[end:].lstrip()[:1] or ""
        local_math = left in "=+×÷−<>≤≥/(" or right in "=+×÷−<>≤≥/)" or left.isdigit() or right.isdigit()
        explicit_cue = (start, end) in cue_positions
        paired_capital = letter.isupper() and uppercase_count >= 2
        likely_symbol = has_math_signal and letter.lower() not in {"a", "i"}
        if local_math or explicit_cue or paired_capital or likely_symbol or text.strip() == letter:
            spans.append((start, end))
    return spans


def _copy_run_with_italic(paragraph: Paragraph, source_run, text: str, italic: bool) -> None:
    new_run = paragraph.add_run(text)
    source_properties = source_run._r.find(qn("w:rPr"))
    if source_properties is not None:
        existing = new_run._r.find(qn("w:rPr"))
        if existing is not None:
            new_run._r.remove(existing)
        new_run._r.insert(0, deepcopy(source_properties))
    new_run.italic = italic
    source_run._r.addprevious(new_run._r)


def _apply_variable_italics(paragraph: Paragraph) -> int:
    changed = 0
    for run in list(paragraph.runs):
        spans = _variable_spans(run.text)
        if not spans:
            continue
        if len(spans) == 1 and spans[0] == (0, len(run.text)):
            if run.italic is not True:
                run.italic = True
                changed += 1
            continue

        cursor = 0
        segments: list[tuple[str, bool]] = []
        for start, end in spans:
            if start > cursor:
                segments.append((run.text[cursor:start], False))
            segments.append((run.text[start:end], True))
            cursor = end
        if cursor < len(run.text):
            segments.append((run.text[cursor:], False))
        for segment, italic in segments:
            if segment:
                _copy_run_with_italic(paragraph, run, segment, italic)
        run._r.getparent().remove(run._r)
        changed += len(spans)
    return changed


def _repair_italics(doc: _Document, findings: list[Finding]) -> None:
    removed = 0
    added = 0
    context = _section_for_paragraphs(doc)
    for paragraph in body_paragraphs(doc):
        section, _question = context.get(paragraph._p, (None, None))
        if section not in {"@Question:@", "@Answers:@", "@Choices:@", "@Solution:@"}:
            continue
        for run in paragraph.runs:
            if not run.italic:
                continue
            text = run.text.strip()
            if VARIABLE_ONLY_RE.fullmatch(text):
                continue
            run.italic = False
            removed += 1
        added += _apply_variable_italics(paragraph)

    # Explicitly keep variables italic and non-variable equation content plain.
    for math_run in doc.element.xpath(".//m:r"):
        text = "".join(math_run.itertext()).strip()
        style = math_run.find("m:rPr/m:sty", namespaces=NS)
        if VARIABLE_ONLY_RE.fullmatch(text):
            if style is None:
                properties = math_run.find(qn("m:rPr"))
                if properties is None:
                    properties = OxmlElement("m:rPr")
                    math_run.insert(0, properties)
                style = OxmlElement("m:sty")
                properties.append(style)
            if style.get(qn("m:val")) != "i":
                style.set(qn("m:val"), "i")
                added += 1
        elif style is not None and style.get(qn("m:val")) == "i":
            style.set(qn("m:val"), "p")
            removed += 1

    if removed or added:
        findings.append(Finding(3, "fixed", f"Corrected italics in student-facing content: removed italics from {removed} non-variable run(s) and italicised {added} variable occurrence(s), including variables in choices."))
    else:
        findings.append(Finding(3, "passed", "Only variables are italic in questions, answers, choices and solutions."))


def _equation_and_fraction_audit(doc: _Document, findings: list[Finding]) -> None:
    native_count = len(doc.element.xpath(".//m:oMath"))
    findings.append(Finding(2, "passed", f"Detected {native_count} native Word equation object(s)."))
    context = _section_for_paragraphs(doc)
    fraction_hits = 0
    image_hits = 0
    math_text_hits = 0
    ambiguous_scope_hits = 0
    for index, paragraph in enumerate(body_paragraphs(doc), 1):
        section, question = context.get(paragraph._p, (None, None))
        if section not in {"@Question:@", "@Solution:@", "@Answers:@", "@Choices:@"}:
            continue
        text = paragraph.text
        ambiguous_scope = ("√" in text and ("/" in text or "^" in text)) or bool(re.search(r"[\^⁰¹²³⁴⁵⁶⁷⁸⁹]", text))
        if SIMPLE_FRACTION_RE.search(text) and not ambiguous_scope:
            fraction_hits += 1
            findings.append(Finding(7, "manual_review", f"Slash-style fraction requires conversion to a stacked Word equation: {text.strip()!r}", question, index))
        if paragraph._p.xpath(".//w:drawing | .//w:pict"):
            image_hits += 1
            findings.append(Finding(2, "manual_review", "An embedded image occurs in mathematical content. Verify that it is a diagram, not an equation screenshot.", question, index))
        if not paragraph._p.xpath(".//m:oMath") and (
            re.fullmatch(r"\s*[A-Za-z0-9().,]+(?:\s*[=+×÷−]\s*[A-Za-z0-9().,]+)+\s*", text)
            or INLINE_EQUATION_RE.search(text)
        ):
            math_text_hits += 1
            findings.append(Finding(2, "manual_review", f"A math-like expression is ordinary text and should be recreated with Insert → Equation: {text.strip()!r}", question, index))
        if ambiguous_scope:
            ambiguous_scope_hits += 1
            findings.append(
                Finding(
                    2,
                    "manual_review",
                    f"An exponent or radical expression has potentially ambiguous scope and was not changed automatically: {text.strip()!r}",
                    question,
                    index,
                )
            )
    if not fraction_hits:
        findings.append(Finding(7, "passed", "No slash-style fractions were detected in student-facing content."))
    if not image_hits and not math_text_hits and not ambiguous_scope_hits:
        findings.append(Finding(2, "passed", "No likely equation screenshots or ordinary-text equations were detected."))


def _math_run(text: str) -> OxmlElement:
    run = OxmlElement("m:r")
    value = OxmlElement("m:t")
    value.text = text
    run.append(value)
    return run


def _math_fraction(numerator: str, denominator: str) -> OxmlElement:
    fraction = OxmlElement("m:f")
    num = OxmlElement("m:num")
    den = OxmlElement("m:den")
    num.append(_math_run(numerator))
    den.append(_math_run(denominator))
    fraction.append(num)
    fraction.append(den)
    return fraction


def _math_radical(radicand: str) -> OxmlElement:
    radical = OxmlElement("m:rad")
    properties = OxmlElement("m:radPr")
    degree_hidden = OxmlElement("m:degHide")
    degree_hidden.set(qn("m:val"), "1")
    properties.append(degree_hidden)
    degree = OxmlElement("m:deg")
    expression = OxmlElement("m:e")
    expression.append(_math_run(radicand))
    radical.extend((properties, degree, expression))
    return radical


def _word_text_run(text: str, run_properties) -> OxmlElement:
    run = OxmlElement("w:r")
    if run_properties is not None:
        run.append(deepcopy(run_properties))
    value = OxmlElement("w:t")
    if text[:1].isspace() or text[-1:].isspace():
        value.set(qn("xml:space"), "preserve")
    value.text = text
    run.append(value)
    return run


def _convert_safe_choice_equations(doc: _Document, findings: list[Finding]) -> None:
    """Convert only unambiguous tagged-choice assignments to native Word equations."""
    context = _section_for_paragraphs(doc)
    converted = 0
    for paragraph in body_paragraphs(doc):
        section, question = context.get(paragraph._p, (None, None))
        if section != "@Choices:@" or paragraph._p.xpath(".//m:oMath"):
            continue
        choice_match = TAGGED_CHOICE_RE.fullmatch(paragraph.text)
        if not choice_match:
            continue
        prefix, expression, suffix = choice_match.groups()
        expression = expression.strip()
        root_match = SAFE_ROOT_ASSIGNMENT_RE.fullmatch(expression)
        fraction_match = SAFE_FRACTION_ASSIGNMENT_RE.fullmatch(expression)
        if not root_match and not fraction_match:
            continue

        paragraph.clear()
        paragraph.add_run(prefix)
        equation = OxmlElement("m:oMath")
        if root_match:
            equation.append(_math_run(root_match.group(1)))
            equation.append(_math_run(" = "))
            equation.append(_math_radical(root_match.group(2)))
        else:
            equation.append(_math_run(fraction_match.group(1)))
            equation.append(_math_run(" = "))
            equation.append(_math_fraction(fraction_match.group(2), fraction_match.group(3)))
        paragraph._p.append(equation)
        if suffix:
            paragraph.add_run(suffix)
        converted += 1
        findings.append(Finding(2, "fixed", f"Converted an unambiguous choice expression to a native Word equation: {expression!r}", question))
    if not converted:
        findings.append(Finding(2, "passed", "No unambiguous ordinary-text radical or fraction assignments in choices required conversion."))


def _convert_inline_slash_fractions(doc: _Document, findings: list[Finding]) -> None:
    """Replace simple fractions inside ordinary runs with native stacked fractions."""
    context = _section_for_paragraphs(doc)
    converted = 0
    for paragraph in body_paragraphs(doc):
        section, question = context.get(paragraph._p, (None, None))
        if section not in {"@Question:@", "@Solution:@", "@Answers:@", "@Choices:@"}:
            continue
        for run in list(paragraph.runs):
            text = run.text
            matches = list(SIMPLE_FRACTION_RE.finditer(text))
            if not matches:
                continue
            # Exponents can change the scope of a nearby fraction. A radical
            # only blocks a fraction when it is directly attached to that
            # numerator; comma-separated fractions such as "√2, 1/5" remain
            # safe to convert.
            if re.search(r"[\^⁰¹²³⁴⁵⁶⁷⁸⁹]", text):
                continue
            matches = [match for match in matches if not text[: match.start()].rstrip().endswith("√")]
            if not matches:
                continue
            # Runs containing tabs, line breaks or drawings need a human-safe edit.
            if any(child.tag not in {qn("w:rPr"), qn("w:t")} for child in run._r):
                findings.append(Finding(7, "manual_review", f"A slash fraction occurs in a complex Word run and could not be converted safely: {text!r}", question))
                continue
            parent = run._r.getparent()
            if parent is None:
                continue
            position = parent.index(run._r)
            run_properties = run._r.rPr
            cursor = 0
            replacement_nodes: list[object] = []
            for match in matches:
                if match.start() > cursor:
                    replacement_nodes.append(_word_text_run(text[cursor : match.start()], run_properties))
                equation = OxmlElement("m:oMath")
                equation.append(_math_fraction(match.group(1), match.group(2)))
                replacement_nodes.append(equation)
                cursor = match.end()
                converted += 1
            if cursor < len(text):
                replacement_nodes.append(_word_text_run(text[cursor:], run_properties))
            parent.remove(run._r)
            for offset, node in enumerate(replacement_nodes):
                parent.insert(position + offset, node)
    if converted:
        findings.append(Finding(7, "fixed", f"Converted {converted} slash-style fraction(s) into native stacked Word fractions."))
    else:
        findings.append(Finding(7, "passed", "No safely convertible slash-style fractions were found in ordinary text runs."))


def _append_native_equation(paragraph: Paragraph, expression: str) -> None:
    equation = OxmlElement("m:oMath")
    cursor = 0
    for match in SIMPLE_FRACTION_RE.finditer(expression):
        if match.start() > cursor:
            equation.append(_math_run(expression[cursor : match.start()]))
        equation.append(_math_fraction(match.group(1), match.group(2)))
        cursor = match.end()
    if cursor < len(expression):
        equation.append(_math_run(expression[cursor:]))
    paragraph._p.append(equation)


def _convert_simple_math_paragraphs(doc: _Document, findings: list[Finding]) -> None:
    """Convert only unambiguous standalone expressions; prose remains for review."""
    context = _section_for_paragraphs(doc)
    converted = 0
    for paragraph in body_paragraphs(doc):
        section, question = context.get(paragraph._p, (None, None))
        if section not in {"@Question:@", "@Solution:@", "@Answers:@", "@Choices:@"}:
            continue
        if paragraph._p.xpath(".//m:oMath | .//w:drawing | .//w:pict"):
            continue
        raw = paragraph.text.strip()
        if not raw or raw.startswith("@"):
            continue
        label_match = re.match(r"^(\([a-z]\))\s+", raw, re.I)
        label = label_match.group(1) if label_match else ""
        expression = raw[label_match.end() :] if label_match else raw
        if re.search(r"[A-Za-z]{2,}", expression):
            continue
        if "^" in expression or not re.fullmatch(r"[A-Za-z0-9.,()\s=+×÷−*/]+", expression):
            continue
        if not re.search(r"[=+×÷−/]", expression):
            continue
        paragraph.clear()
        if label:
            paragraph.add_run(label + " ")
        _append_native_equation(paragraph, _normalise_spacing(expression))
        converted += 1
        findings.append(Finding(2, "fixed", f"Converted a standalone mathematical expression to a native Word equation: {raw!r}", question))
    if converted:
        findings.append(Finding(7, "fixed", f"Converted slash fractions to stacked fractions wherever they occurred in {converted} unambiguous standalone expression(s)."))


def _split_answers(doc: _Document, findings: list[Finding]) -> None:
    context = _section_for_paragraphs(doc)
    splits = 0
    for paragraph in list(body_paragraphs(doc)):
        section, question = context.get(paragraph._p, (None, None))
        if section != "@Answers:@":
            continue
        text = paragraph.text.strip()
        matches = list(re.finditer(r"(?<!\w)(\([a-z]\)|[a-z]\))\s*", text, re.I))
        if len(matches) < 2:
            continue
        parts: list[str] = []
        for idx, match in enumerate(matches):
            end = matches[idx + 1].start() if idx + 1 < len(matches) else len(text)
            parts.append(text[match.start() : end].strip())
        _set_text(paragraph, parts[0])
        cursor = paragraph
        for part in parts[1:]:
            cursor = _insert_after(cursor, part)
        splits += len(parts) - 1
        findings.append(Finding(8, "fixed", f"Placed {len(parts)} labelled answers on separate lines.", question))
    if not splits:
        findings.append(Finding(8, "passed", "No combined labelled answer lines required splitting."))


def _normalise_subparts(doc: _Document, findings: list[Finding]) -> None:
    context = _section_for_paragraphs(doc)
    candidates: list[tuple[Paragraph, re.Match[str], int | None]] = []
    alpha_exists = False
    for paragraph in body_paragraphs(doc):
        section, question = context.get(paragraph._p, (None, None))
        if section != "@Question:@":
            continue
        alpha_exists = alpha_exists or bool(ALPHA_RE.match(paragraph.text))
        match = ROMAN_RE.match(paragraph.text)
        if match:
            candidates.append((paragraph, match, question))
    if candidates and not alpha_exists:
        for idx, (paragraph, match, question) in enumerate(candidates):
            label = chr(ord("a") + idx)
            _set_text(paragraph, f"({label}) {paragraph.text[match.end():].strip()}")
        findings.append(Finding(9, "fixed", f"Converted {len(candidates)} top-level Roman-numeral subpart label(s) to alphabetic labels."))
    elif candidates and alpha_exists:
        findings.append(Finding(9, "manual_review", "Both alphabetic and Roman-numeral labels occur at the same visible level. Verify which Roman numerals are nested."))
    else:
        findings.append(Finding(9, "passed", "Top-level subpart labels are compatible with the required format."))


ASSERTION_CHOICES = [
    "Both Assertion (A) and Reason (R) are true and Reason (R) is the correct explanation of Assertion (A)",
    "Both Assertion (A) and Reason (R) are true and Reason (R) is not the correct explanation of Assertion (A)",
    "Assertion (A) is true but Reason (R) is false",
    "Assertion (A) is false but Reason (R) is true",
]


def _audit_assertion_reason(doc: _Document, findings: list[Finding]) -> None:
    starts = _question_starts(doc)
    found_any = False
    for ordinal, (start, number) in enumerate(starts):
        next_start = starts[ordinal + 1][0] if ordinal + 1 < len(starts) else None
        block = _block_paragraphs(start, next_start)
        texts = [p.text.strip() for p in block]
        if not any(t.startswith("Assertion (A):") for t in texts) or not any(t.startswith("Reason (R):") for t in texts):
            continue
        found_any = True
        intro = "A statement of Assertion (A) is followed by a statement of Reason (R)."
        assertion_index = next(i for i, t in enumerate(texts) if t.startswith("Assertion (A):"))
        if intro not in texts[:assertion_index]:
            block[assertion_index].insert_paragraph_before(intro)
            findings.append(Finding(10, "fixed", "Inserted the required Assertion–Reason introductory sentence.", number))

        choice_marker = next((p for p in block if p.text.strip() == "@Choices:@"), None)
        if choice_marker is None:
            findings.append(Finding(10, "manual_review", "Assertion–Reason item has no @Choices:@ block.", number))
            continue
        option_paragraphs: list[Paragraph] = []
        node = choice_marker._p.getnext()
        while node is not None and node is not (next_start._p if next_start else None):
            if node.tag == qn("w:p"):
                paragraph = Paragraph(node, choice_marker._parent)
                if paragraph.text.strip() in {"@e@", "@Solution:@"}:
                    break
                if OPTION_RE.match(paragraph.text):
                    option_paragraphs.append(paragraph)
            node = node.getnext()
        if len(option_paragraphs) != 4:
            findings.append(Finding(10, "manual_review", f"Assertion–Reason item has {len(option_paragraphs)} detectable choices; four are required.", number))
            continue
        correct_indexes = [i for i, p in enumerate(option_paragraphs) if "@correct answer@" in p.text]
        if len(correct_indexes) != 1:
            findings.append(Finding(10, "manual_review", "Assertion–Reason item must have exactly one @correct answer@ marker.", number))
        for i, paragraph in enumerate(option_paragraphs):
            suffix = " @correct answer@" if i in correct_indexes else ""
            _set_text(paragraph, f"@{i + 1}@ {ASSERTION_CHOICES[i]}{suffix}")
        findings.append(Finding(10, "fixed", "Normalised the four Assertion–Reason choice sentences while preserving the existing correct-answer position.", number))
    if not found_any:
        findings.append(Finding(10, "passed", "No Assertion–Reason items were detected."))


def _preserve_table_formatting(doc: _Document, findings: list[Finding]) -> None:
    if doc.tables:
        findings.append(Finding(11, "passed", "Preserved existing table formatting. Header rows or columns must be identified and bolded by the author."))
    else:
        findings.append(Finding(11, "passed", "The document has no tables."))


def _normalise_choices(doc: _Document, findings: list[Finding]) -> None:
    context = _section_for_paragraphs(doc)
    changed = 0
    for paragraph in body_paragraphs(doc):
        section, question = context.get(paragraph._p, (None, None))
        if section != "@Choices:@":
            continue
        match = OPTION_RE.fullmatch(paragraph.text.strip())
        if not match:
            continue
        digit, letter, rest = match.groups()
        option = int(digit) if digit else ord(letter.upper()) - ord("A") + 1
        if 1 <= option <= 4:
            canonical = f"@{option}@" + (f" {rest.strip()}" if rest.strip() else "")
            if paragraph.text.strip() != canonical:
                _set_text(paragraph, canonical)
                changed += 1
    if changed:
        findings.append(Finding(0, "fixed", f"Normalised {changed} MCQ option marker(s)."))


def process_docx(
    source: bytes | str | Path,
    original_filename: str,
    options: ProcessorOptions,
    storage_dir: str | Path | None = None,
) -> ProcessResult:
    if not original_filename.lower().endswith(".docx"):
        raise ValueError("Only .docx files are supported.")

    root = _storage_root(storage_dir)
    uploads = root / "uploads"
    processed = root / "processed"
    reports = root / "reports"
    queue = root / "verification_queue"
    images_root = root / "images"
    packages = root / "packages"
    for folder in (uploads, processed, reports, queue, images_root, packages):
        folder.mkdir(parents=True, exist_ok=True)

    run_id = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ") + "_" + uuid.uuid4().hex[:8]
    safe_name = _safe_filename(original_filename)
    original_path = uploads / f"{run_id}_{safe_name}"
    if isinstance(source, (str, Path)):
        shutil.copyfile(source, original_path)
    else:
        original_path.write_bytes(source)

    doc = Document(original_path)
    findings: list[Finding] = []
    preserve_cms_mode = options.processing_mode in {"images_only", "images_math"}
    if preserve_cms_mode:
        question_count = len(_question_starts(doc))
        if question_count:
            if options.processing_mode == "images_only":
                findings.append(Finding(0, "passed", "Images-only mode preserved the existing document text, CMS tags, metadata and formatting."))
            else:
                findings.append(Finding(0, "passed", "Safe-Math image mode preserved existing question IDs, snippet IDs and document structure; only a deterministically mismatched Type tag may be corrected."))
            for paragraph, number in _question_starts(doc):
                suffix_match = QUESTION_START_WITH_SUFFIX_RE.fullmatch(paragraph.text.strip())
                if suffix_match:
                    findings.append(
                        Finding(
                            0,
                            "manual_review",
                            f"The question heading contains text after its closing tag: {suffix_match.group(2).strip()!r}. It was counted as a question and preserved in the selected existing-CMS mode; move the text into the Question block before CMS upload.",
                            number,
                        )
                    )
        else:
            findings.append(Finding(0, "manual_review", "No @Question: n@ records were detected. Images cannot be assigned reliable CMS filenames without existing question tags."))
        _audit_question_types(doc, findings, fix_mismatches=options.processing_mode == "images_math")
        if options.processing_mode == "images_math":
            _remove_empty_scripts(doc, findings)
            _repair_spacing(doc, findings)
            _repair_italics(doc, findings)
            _convert_safe_choice_equations(doc, findings)
            _convert_inline_slash_fractions(doc, findings)
            _convert_simple_math_paragraphs(doc, findings)
            _equation_and_fraction_audit(doc, findings)
            _preserve_table_formatting(doc, findings)
    else:
        question_count = _ensure_cms_records(doc, options, findings)
        _audit_question_types(doc, findings, fix_mismatches=True)
        _normalise_choices(doc, findings)
        _remove_bold_from_questions_and_solutions(doc, findings)
        _remove_empty_scripts(doc, findings)
        _repair_spacing(doc, findings)
        _repair_italics(doc, findings)
        _split_answers(doc, findings)
        _normalise_subparts(doc, findings)
        _convert_safe_choice_equations(doc, findings)
        _convert_inline_slash_fractions(doc, findings)
        _convert_simple_math_paragraphs(doc, findings)
        _equation_and_fraction_audit(doc, findings)
        _audit_assertion_reason(doc, findings)
        _preserve_table_formatting(doc, findings)
        _bold_cms_tags(doc, findings)

    project_id = re.sub(r"_q$", "", options.project_question_prefix.strip(), flags=re.I).rstrip("_")
    if preserve_cms_mode:
        document_project_id, document_project_ids = _existing_project_id(doc)
        if document_project_id:
            project_id = document_project_id
            findings.append(Finding(0, "passed", f"Used project ID {project_id} from the document's existing Question id tags."))
        elif len(document_project_ids) > 1:
            findings.append(Finding(0, "manual_review", f"Existing Question id tags contain multiple project IDs: {', '.join(sorted(document_project_ids))}. The app used the fallback Project ID {project_id}."))
        else:
            findings.append(Finding(0, "passed", f"No project ID could be derived from existing Question id tags; used fallback Project ID {project_id}."))
    run_images_dir = images_root / run_id
    images_zip_path = packages / f"{run_id}_{project_id}_images.zip"
    image_result = process_document_images(doc, project_id, run_images_dir, images_zip_path)
    for warning in image_result.warnings:
        findings.append(Finding(0, "manual_review", warning))
    if image_result.artifacts:
        duplicate_count = sum(bool(item.duplicate_of) for item in image_result.artifacts)
        duplicate_note = f" Reused canonical filenames for {duplicate_count} exact duplicate occurrence(s)." if duplicate_count else ""
        findings.append(Finding(0, "fixed", f"Extracted {len(image_result.artifacts)} image occurrence(s), applied CMS filenames and wrote resolved filenames into Word Alt Text.{duplicate_note}"))
    else:
        findings.append(Finding(0, "passed", "No embedded images were detected."))

    suffix = {
        "images_only": "Images_Alt_Text_Ready",
        "images_math": "Images_Alt_Text_Math_Ready",
        "full": "CMS_Verification_Ready",
    }[options.processing_mode]
    output_name = f"{Path(safe_name).stem}_{suffix}.docx"
    output_path = processed / f"{run_id}_{output_name}"
    doc.core_properties.title = f"{Path(safe_name).stem} - CMS Verification Ready"
    doc.core_properties.subject = "Processed and audited for HeyMath CMS verification"
    doc.save(output_path)

    report_path = reports / f"{run_id}_{Path(safe_name).stem}_report.json"
    manual_review_count = sum(f.status == "manual_review" for f in findings)
    verification_status = "PENDING_MANUAL_REVIEW" if manual_review_count else "READY_FOR_VERIFICATION"
    report = {
        "processor_build_id": PROCESSOR_BUILD_ID,
        "run_id": run_id,
        "created_utc": datetime.now(timezone.utc).isoformat(),
        "original_filename": safe_name,
        "server_original_path": str(original_path),
        "server_output_path": str(output_path),
        "question_count": question_count,
        "image_count": len(image_result.artifacts),
        "image_manifest": str(image_result.manifest_json),
        "images_zip": str(images_zip_path),
        "fixed_count": sum(f.status == "fixed" for f in findings),
        "manual_review_count": manual_review_count,
        "verification_status": verification_status,
        "options": asdict(options),
        "effective_project_id": project_id,
        "findings": [asdict(f) for f in findings],
    }
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")

    manifest_path = queue / f"{run_id}.json"
    manifest = {
        "run_id": run_id,
        "status": verification_status,
        "created_utc": report["created_utc"],
        "original_filename": safe_name,
        "processed_document": str(output_path),
        "audit_report": str(report_path),
        "question_count": question_count,
        "image_count": len(image_result.artifacts),
        "image_manifest": str(image_result.manifest_json),
        "images_zip": str(images_zip_path),
        "manual_review_count": manual_review_count,
    }
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")

    return ProcessResult(
        output_path,
        report_path,
        manifest_path,
        image_result.manifest_json,
        images_zip_path,
        original_path,
        run_id,
        question_count,
        len(image_result.artifacts),
        findings,
    )


def process_docx_bytes(data: bytes, filename: str, options: ProcessorOptions, storage_dir: str | Path | None = None) -> ProcessResult:
    return process_docx(data, filename, options, storage_dir)
