from __future__ import annotations

import io
import csv
import json
import os
import re
import shutil
import uuid
import zipfile
from copy import deepcopy
from collections import Counter
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

from image_pipeline import ImagePipelineResult, process_document_images
from safe_math import parse as parse_safe_math, replace_span as replace_math_span


PROCESSOR_BUILD_ID = "2026.09.25-preserve-existing-snippets-v14.7"

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
    r"^@(Type|Question id|New snippet id|Snippet id|Difficulty level|Objective):\s*(.*?)\s*@$",
    re.I,
)
PLAIN_METADATA_RE = re.compile(
    r"^(Type|Question type|Question id|New snippet id|Snippet id|Difficulty(?: level)?|Objective):\s*(.*?)\s*$",
    re.I,
)
PLAIN_QUESTION_START_WITH_SUFFIX_RE = re.compile(
    r"^(?:Question|Q)\s*[:.\-]?\s*(\d+)\s*[).:]?\s+(.+?)\s*$",
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
LETTER_ANSWER_RE = re.compile(r"^\s*(?:Answer|Correct answer)\s*:?\s*\(?([A-Da-d])\)?\.?\s*$", re.I)
ANSWER_VALUE_RE = re.compile(r"^\s*(?:Answer|Correct answer)\s*:?\s*(.+?)\s*$", re.I)
MCQ_ANSWER_RE = re.compile(
    r"^\s*(?:Answer|Correct answer)\s*:?\s*(?:\(?([A-Da-d])\)?|([1-4]))(?:[.)])?(?:\s+(.+?))?\s*$",
    re.I,
)
MAPPING_LABEL_RE = re.compile(r"^(Curriculum|Taxonomy)\s*:\s*(.*?)\s*$", re.I)
SNIPPET_ID_RE = re.compile(r"^@?(?:New snippet id|Snippet id)\s*:\s*(\d+)\s*@?$", re.I)
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
    add_cms_tags: bool | None = None
    process_images: bool | None = None
    format_math: bool | None = None
    normalize_text_structure: bool | None = None
    prepare_mappings: bool | None = None
    format_math_structure: bool | None = None  # Legacy Version 13 switch.


@dataclass
class Finding:
    rule: int
    status: str
    message: str
    question: int | None = None
    paragraph: int | None = None


@dataclass(frozen=True)
class MappingInstruction:
    question: int
    snippet_id: int | None
    mapping_type: str
    path: str


@dataclass
class ProcessResult:
    output_path: Path
    report_path: Path
    manifest_path: Path
    image_manifest_path: Path
    images_zip_path: Path
    mapping_csv_path: Path
    original_path: Path
    run_id: str
    question_count: int
    image_count: int
    mapping_count: int
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
    if _has_embedded_content(paragraph):
        raise ValueError("Processing stopped to protect an equation or embedded object from a text-only rewrite. The original document is preserved.")
    paragraph.clear()
    paragraph.add_run(text)


def _has_embedded_content(paragraph: Paragraph) -> bool:
    return bool(paragraph._p.xpath(".//m:oMath | .//w:drawing | .//w:pict | .//w:object"))


def _equation_inventory(doc: _Document) -> Counter:
    # Keep structural operators as well as text: a radical and a fraction with
    # the same digits must never be treated as interchangeable.
    result = Counter()
    for equation in doc.element.xpath(".//m:oMath"):
        tokens = tuple(
            token
            for n in equation.iter(qn("m:t"))
            if (token := re.sub(r"\s+", "", n.text or ""))
        )
        if not any(tokens):
            continue  # Empty equation templates are intentionally removed.
        operators = tuple(n.tag for n in equation.iter() if n.tag in {
            qn("m:rad"), qn("m:f"), qn("m:sSup"), qn("m:sSub"), qn("m:sSubSup"), qn("m:nary")
        } and any((t.text or "").strip() for t in n.iter(qn("m:t"))))
        result[(tokens, operators)] += 1
    return result


def _replace_text_prefix(paragraph: Paragraph, length: int, replacement: str = "") -> None:
    """Edit only the ordinary-text prefix, preserving equations, objects and run properties."""
    for run in paragraph.runs:
        for node in run._r.findall(qn("w:t")):
            text = node.text or ""
            take = min(length, len(text))
            node.text = text[take:]
            length -= take
            if not length:
                break
        if not length:
            break
    if length:
        raise ValueError("Cannot safely edit this Word paragraph's prefix.")
    if replacement:
        run = paragraph.add_run(replacement)
        paragraph._p.remove(run._r)
        paragraph._p.insert(1 if paragraph._p.pPr is not None else 0, run._r)


def _is_bold_cms_tag(text: str) -> bool:
    """Match the standalone metadata and section tags that are bold in the CMS template."""
    stripped = text.strip()
    return bool(
        QUESTION_START_RE.fullmatch(stripped)
        or QUESTION_START_WITH_SUFFIX_RE.fullmatch(stripped)
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
    paragraphs = body_paragraphs(doc)
    for index, paragraph in enumerate(paragraphs):
        text = paragraph.text.strip()
        match = (
            QUESTION_START_RE.fullmatch(text)
            or QUESTION_START_WITH_SUFFIX_RE.fullmatch(text)
            or PLAIN_QUESTION_START_RE.fullmatch(text)
        )
        if match is None:
            suffix_match = PLAIN_QUESTION_START_WITH_SUFFIX_RE.fullmatch(text)
            next_paragraph = next((candidate for candidate in paragraphs[index + 1 :] if candidate.text.strip()), None)
            if suffix_match and next_paragraph is not None and _metadata_prefix_length(next_paragraph.text) > 0:
                match = suffix_match
        if match:
            starts.append((paragraph, int(match.group(1))))
    return starts


def _audit_source_question_numbering(
    doc: _Document,
    starts: list[tuple[Paragraph, int]],
    findings: list[Finding],
) -> None:
    """Block CMS-ready outputs when author question headings are not sequential."""
    if not starts:
        return
    occurrences: dict[int, list[int]] = {}
    source_numbers: list[int] = []
    for detected_position, (_paragraph, number) in enumerate(starts, 1):
        source_numbers.append(number)
        occurrences.setdefault(number, []).append(detected_position)

    errors = 0
    for number, positions in sorted(occurrences.items()):
        if len(positions) < 2:
            continue
        errors += 1
        locations = ", ".join(str(value) for value in positions)
        findings.append(
            Finding(
                15,
                "manual_review",
                f"Duplicate source question number {number} appears {len(positions)} times at detected question positions {locations}. Correct the source headings and reprocess; CMS-ready document, image and mapping downloads are blocked.",
                number,
            )
        )

    unique_numbers = sorted(occurrences)
    if unique_numbers:
        missing = [number for number in range(unique_numbers[0], unique_numbers[-1] + 1) if number not in occurrences]
        if missing:
            errors += 1
            findings.append(
                Finding(
                    15,
                    "manual_review",
                    f"Source question numbering has a gap. Missing question number(s): {', '.join(map(str, missing))}. Correct the source headings and reprocess; CMS-ready downloads are blocked.",
                )
            )

    if any(current < previous for previous, current in zip(source_numbers, source_numbers[1:])):
        errors += 1
        findings.append(
            Finding(
                15,
                "manual_review",
                f"Source question headings are out of order: {', '.join(map(str, source_numbers))}. Correct their order and reprocess; CMS-ready downloads are blocked.",
            )
        )

    if not errors:
        findings.append(Finding(15, "passed", "Source question headings are unique, sequential and in ascending order."))


def _canonicalise_section_labels(doc: _Document, findings: list[Finding]) -> None:
    for index, paragraph in enumerate(iter_paragraphs(doc), 1):
        raw = paragraph.text.strip()
        key = raw.casefold()
        if key in SECTION_ALIASES and raw != SECTION_ALIASES[key]:
            if _has_embedded_content(paragraph):
                if key in {"answer", "answer:", "answers", "answers:"}:
                    paragraph.insert_paragraph_before("@Answers:@")
                    _replace_text_prefix(paragraph, len(paragraph.text))
                continue
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
        for line in p.text.replace("\u00a0", " ").splitlines():
            match = KNOWN_METADATA_RE.fullmatch(line.strip()) or PLAIN_METADATA_RE.fullmatch(line.strip())
            field_name = match.group(1).casefold() if match else ""
            if field_name == "difficulty":
                field_name = "difficulty level"
            elif field_name == "question type":
                field_name = "type"
            elif field_name == "snippet id":
                field_name = "new snippet id"
            if match and field_name == name.casefold():
                return match.group(2).strip()
    return None


def document_has_complete_snippet_ids(data: bytes) -> bool:
    """Return true when every detected source question already has a positive snippet ID."""
    doc = Document(io.BytesIO(data))
    starts = _question_starts(doc)
    if not starts:
        return False
    for ordinal, (start, _) in enumerate(starts):
        next_start = starts[ordinal + 1][0] if ordinal + 1 < len(starts) else None
        value = _metadata_value(_block_paragraphs(start, next_start), "New snippet id")
        if not value or not value.isdigit() or int(value) <= 0:
            return False
    return True


def _contains_only_metadata(paragraph: Paragraph) -> bool:
    lines = [line.strip() for line in paragraph.text.replace("\u00a0", " ").splitlines() if line.strip()]
    return bool(lines) and all(
        KNOWN_METADATA_RE.fullmatch(line) or PLAIN_METADATA_RE.fullmatch(line)
        for line in lines
    )


def _metadata_prefix_length(text: str) -> int:
    """Return the character length of consecutive metadata lines at a paragraph's start."""
    consumed = 0
    found = False
    for line in text.replace("\u00a0", " ").splitlines(keepends=True):
        stripped = line.rstrip("\r\n").strip()
        if stripped and (KNOWN_METADATA_RE.fullmatch(stripped) or PLAIN_METADATA_RE.fullmatch(stripped)):
            consumed += len(line)
            found = True
            continue
        break
    return consumed if found else 0


def _remove_run_text_prefix(paragraph: Paragraph, length: int) -> None:
    """Remove a text prefix across Word runs, including manual line-break runs."""
    remaining = length
    for run in list(paragraph.runs):
        if remaining <= 0:
            break
        run_text = run.text
        if len(run_text) <= remaining:
            remaining -= len(run_text)
            run._r.getparent().remove(run._r)
        else:
            run.text = run_text[remaining:]
            remaining = 0
    if remaining:
        raise ValueError("Cannot safely separate metadata from question content in this Word paragraph.")


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
            findings.append(Finding(12, "manual_review", f"The Type tag says {declared_type}, but the structure indicates {expected_type} because the record contains {structure_description}. The selected processing sections preserved the tag unchanged.", question_number))

    if checked and not issues:
        findings.append(Finding(12, "passed", f"Question Type tags matched the detected Answers/Choices structure in {checked} record(s)."))


def _has_visible_content(paragraph: Paragraph) -> bool:
    return bool(paragraph.text.strip() or paragraph._p.xpath(".//w:drawing | .//w:pict | .//m:oMath"))


def _split_formatted_text(paragraph: Paragraph, spans: list[tuple[int, int]]) -> list[Paragraph]:
    """Clone plain-text runs for each slice, retaining their original properties."""
    original = [(deepcopy(r._r), r.text) for r in paragraph.runs]
    result = [paragraph]
    for _ in spans[1:]:
        result.append(_insert_after(result[-1]))
    for target, (start, end) in zip(result, spans):
        full_text = "".join(text for _, text in original)
        while start < end and full_text[start].isspace():
            start += 1
        while end > start and full_text[end - 1].isspace():
            end -= 1
        target.clear()
        offset = 0
        for node, text in original:
            left, right = max(0, start - offset), min(len(text), end - offset)
            if left < right:
                run = target.add_run(text[left:right])
                properties = node.find(qn("w:rPr"))
                if properties is not None:
                    run._r.insert(0, deepcopy(properties))
            offset += len(text)
    return result


def _split_compound_choices(paragraph: Paragraph) -> list[Paragraph] | None:
    if _has_embedded_content(paragraph):
        return None
    spans = [(m.start(), m.end()) for m in re.finditer(r"[^\r\n]+", paragraph.text) if m.group().strip()]
    if len(spans) != 4 or not all(OPTION_RE.fullmatch(paragraph.text[a:b]) for a, b in spans):
        return None
    return _split_formatted_text(paragraph, spans)


def _write_choice_marker(paragraph: Paragraph, option: int, correct: bool) -> None:
    suffix = " @correct answer@" if correct else ""
    match = re.match(r"^\s*(?:@[1-4]@|\(?[A-Da-d]\)?[.)])\s*", paragraph.text)
    _replace_text_prefix(paragraph, match.end() if match else 0, f"@{option}@ ")
    if correct:
        paragraph.add_run(suffix)


def _parse_mcq_answer_key(text: str) -> tuple[int, str | None] | None:
    match = MCQ_ANSWER_RE.fullmatch(text)
    if not match:
        return None
    letter, digit, supplied_text = match.groups()
    option = ord(letter.upper()) - ord("A") + 1 if letter else int(digit)
    return option, supplied_text.strip() if supplied_text and supplied_text.strip() else None


def _normalise_answer_text(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip().casefold()


def _infer_untagged_mcq_sections(doc: _Document, findings: list[Finding]) -> None:
    """Structure common untagged MCQs only when a letter answer and four choices are clear."""
    starts = _question_starts(doc)
    for ordinal, (start, detected_number) in enumerate(starts):
        next_start = starts[ordinal + 1][0] if ordinal + 1 < len(starts) else None
        block = _block_paragraphs(start, next_start)
        if any(p.text.strip() == "@Answers:@" for p in block):
            continue
        # A Question or Solution heading does not imply that choices are already structured.
        solution_index = next((i for i, p in enumerate(block) if p.text.strip() == "@Solution:@"), len(block))
        block = block[:solution_index]
        answer_candidates = [
            (i, parsed)
            for i, p in enumerate(block)
            if not _has_embedded_content(p) and (parsed := _parse_mcq_answer_key(p.text.strip())) is not None
        ]
        if len(answer_candidates) != 1:
            continue
        answer_index, (correct_option, supplied_answer_text) = answer_candidates[0]

        choice_paragraphs: list[Paragraph] | None = None
        for paragraph in block[1:answer_index]:
            compound = _split_compound_choices(paragraph)
            if compound is not None:
                choice_paragraphs = compound
                break
        if choice_paragraphs is None:
            choices_index = next((i for i, p in enumerate(block) if p.text.strip() == "@Choices:@"), None)
            candidates = [p for p in block[(choices_index + 1 if choices_index is not None else 1):answer_index]
                          if _has_visible_content(p) and p.text.strip() not in SECTION_MARKERS | {"@e@"}
                          and not (PLAIN_METADATA_RE.fullmatch(p.text.strip()) or KNOWN_METADATA_RE.fullmatch(p.text.strip()))]
            if (choices_index is not None and len(candidates) == 4) or (choices_index is None and len(candidates) >= 5):
                choice_paragraphs = candidates[-4:]
        if not choice_paragraphs or len(choice_paragraphs) != 4:
            findings.append(Finding(0, "manual_review", "A letter answer was found, but four MCQ choices could not be identified safely.", detected_number))
            continue

        if supplied_answer_text is not None:
            selected = OPTION_RE.fullmatch(choice_paragraphs[correct_option - 1].text.strip())
            selected_text = selected.group(3).strip() if selected else choice_paragraphs[correct_option - 1].text.strip()
            if _normalise_answer_text(selected_text) != _normalise_answer_text(supplied_answer_text):
                findings.append(
                    Finding(
                        0,
                        "manual_review",
                        f"The correct-answer label points to choice {correct_option}, but its supplied text does not match that choice. The record was preserved for review.",
                        detected_number,
                    )
                )
                continue

        if any("@correct answer@" in p.text.lower() for p in choice_paragraphs):
            findings.append(Finding(0, "manual_review", "Both a letter answer key and an existing correct-answer marker were found; check that they agree.", detected_number))
            continue

        if not any(p.text.strip() == "@Choices:@" for p in block):
            choice_paragraphs[0].insert_paragraph_before("@Choices:@")
        for option, paragraph in enumerate(choice_paragraphs, 1):
            _write_choice_marker(paragraph, option, option == correct_option)

        # Locate the answer paragraph again because splitting a compound choice can change the block.
        refreshed = _block_paragraphs(start, next_start)
        answer_paragraph = next((p for p in refreshed if _parse_mcq_answer_key(p.text.strip()) is not None), None)
        if answer_paragraph is not None:
            if not any(p.text.strip() == "@Solution:@" for p in refreshed):
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
        if any(p.text.strip() in {"@Answers:@", "@Choices:@"} for p in block):
            continue
        solution_index = next((i for i, p in enumerate(block) if p.text.strip() == "@Solution:@"), len(block))
        answer_block = block[:solution_index]
        declared_type = (_metadata_value(block, "Type") or "").upper()
        option_count = sum(bool(OPTION_RE.fullmatch(p.text.strip())) for p in answer_block)
        candidates: list[tuple[Paragraph, re.Match[str]]] = []
        for paragraph in answer_block:
            match = ANSWER_VALUE_RE.fullmatch(paragraph.text.strip())
            mcq_key = _parse_mcq_answer_key(paragraph.text.strip())
            if match and mcq_key and declared_type != "FIB" and option_count >= 4:
                continue
            if match and (_has_embedded_content(paragraph) or not LETTER_ANSWER_RE.fullmatch(paragraph.text.strip()) or declared_type == "FIB"):
                candidates.append((paragraph, match))
        if len(candidates) != 1:
            continue
        answer_paragraph, answer_match = candidates[0]
        answer_value = answer_match.group(1).strip()
        if not answer_value or len(answer_value) > 100 or "\n" in answer_value:
            findings.append(Finding(0, "manual_review", "An FIB-style answer was found but is too complex to structure automatically.", detected_number))
            continue
        answer_paragraph.insert_paragraph_before("@Answers:@")
        prefix = re.match(r"^\s*(?:Answer|Correct answer)\s*:?\s*", answer_paragraph.text, re.I)
        _replace_text_prefix(answer_paragraph, prefix.end())
        if solution_index == len(block):
            solution_marker = _insert_after(answer_paragraph, "@Solution:@")
            # Keep the explicit answer in @Answers:@, but do not invent solution prose.
            _insert_after(solution_marker, "@e@")
        findings.append(Finding(0, "fixed", "Identified an untagged FIB and created its Answers and Solution sections.", detected_number))


def _ensure_cms_records(doc: _Document, options: ProcessorOptions, findings: list[Finding]) -> int:
    _canonicalise_section_labels(doc, findings)
    _infer_untagged_mcq_sections(doc, findings)
    _infer_untagged_fib_sections(doc, findings)
    starts = _question_starts(doc)
    _audit_source_question_numbering(doc, starts, findings)
    if not starts:
        findings.append(Finding(0, "manual_review", "No question headings were detected. Use @Question: n@ or a standalone heading such as Question 1."))
        return 0

    # Work backwards so insertions do not disturb unprocessed record boundaries.
    for ordinal in range(len(starts) - 1, -1, -1):
        start, detected_number = starts[ordinal]
        next_start = starts[ordinal + 1][0] if ordinal + 1 < len(starts) else None
        block = _block_paragraphs(start, next_start)
        assigned_number = options.start_question_number + ordinal
        heading_match = PLAIN_QUESTION_START_RE.fullmatch(start.text.strip())
        suffix_match = QUESTION_START_WITH_SUFFIX_RE.fullmatch(start.text.strip())
        if suffix_match is None and heading_match is None:
            suffix_match = PLAIN_QUESTION_START_WITH_SUFFIX_RE.fullmatch(start.text.strip())
        heading_suffix = suffix_match.group(2).strip() if suffix_match else None
        heading_difficulty = DIFFICULTY_ALIASES.get((heading_match.group(2) or "").casefold()) if heading_match else None
        heading_objective = heading_match.group(3).title() if heading_match and heading_match.group(3) else None
        canonical_heading = f"@Question: {assigned_number}@" + (f" {heading_suffix}" if heading_suffix else "")
        _set_text(start, canonical_heading)

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
            if _contains_only_metadata(paragraph):
                _remove_paragraph(paragraph)
                continue
            metadata_prefix = _metadata_prefix_length(paragraph.text)
            if metadata_prefix:
                _remove_run_text_prefix(paragraph, metadata_prefix)
                findings.append(Finding(0, "fixed", "Separated leading metadata lines from question content stored in the same Word paragraph.", assigned_number))

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
            findings.append(Finding(0, "fixed", f"Retained question-heading text after the canonical tag: {heading_suffix!r}.", assigned_number))

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
            marker = _insert_after(block[-1], "@Solution:@")
            _insert_after(marker, "@e@")
            findings.append(Finding(1, "manual_review", "Added an empty Solution section because no solution was supplied. No solution content was invented.", assigned_number))

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


def _audit_unsectioned_record_text(doc: _Document, findings: list[Finding]) -> None:
    """Flag preserved text that sits between CMS sections within a question record."""
    context = _section_for_paragraphs(doc)
    for index, paragraph in enumerate(body_paragraphs(doc), 1):
        section, question = context.get(paragraph._p, (None, None))
        text = paragraph.text.strip()
        if question is None or section is not None or not text:
            continue
        if (
            QUESTION_START_RE.fullmatch(text)
            or QUESTION_START_WITH_SUFFIX_RE.fullmatch(text)
            or PLAIN_QUESTION_START_RE.fullmatch(text)
            or PLAIN_QUESTION_START_WITH_SUFFIX_RE.fullmatch(text)
            or KNOWN_METADATA_RE.fullmatch(text)
            or PLAIN_METADATA_RE.fullmatch(text)
            or text in SECTION_MARKERS
            or text == "@e@"
            or MAPPING_LABEL_RE.fullmatch(text)
            or ">>" in text
        ):
            continue
        findings.append(
            Finding(
                8,
                "manual_review",
                f"Text occurs outside a recognised Question, Answers/Choices or Solution section and was preserved in place: {text!r}",
                question,
                index,
            )
        )


def _section_for_paragraphs(doc: _Document) -> dict[object, tuple[str | None, int | None]]:
    state: dict[object, tuple[str | None, int | None]] = {}
    question_numbers = {paragraph._p: number for paragraph, number in _question_starts(doc)}
    section: str | None = None
    question: int | None = None
    for paragraph in body_paragraphs(doc):
        text = paragraph.text.strip()
        if paragraph._p in question_numbers:
            question = question_numbers[paragraph._p]
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
        if section not in {"@Question:@", "@Choices:@", "@Solution:@"}:
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
                if re.sub(r"@[1-4]@|@correct answer@", "", run.text).strip():
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
            label = section.strip("@:")
            findings.append(Finding(1, "fixed", f"Removed whole-block bold formatting from the {label.lower()} while preserving the text.", question))
        elif any(bold_flags):
            emphasis_sections += 1
            label = section.strip("@:").lower()
            findings.append(Finding(1, "passed", f"Preserved selective bold emphasis within the {label}; the entire block was not bold.", question))

    if not changed_sections and not emphasis_sections:
        findings.append(Finding(1, "passed", "No whole-question, whole-choices or whole-solution bold formatting was found."))
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
    text = re.sub(r"(?<=\d),\s+(?=\d)", ",", text)
    text = re.sub(r"\s*([=+×÷−])\s*", r" \1 ", text)
    text = re.sub(r" +([.;:?!])", r"\1", text)
    return text


def _convert_rupee_symbols(doc: _Document, findings: list[Finding]) -> None:
    changes = 0
    for node in doc.element.xpath(".//w:t | .//m:t"):
        original = node.text or ""
        updated, count = re.subn(r"₹\s*(?=\d)", "Rs ", original)
        if count:
            node.text = updated
            changes += count
    findings.append(
        Finding(
            13,
            "fixed" if changes else "passed",
            f"Converted {changes} rupee symbol occurrence(s) to Rs followed by one space."
            if changes
            else "No rupee symbols before numeric amounts required conversion.",
        )
    )


def _normalise_paragraph_run_boundaries(paragraph: Paragraph) -> int:
    """Normalise whitespace across adjacent ordinary Word runs without flattening formatting."""
    groups: list[list[object]] = []
    current: list[object] = []
    for child in paragraph._p:
        if child.tag == qn("w:pPr"):
            continue
        if child.tag != qn("w:r"):
            if current:
                groups.append(current)
                current = []
            continue
        nodes = child.findall(qn("w:t"))
        if not nodes:
            if current:
                groups.append(current)
                current = []
            continue
        current.extend(nodes)
    if current:
        groups.append(current)

    changes = 0
    for nodes in groups:
        previous = None
        for node in nodes:
            if previous is None:
                previous = node
                continue
            left = previous.text or ""
            right = node.text or ""
            right_visible = right.lstrip()
            if right_visible.startswith(("?", ".", ",", ":", ";", "!")):
                new_left = left.rstrip()
                new_right = right_visible
            elif left[-1:].isspace() and right[:1].isspace():
                new_left = left.rstrip() + " "
                new_right = right.lstrip()
            else:
                previous = node
                continue
            if new_left != left:
                previous.text = new_left
                changes += 1
            if new_right != right:
                node.text = new_right
                changes += 1
            previous = node
    return changes


def _repair_spacing(doc: _Document, findings: list[Finding]) -> None:
    changes = 0
    for node in doc.element.xpath(".//w:t | .//m:t"):
        original = node.text or ""
        updated = _normalise_spacing(original)
        if updated != original:
            node.text = updated
            changes += 1
    boundary_changes = sum(_normalise_paragraph_run_boundaries(paragraph) for paragraph in iter_paragraphs(doc))
    changes += boundary_changes
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
        if any(child.tag not in {qn("w:rPr"), qn("w:t")} for child in run._r):
            continue
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


GEOMETRY_SINGLE_CUE_RE = re.compile(r"\b(?i:point|centre|center|vertex)\s+([A-Z])\b")
GEOMETRY_LABEL_CUE_RE = re.compile(
    r"\b(?i:line segment|line|segment|ray|chord|arc|angle|triangle|circle)\s+([A-Z]{1,3})\b",
)
GEOMETRY_SYMBOL_RE = re.compile(r"(?:∠|△|Δ)\s*([A-Z]{1,3})\b")
VARIABLE_CUE_RE = re.compile(
    r"\b(?i:let|variable|value\s+of|solve\s+for|find|where|radius|diameter|length|breadth|height|distance|speed|time|mass)\s+([A-Za-z])\b",
)
PROTECTED_LABEL_RE = re.compile(r"\b(?i:option|choice|part|class)\s+([A-Da-d])\b")
ASSERTION_REASON_LABEL_RE = re.compile(r"\b(?i:Assertion|Reason)\s*\(\s*([AR])\s*\)")
UNIT_TOKENS = {"mm", "cm", "m", "km", "mg", "g", "kg", "ml", "l", "s", "min", "h"}


def _add_match_group_span(spans: list[tuple[int, int]], match: re.Match[str], group: int = 1) -> str:
    spans.append(match.span(group))
    return match.group(group)


def _measurement_unit_ranges(text: str, paragraph: Paragraph) -> list[tuple[int, int]]:
    ranges: list[tuple[int, int]] = []
    unit_pattern = r"(?:mm|cm|km|mg|kg|mL|min|m|g|L|l|s|h)(?:²|³)?"
    for match in re.finditer(rf"(?<![A-Za-z])\d+(?:\.\d+)?\s+({unit_pattern})\b", text):
        ranges.append(match.span(1))
    for match in re.finditer(rf"\b({unit_pattern})\s*/\s*({unit_pattern})\b", text):
        ranges.extend((match.span(1), match.span(2)))
    # When a native equation is followed by a unit, paragraph.text contains only
    # the ordinary-text suffix. Treat a leading unit as measurement text.
    if paragraph._p.xpath(".//m:oMath"):
        match = re.match(rf"\s*({unit_pattern})\b", text)
        if match:
            ranges.append(match.span(1))
    return ranges


def _range_contains(ranges: list[tuple[int, int]], start: int, end: int) -> bool:
    return any(left <= start and end <= right for left, right in ranges)


def _looks_like_article(text: str, start: int, end: int) -> bool:
    if text[start:end].casefold() != "a":
        return False
    before = text[:start].rstrip()
    after = text[end:].lstrip()
    next_word = re.match(r"([A-Za-z]+)", after)
    if not next_word:
        return False
    sentence_start = not before or before[-1:] in ".?!:;"
    preceding_word = re.search(r"([A-Za-z]+)\s*$", before)
    ordinary_article_position = sentence_start or not preceding_word or preceding_word.group(1).casefold() in {
        "of", "in", "on", "with", "from", "by", "for", "as", "draw", "construct"
    }
    relation_verbs = {"is", "are", "lies", "meets", "intersects", "equals", "represents", "denotes", "has"}
    return ordinary_article_position and next_word.group(1).casefold() not in relation_verbs


def _italicise_absolute_spans(paragraph: Paragraph, spans: list[tuple[int, int]]) -> int:
    spans = sorted({span for span in spans if span[0] < span[1]})
    if not spans:
        return 0
    changed = 0
    offset = 0
    for source in list(paragraph.runs):
        text = source.text
        run_start, run_end = offset, offset + len(text)
        offset = run_end
        intersections = [
            (max(left, run_start) - run_start, min(right, run_end) - run_start)
            for left, right in spans
            if left < run_end and right > run_start
        ]
        if not intersections:
            continue
        boundaries = {0, len(text)}
        for left, right in intersections:
            boundaries.update((left, right))
        points = sorted(boundaries)
        source_properties = source._r.find(qn("w:rPr"))
        for left, right in zip(points, points[1:]):
            if left == right:
                continue
            new_run = paragraph.add_run(text[left:right])
            if source_properties is not None:
                existing = new_run._r.find(qn("w:rPr"))
                if existing is not None:
                    new_run._r.remove(existing)
                new_run._r.insert(0, deepcopy(source_properties))
            if any(span_left <= left and right <= span_right for span_left, span_right in intersections):
                if new_run.italic is not True:
                    new_run.italic = True
                    changed += 1
            source._r.addprevious(new_run._r)
        source._r.getparent().remove(source._r)
    return changed


def _repair_italics(doc: _Document, findings: list[Finding]) -> None:
    context = _section_for_paragraphs(doc)
    declared: dict[int, set[str]] = {}
    direct_spans: dict[object, list[tuple[int, int]]] = {}

    for paragraph in body_paragraphs(doc):
        section, question = context.get(paragraph._p, (None, None))
        if question is None or section not in {"@Question:@", "@Choices:@", "@Answers:@", "@Solution:@"}:
            continue
        symbols = declared.setdefault(question, set())
        for node in paragraph._p.xpath(".//m:oMath//m:t"):
            value = (node.text or "").strip()
            if re.fullmatch(r"[A-Za-z]", value):
                symbols.add(value)
        if paragraph._p.xpath(".//w:drawing | .//w:pict | .//w:object"):
            continue
        text = paragraph.text
        spans: list[tuple[int, int]] = []
        for pattern in (GEOMETRY_SINGLE_CUE_RE, GEOMETRY_LABEL_CUE_RE, GEOMETRY_SYMBOL_RE, VARIABLE_CUE_RE):
            for match in pattern.finditer(text):
                symbols.add(_add_match_group_span(spans, match))
        # A single letter touching a mathematical operator, digit or explicit
        # superscript is a high-confidence variable even when prose separates
        # it from cues such as "the equation".
        protected_labels = [match.span(1) for match in ASSERTION_REASON_LABEL_RE.finditer(text)]
        for match in re.finditer(r"(?<![A-Za-z])([A-Za-z])(?![A-Za-z])", text):
            start, end = match.span(1)
            if _range_contains(protected_labels, start, end):
                continue
            left = text[start - 1:start]
            right = text[end:end + 1]
            math_adjacent = (
                bool(left) and left in "=+×÷−-<>≤≥/(²³"
                or bool(right) and right in "=+×÷−-<>≤≥/)²³"
                or left.isdigit()
                or right.isdigit()
            )
            if math_adjacent:
                symbols.add(_add_match_group_span(spans, match))
        for match in re.finditer(r"\b([A-Z])\s+and\s+([A-Z])\s+are\s+(?:two\s+)?(?:fixed\s+)?points\b", text):
            symbols.add(_add_match_group_span(spans, match, 1))
            symbols.add(_add_match_group_span(spans, match, 2))
        for match in re.finditer(r"\b([A-Z]{1,3})\s+(?i:is\s+(?:a|an|the)\s+(?:diameter|radius|chord|line|segment|ray|angle|triangle))\b", text):
            symbols.add(_add_match_group_span(spans, match))
        direct_spans[paragraph._p] = spans

    changed = 0
    uncertain_by_question: dict[int, set[str]] = {}
    for paragraph in body_paragraphs(doc):
        section, question = context.get(paragraph._p, (None, None))
        if question is None or section not in {"@Question:@", "@Choices:@", "@Answers:@", "@Solution:@"}:
            continue
        if paragraph._p.xpath(".//w:drawing | .//w:pict | .//w:object"):
            continue
        if any(child.tag not in {qn("w:pPr"), qn("w:r"), qn("m:oMath")} for child in paragraph._p):
            continue
        text = paragraph.text
        if text.strip() in SECTION_MARKERS | {"@e@"}:
            continue
        spans = list(direct_spans.get(paragraph._p, []))
        protected = [match.span(1) for match in PROTECTED_LABEL_RE.finditer(text)]
        protected.extend(match.span(1) for match in ASSERTION_REASON_LABEL_RE.finditer(text))
        protected.extend(_measurement_unit_ranges(text, paragraph))
        choice_label = re.match(r"^\s*([a-dA-D])(?:[.)]|\s*@)", text)
        if choice_label:
            protected.append(choice_label.span(1))

        symbols = declared.get(question, set())
        for symbol in sorted(symbols, key=len, reverse=True):
            for match in re.finditer(rf"(?<![A-Za-z]){re.escape(symbol)}(?![A-Za-z])", text):
                start, end = match.span()
                if _range_contains(protected, start, end) or _looks_like_article(text, start, end):
                    continue
                spans.append((start, end))

        geometry_context = bool(re.search(r"\b(?:point|line|segment|ray|angle|triangle|centre|center|vertex|chord|arc|circle)\b|[∠△Δ]", text, re.I))
        if geometry_context:
            for match in re.finditer(r"(?<![A-Za-z])([A-Z])(?![A-Za-z])", text):
                start, end = match.span(1)
                if any(left <= start and end <= right for left, right in spans + protected):
                    continue
                if _looks_like_article(text, start, end):
                    continue
                uncertain_by_question.setdefault(question, set()).add(match.group(1))

        changed += _italicise_absolute_spans(paragraph, spans)

    if changed:
        findings.append(Finding(3, "fixed", f"Italicised {changed} high-confidence variable or geometry-label occurrence(s) while protecting recognised units and ordinary labels."))
    else:
        findings.append(Finding(3, "passed", "No high-confidence ordinary-text variable or geometry-label occurrences required italic formatting."))
    for question, symbols in sorted(uncertain_by_question.items()):
        findings.append(
            Finding(
                3,
                "manual_review",
                f"Possible geometry label(s) {', '.join(sorted(symbols))} could not be distinguished safely from ordinary text and were left unchanged.",
                question,
            )
        )


def _equation_and_fraction_audit(doc: _Document, findings: list[Finding]) -> None:
    native_count = len(doc.element.xpath(".//m:oMath"))
    findings.append(Finding(2, "passed", f"Detected {native_count} native Word equation object(s)."))
    context = _section_for_paragraphs(doc)
    fraction_hits = 0
    image_hits = 0
    ambiguous_scope_hits = 0
    for index, paragraph in enumerate(body_paragraphs(doc), 1):
        section, question = context.get(paragraph._p, (None, None))
        if section not in {"@Question:@", "@Solution:@", "@Answers:@", "@Choices:@"}:
            continue
        text = paragraph.text
        # A radical or caret expression left in ordinary text was rejected by
        # the conservative parser and requires an author to confirm its scope.
        # Unicode superscripts such as 7³ and x² have an explicit single base
        # and are valid ordinary Word text, so they are not warnings.
        ambiguous_scope = "√" in text or "^" in text
        if SIMPLE_FRACTION_RE.search(text) and not ambiguous_scope:
            fraction_hits += 1
            findings.append(Finding(7, "manual_review", f"Slash-style fraction requires conversion to a stacked Word equation: {text.strip()!r}", question, index))
        if paragraph._p.xpath(".//w:drawing | .//w:pict"):
            image_hits += 1
            findings.append(Finding(2, "manual_review", "An embedded image occurs in mathematical content. Verify that it is a diagram, not an equation screenshot.", question, index))
        if ambiguous_scope:
            ambiguous_scope_hits += 1
            findings.append(
                Finding(
                    2,
                    "manual_review",
                    f"A radical or caret expression has potentially ambiguous scope and was left unchanged: {text.strip()!r}",
                    question,
                    index,
                )
            )
    if not fraction_hits:
        findings.append(Finding(7, "passed", "No slash-style fractions were detected in student-facing content."))
    if not image_hits and not ambiguous_scope_hits:
        findings.append(Finding(2, "passed", "No likely equation screenshots or genuinely ambiguous radical/caret expressions were detected."))


def _math_run(text: str) -> OxmlElement:
    run = OxmlElement("m:r")
    value = OxmlElement("m:t")
    if text[:1].isspace() or text[-1:].isspace():
        value.set(qn("xml:space"), "preserve")
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


def _inline_radical_spans(text: str) -> list[tuple[int, int]]:
    """Find self-contained radicals whose scope is explicit in ordinary text."""
    spans: list[tuple[int, int]] = []
    for marker in re.finditer("√", text):
        start = marker.start()
        if text[:start].rstrip().endswith("/"):
            continue
        cursor = marker.end()
        if cursor >= len(text):
            continue
        if text[cursor] == "(":
            depth = 0
            end = None
            for index in range(cursor, len(text)):
                if text[index] == "(":
                    depth += 1
                elif text[index] == ")":
                    depth -= 1
                    if depth == 0:
                        end = index + 1
                        break
            if end is not None:
                spans.append((start, end))
            continue
        number = re.match(r"\d+(?:\.\d+)?", text[cursor:])
        if not number:
            continue
        end = cursor + number.end()
        if text[end:end + 1] in {"/", "^", "²", "³"}:
            continue
        spans.append((start, end))
    return spans


def _semantic_math_text(paragraph: Paragraph) -> str:
    """Expose Word superscript 2/3 formatting to the plain-text math parser.

    Replacing a one-character run with its Unicode superscript keeps every
    offset aligned with paragraph.runs, so surgical OOXML replacement remains
    safe while expressions such as r + superscript 3 parse as r³.
    """
    parts: list[str] = []
    superscripts = str.maketrans({"2": "²", "3": "³"})
    for run in paragraph.runs:
        text = run.text
        if run.font.superscript is True:
            text = text.translate(superscripts)
        parts.append(text)
    return "".join(parts)


def _inline_equality_spans(text: str) -> list[tuple[int, int]]:
    """Find self-contained equalities embedded in prose."""
    operand = r"(?:\d+(?:\.\d+)?|(?:\d+(?:\.\d+)?)?[A-Za-z](?:[²³])?)"
    operator_tail = rf"(?:\s*[+−\-×*÷]\s*{operand})*"
    pattern = re.compile(
        rf"(?<![A-Za-z0-9]){operand}{operator_tail}\s*=\s*{operand}{operator_tail}(?![A-Za-z0-9])"
    )
    return [match.span() for match in pattern.finditer(text)]


def _replace_inline_math_with_spacing(
    paragraph: Paragraph,
    text: str,
    start: int,
    end: int,
    equation: OxmlElement,
) -> None:
    """Replace an inline expression and preserve one visible prose boundary space."""
    # Keep the author's surrounding prose spaces outside the OMML object.
    # Word suppresses ordinary spaces stored inside an inline equation, even
    # when m:t carries xml:space="preserve".  The boundary pass below turns
    # these prose spaces into visible non-breaking spaces.
    replace_math_span(paragraph, start, end, equation)


VISIBLE_EQUATION_SPACE = "\u2002"  # En space: Word renders it reliably beside OMML objects.


def _normalise_native_equation_boundaries(doc: _Document, findings: list[Finding]) -> None:
    """Keep one visible space where an inline equation meets surrounding prose."""
    changed = 0
    for paragraph in body_paragraphs(doc):
        children = list(paragraph._p)
        for index, equation in enumerate(children):
            if equation.tag != qn("m:oMath"):
                continue
            previous = children[index - 1] if index > 0 and children[index - 1].tag == qn("w:r") else None
            following = children[index + 1] if index + 1 < len(children) and children[index + 1].tag == qn("w:r") else None
            previous_text_nodes = list(previous.iter(qn("w:t"))) if previous is not None else []
            following_text_nodes = list(following.iter(qn("w:t"))) if following is not None else []

            if previous_text_nodes:
                previous_value = previous_text_nodes[-1].text or ""
                prefix = "".join((node.text or "") for child in children[:index] for node in child.iter(qn("w:t")))
                label_prefix = bool(re.search(r"(?:@[1-4]@|\([a-zivx]+\))\s*$", prefix, re.I))
                needs_space = bool(previous_value) and (previous_value[-1].isspace() or previous_value[-1].isalnum())
                if needs_space and not label_prefix:
                    visible_value = previous_value.rstrip(" \u00a0\u2002\t") + VISIBLE_EQUATION_SPACE
                    if visible_value != previous_value:
                        previous_text_nodes[-1].text = visible_value
                        previous_text_nodes[-1].set(qn("xml:space"), "preserve")
                        changed += 1

            if following_text_nodes:
                following_value = following_text_nodes[0].text or ""
                needs_space = bool(following_value) and (following_value[0].isspace() or following_value[0].isalnum())
                if needs_space:
                    visible_value = VISIBLE_EQUATION_SPACE + following_value.lstrip(" \u00a0\u2002\t")
                    if visible_value != following_value:
                        following_text_nodes[0].text = visible_value
                        following_text_nodes[0].set(qn("xml:space"), "preserve")
                        changed += 1
    if changed:
        findings.append(Finding(4, "fixed", f"Normalised {changed} prose-to-equation boundary space(s)."))


def _overlaps(span: tuple[int, int], others: list[tuple[int, int]]) -> bool:
    return any(span[0] < right and span[1] > left for left, right in others)


def _convert_explicit_math(doc: _Document, findings: list[Finding]) -> None:
    context = _section_for_paragraphs(doc)
    atom = r"(?:\d+(?:\.\d+)?|[A-Za-z](?![A-Za-z])|[²³=+−\-×*÷/()√^])"
    assignment = re.compile(r"\b[A-Za-z]\s*=\s*" + atom + r"(?:\s*" + atom + r")*")
    fraction = re.compile(r"(?<![\w/])\d+\s*/\s*\d+(?![\w/])")
    for paragraph in body_paragraphs(doc):
        section, question = context.get(paragraph._p, (None, None))
        if section not in {"@Question:@", "@Choices:@", "@Answers:@", "@Solution:@"}:
            continue
        # Existing equations may coexist with convertible ordinary text. Drawings,
        # objects and hyperlinks still require a human-safe edit.
        if paragraph._p.xpath(".//w:drawing | .//w:pict | .//w:object"):
            continue
        if any(child.tag not in {qn('w:pPr'), qn('w:r'), qn('m:oMath')} for child in paragraph._p) or any(
            child.tag not in {qn('w:rPr'), qn('w:t')} for r in paragraph.runs for child in r._r
        ):
            continue
        text = _semantic_math_text(paragraph)
        if text.strip() in SECTION_MARKERS | {"@e@"}:
            continue
        spans = []
        choice = TAGGED_CHOICE_RE.fullmatch(text)
        if choice:
            expression = choice.group(2).strip()
            expression = re.sub(r"\s+(?:units?|litres?|liters?|cm|mm|km|m|kg|g|mL|L)\s*$", "", expression, flags=re.I)
            start = text.find(expression, len(choice.group(1)))
            if re.search(r"[²³√=+−×÷/^]|\d\s*\(", expression):
                try:
                    parse_safe_math(expression)
                    spans = [(start, start + len(expression))]
                except ValueError:
                    spans = [(start + left, start + right) for left, right in _inline_radical_spans(expression)]
                    for match in fraction.finditer(expression):
                        candidate = (start + match.start(), start + match.end())
                        if not _overlaps(candidate, spans):
                            spans.append(candidate)
        else:
            spans = [m.span() for m in assignment.finditer(text)]
            if not spans and "=" not in text and re.fullmatch(r"\s*(?:\([a-z]\)\s*)?[\dA-Za-z²³√=+−\-×*÷/()^ .]+\s*", text) and not re.search(r"[A-Za-z]{2}", text) and re.search(r"[²³√=+−×÷/]", text):
                start = re.match(r"\s*(?:\([a-z]\)\s*)?", text).end()
                spans = [(start, len(text.rstrip()))]
            for radical_span in _inline_radical_spans(text):
                if not _overlaps(radical_span, spans):
                    spans.append(radical_span)
            for match in fraction.finditer(text):
                fraction_span = match.span()
                if not _overlaps(fraction_span, spans):
                    spans.append(fraction_span)
        for start, end in reversed(spans):
            expression = text[start:end].strip()
            # A literal space after a simple slash fraction ends the denominator.
            # Without a space, an attached variable or bracket remains ambiguous.
            immediate_after = text[end:end + 1]
            if '/' in expression and (
                re.search(r"[/\w]", text[max(0, start - 1):start])
                or text[:start].rstrip().endswith("√")
                or immediate_after in {"/", "(", "^", "²", "³"}
                or bool(immediate_after and immediate_after.isalpha())
            ):
                findings.append(Finding(2, "manual_review", f"Preserved ambiguous fraction context: {expression!r}. Use Word Equation Editor to specify its meaning.", question))
                continue
            try:
                equation = parse_safe_math(expression)
                _replace_inline_math_with_spacing(paragraph, text, start, end, equation)
            except ValueError:
                findings.append(Finding(2, "manual_review", f"Preserved uncertain mathematical expression: {expression!r}. Use explicit parentheses or Word Equation Editor.", question))
                continue
            findings.append(Finding(7 if '/' in expression else 2, "fixed", f"Converted explicit expression to a native Word equation with italic variables: {expression!r}", question))


def _convert_safe_choice_equations(doc: _Document, findings: list[Finding]) -> None:
    """Convert only unambiguous tagged-choice assignments to native Word equations."""
    context = _section_for_paragraphs(doc)
    converted = 0
    for paragraph in body_paragraphs(doc):
        if any(r.italic is not None or r.bold is True for r in paragraph.runs):
            continue
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
        if any(r.italic is not None or r.bold is True for r in paragraph.runs):
            continue
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
        if any(r.italic is not None or r.bold is True for r in paragraph.runs):
            continue
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
        if _has_embedded_content(paragraph):
            findings.append(Finding(8, "manual_review", "Combined answers contain equations or objects; preserved intact. Put each labelled answer in its own paragraph.", question))
            continue
        # Match against the original string so offsets include leading whitespace.
        matches = list(re.finditer(r"(?<!\w)(\([a-z]\)|[a-z]\))\s*", paragraph.text, re.I))
        spans = [(m.start(), matches[i + 1].start() if i + 1 < len(matches) else len(paragraph.text)) for i, m in enumerate(matches)]
        parts = _split_formatted_text(paragraph, spans)
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
            _replace_text_prefix(paragraph, match.end(), f"({label}) ")
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
        findings.append(Finding(10, "passed", "Preserved author-supplied Assertion–Reason choice wording and formatting.", number))
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
                _write_choice_marker(paragraph, option, False)
                changed += 1
    if changed:
        findings.append(Finding(0, "fixed", f"Normalised {changed} MCQ option marker(s)."))


def _normalise_mapping_path(raw: str) -> str | None:
    parts = [re.sub(r"\s+", " ", part).strip() for part in raw.split(">>")]
    if len(parts) < 2 or any(not part for part in parts):
        return None
    return " >> ".join(parts)


def _extract_mapping_instructions(doc: _Document, findings: list[Finding]) -> list[MappingInstruction]:
    paragraphs = body_paragraphs(doc)
    question_numbers = {paragraph._p: number for paragraph, number in _question_starts(doc)}
    snippet_ids: dict[int, int] = {}
    current_question: int | None = None
    for paragraph in paragraphs:
        text = paragraph.text.strip()
        if paragraph._p in question_numbers:
            current_question = question_numbers[paragraph._p]
            continue
        snippet_match = SNIPPET_ID_RE.fullmatch(text)
        if current_question is not None and snippet_match:
            snippet_ids[current_question] = int(snippet_match.group(1))

    instructions: list[MappingInstruction] = []
    seen: set[tuple[int, str, str]] = set()
    remove: list[Paragraph] = []
    active_type: str | None = None
    active_label: Paragraph | None = None
    mapping_labels: list[tuple[Paragraph, int | None, str]] = []
    label_path_counts: dict[int, int] = {}
    current_question = None
    questions_with_mappings: set[int] = set()

    def add_path(question: int | None, mapping_type: str, raw_path: str, paragraph: Paragraph) -> bool:
        if question is None:
            findings.append(Finding(14, "manual_review", f"A {mapping_type} mapping occurs before an identifiable question: {raw_path!r}."))
            return False
        path = _normalise_mapping_path(raw_path)
        if path is None:
            findings.append(Finding(14, "manual_review", f"The {mapping_type} path is incomplete or malformed: {raw_path!r}.", question))
            return False
        key = (question, mapping_type, path.casefold())
        if key not in seen:
            seen.add(key)
            instructions.append(MappingInstruction(question, snippet_ids.get(question), mapping_type, path))
            questions_with_mappings.add(question)
        remove.append(paragraph)
        return True

    for paragraph in paragraphs:
        text = paragraph.text.strip()
        if paragraph._p in question_numbers:
            current_question = question_numbers[paragraph._p]
            active_type = None
            active_label = None
            continue
        label_match = MAPPING_LABEL_RE.fullmatch(text)
        if label_match:
            active_type = label_match.group(1).title()
            active_label = paragraph
            mapping_labels.append((paragraph, current_question, active_type))
            label_path_counts[id(paragraph._p)] = 0
            remove.append(paragraph)
            inline_path = label_match.group(2).strip()
            if inline_path:
                if add_path(current_question, active_type, inline_path, paragraph):
                    label_path_counts[id(paragraph._p)] += 1
            continue
        if active_type and ">>" in text and not text.startswith("@"):
            if add_path(current_question, active_type, text, paragraph) and active_label is not None:
                label_path_counts[id(active_label._p)] += 1
            continue
        if text:
            active_type = None
            active_label = None

    for label, question, mapping_type in mapping_labels:
        if label_path_counts[id(label._p)] == 0:
            findings.append(Finding(14, "manual_review", f"The {mapping_type} mapping block has no valid path.", question))

    removed_nodes: set[int] = set()
    for paragraph in remove:
        node_id = id(paragraph._p)
        if node_id not in removed_nodes:
            removed_nodes.add(node_id)
            _remove_paragraph(paragraph)

    missing_snippets = sorted(question for question in questions_with_mappings if question not in snippet_ids)
    for question in missing_snippets:
        findings.append(Finding(14, "manual_review", "Mapping paths were prepared, but the question has no usable New snippet id. Add the snippet ID before applying mappings to CMS.", question))
    if instructions:
        findings.append(Finding(14, "fixed", f"Prepared {len(instructions)} curriculum/taxonomy mapping path(s) for {len(questions_with_mappings)} question(s) and removed the author-only mapping lines from the CMS-ready document."))
    else:
        findings.append(Finding(14, "passed", "No Curriculum or Taxonomy mapping instructions were found."))
    return instructions


def _write_mapping_csv(path: Path, instructions: list[MappingInstruction]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.DictWriter(handle, fieldnames=["Question", "Snippet ID", "Mapping Type", "Path"])
        writer.writeheader()
        for item in instructions:
            writer.writerow(
                {
                    "Question": item.question,
                    "Snippet ID": item.snippet_id or "",
                    "Mapping Type": item.mapping_type,
                    "Path": item.path,
                }
            )


def _selected_processing_sections(options: ProcessorOptions) -> tuple[bool, bool, bool, bool, bool]:
    """Resolve the new independent switches while retaining old API callers."""
    if any(
        value is not None
        for value in (
            options.add_cms_tags,
            options.process_images,
            options.format_math,
            options.normalize_text_structure,
            options.prepare_mappings,
            options.format_math_structure,
        )
    ):
        legacy_formatting = bool(options.format_math_structure)
        return (
            bool(options.add_cms_tags),
            bool(options.process_images),
            bool(options.format_math) if options.format_math is not None else legacy_formatting,
            bool(options.normalize_text_structure) if options.normalize_text_structure is not None else legacy_formatting,
            bool(options.prepare_mappings),
        )
    legacy = {
        "full": (True, True, True, True, False),
        "images_only": (False, True, False, False, False),
        "images_math": (False, True, True, True, False),
    }
    try:
        return legacy[options.processing_mode]
    except KeyError as exc:
        raise ValueError(f"Unsupported processing mode: {options.processing_mode!r}") from exc


def _output_suffix(
    add_cms_tags: bool,
    process_images: bool,
    format_math: bool,
    normalize_text_structure: bool,
    prepare_mappings: bool,
) -> str:
    selected = (add_cms_tags, process_images, format_math, normalize_text_structure)
    if not any(selected) and not prepare_mappings:
        raise ValueError("Select at least one processing section.")
    if selected == (True, True, True, True):
        return "CMS_Verification_Ready"
    parts: list[str] = []
    if add_cms_tags:
        parts.append("CMS_Tags")
    if process_images:
        parts.append("Images_Alt_Text")
    if format_math:
        parts.append("Math")
    if normalize_text_structure:
        parts.append("Text_Format")
    if not parts:
        parts.append("Mapping_Data")
    return "_".join(parts) + "_Ready"


def _empty_image_result(run_images_dir: Path, images_zip_path: Path) -> ImagePipelineResult:
    """Create stable empty image artifacts when image handling is not selected."""
    run_images_dir.mkdir(parents=True, exist_ok=True)
    manifest_json = run_images_dir / "image_manifest.json"
    manifest_csv = run_images_dir / "image_manifest.csv"
    manifest_json.write_text(json.dumps({"images": [], "warnings": []}, indent=2), encoding="utf-8")
    manifest_csv.write_text(
        "question,section,choice,section_index,filename,original_part,width_px,height_px,sha256,visual_sha256,duplicate_of,alt_text_written\n",
        encoding="utf-8",
    )
    with zipfile.ZipFile(images_zip_path, "w", compression=zipfile.ZIP_DEFLATED):
        pass
    return ImagePipelineResult([], [], manifest_json, manifest_csv, images_zip_path)


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
    original_equations = _equation_inventory(doc)
    findings: list[Finding] = []
    add_cms_tags, process_images, format_math, normalize_text_structure, prepare_mappings = _selected_processing_sections(options)
    suffix = _output_suffix(add_cms_tags, process_images, format_math, normalize_text_structure, prepare_mappings)
    preserve_cms_mode = not add_cms_tags
    legacy_type_repair = (
        options.processing_mode == "images_math"
        and all(
            value is None
            for value in (
                options.add_cms_tags,
                options.process_images,
                options.format_math,
                options.normalize_text_structure,
                options.prepare_mappings,
                options.format_math_structure,
            )
        )
    )

    if add_cms_tags:
        question_count = _ensure_cms_records(doc, options, findings)
        _audit_question_types(doc, findings, fix_mismatches=True)
        _normalise_choices(doc, findings)
    else:
        question_count = len(_question_starts(doc))
        if question_count:
            findings.append(Finding(0, "passed", "CMS tag insertion was not selected; existing question IDs, snippet IDs and CMS record structure were preserved."))
            for paragraph, number in _question_starts(doc):
                plain_heading = PLAIN_QUESTION_START_RE.fullmatch(paragraph.text.strip())
                suffix_match = QUESTION_START_WITH_SUFFIX_RE.fullmatch(paragraph.text.strip())
                if suffix_match is None and plain_heading is None:
                    suffix_match = PLAIN_QUESTION_START_WITH_SUFFIX_RE.fullmatch(paragraph.text.strip())
                if suffix_match:
                    findings.append(
                        Finding(
                            0,
                            "passed",
                            f"Preserved question-heading text after the question number: {suffix_match.group(2).strip()!r}.",
                            number,
                        )
                    )
        else:
            findings.append(Finding(0, "manual_review", "No tagged or standard plain-text question headings were detected."))
        _audit_question_types(doc, findings, fix_mismatches=legacy_type_repair)

    mapping_instructions: list[MappingInstruction] = []
    if prepare_mappings:
        mapping_instructions = _extract_mapping_instructions(doc, findings)
    else:
        findings.append(Finding(14, "passed", "Mapping-data preparation was not selected; Curriculum and Taxonomy lines were preserved unchanged."))

    if add_cms_tags:
        _audit_unsectioned_record_text(doc, findings)

    if normalize_text_structure:
        _remove_bold_from_questions_and_solutions(doc, findings)

    if format_math:
        _remove_empty_scripts(doc, findings)

    if normalize_text_structure:
        _convert_rupee_symbols(doc, findings)
        _repair_spacing(doc, findings)
        _split_answers(doc, findings)
        _normalise_subparts(doc, findings)

    if format_math:
        _convert_explicit_math(doc, findings)
        _normalise_native_equation_boundaries(doc, findings)
        _repair_italics(doc, findings)
        _equation_and_fraction_audit(doc, findings)

    if normalize_text_structure:
        _audit_assertion_reason(doc, findings)
        _preserve_table_formatting(doc, findings)

    if add_cms_tags:
        _bold_cms_tags(doc, findings)

    project_id = re.sub(r"_q$", "", options.project_question_prefix.strip(), flags=re.I).rstrip("_")
    if original_equations - _equation_inventory(doc):
        raise ValueError("Processing stopped: an original Word equation changed or went missing. No processed document was released. Please share the original for review.")
    findings.append(Finding(2, "passed", "Verified that all non-empty original Word equations retained their mathematical text and operators."))
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
    mapping_csv_path = packages / f"{run_id}_{Path(safe_name).stem}_mapping_data.csv"
    _write_mapping_csv(mapping_csv_path, mapping_instructions)
    if process_images:
        image_result = process_document_images(doc, project_id, run_images_dir, images_zip_path)
        for warning in image_result.warnings:
            findings.append(Finding(0, "manual_review", warning))
        if image_result.artifacts:
            duplicate_count = sum(bool(item.duplicate_of) for item in image_result.artifacts)
            duplicate_note = f" Reused canonical filenames for {duplicate_count} exact duplicate occurrence(s)." if duplicate_count else ""
            findings.append(Finding(0, "fixed", f"Extracted {len(image_result.artifacts)} image occurrence(s), applied CMS filenames and wrote resolved filenames into Word Alt Text.{duplicate_note}"))
        else:
            findings.append(Finding(0, "passed", "No embedded images were detected."))
    else:
        image_result = _empty_image_result(run_images_dir, images_zip_path)
        findings.append(Finding(0, "passed", "Image handling was not selected; embedded images and existing Alt Text were preserved unchanged."))

    output_name = f"{Path(safe_name).stem}_{suffix}.docx"
    output_path = processed / f"{run_id}_{output_name}"
    doc.core_properties.title = f"{Path(safe_name).stem} - CMS Verification Ready"
    doc.core_properties.subject = "Processed and audited for HeyMath CMS verification"
    doc.save(output_path)

    report_path = reports / f"{run_id}_{Path(safe_name).stem}_report.json"
    manual_review_count = sum(f.status == "manual_review" for f in findings)
    numbering_error_count = sum(f.rule == 15 and f.status == "manual_review" for f in findings)
    verification_status = (
        "BLOCKED_NUMBERING_ERROR"
        if numbering_error_count
        else "PENDING_MANUAL_REVIEW"
        if manual_review_count
        else "READY_FOR_VERIFICATION"
    )
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
        "numbering_error_count": numbering_error_count,
        "verification_status": verification_status,
        "options": asdict(options),
        "selected_processing_sections": {
            "add_cms_tags": add_cms_tags,
            "process_images": process_images,
            "format_math": format_math,
            "normalize_text_structure": normalize_text_structure,
            "prepare_mappings": prepare_mappings,
        },
        "mapping_count": len(mapping_instructions),
        "mapping_csv": str(mapping_csv_path),
        "mappings": [asdict(item) for item in mapping_instructions],
        "download_filename": output_name,
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
        "mapping_csv": str(mapping_csv_path),
        "mapping_count": len(mapping_instructions),
        "manual_review_count": manual_review_count,
        "numbering_error_count": numbering_error_count,
    }
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")

    return ProcessResult(
        output_path,
        report_path,
        manifest_path,
        image_result.manifest_json,
        images_zip_path,
        mapping_csv_path,
        original_path,
        run_id,
        question_count,
        len(image_result.artifacts),
        len(mapping_instructions),
        findings,
    )


def process_docx_bytes(data: bytes, filename: str, options: ProcessorOptions, storage_dir: str | Path | None = None) -> ProcessResult:
    return process_docx(data, filename, options, storage_dir)
