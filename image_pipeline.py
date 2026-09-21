from __future__ import annotations

import csv
import hashlib
import io
import json
import re
import zipfile
from dataclasses import asdict, dataclass
from pathlib import Path

from docx.document import Document as _Document
from docx.oxml.ns import qn
from PIL import Image, ImageChops


QUESTION_RE = re.compile(r"^@Question:\s*(\d+)@$", re.I)
CHOICE_RE = re.compile(r"^@([1-4])@")
SECTION_MARKERS = {"@Question:@", "@Choices:@", "@Answers:@", "@Solution:@"}
WHITE_TOLERANCE = 12


@dataclass
class ImageArtifact:
    question: int | None
    section: str
    choice: int | None
    section_index: int
    filename: str
    original_part: str
    width_px: int | None
    height_px: int | None
    sha256: str
    visual_sha256: str
    duplicate_of: str | None
    alt_text_written: bool


@dataclass
class ImagePipelineResult:
    artifacts: list[ImageArtifact]
    warnings: list[str]
    manifest_json: Path
    manifest_csv: Path
    images_zip: Path


@dataclass
class _Occurrence:
    question: int | None
    section: str
    choice: int | None
    blob: bytes
    original_part: str
    alt_node: object | None
    alt_attribute: str
    filename: str = ""
    section_index: int = 0
    warning: str | None = None


def _remove_outer_white_space(image: Image.Image) -> Image.Image:
    rgba = image.convert("RGBA")
    white = Image.new("RGBA", rgba.size, (255, 255, 255, 255))
    white.alpha_composite(rgba)
    rgb = white.convert("RGB")
    background = Image.new("RGB", rgb.size, (255, 255, 255))
    diff = ImageChops.difference(rgb, background).convert("L")
    diff = diff.point(lambda value: 255 if value > WHITE_TOLERANCE else 0)
    bbox = diff.getbbox()
    return image.copy() if bbox is None else image.crop(bbox)


def _save_gif(blob: bytes, output: Path) -> tuple[int, int]:
    image = Image.open(io.BytesIO(blob))
    image.load()
    image = _remove_outer_white_space(image.copy())
    if "A" in image.getbands():
        rgba = image.convert("RGBA")
        background = Image.new("RGBA", rgba.size, (255, 255, 255, 255))
        background.alpha_composite(rgba)
        image = background.convert("RGB")
    else:
        image = image.convert("RGB")
    size = image.size
    image.save(output, format="GIF", optimize=True)
    return size


def _normalised_visual_signature(blob: bytes) -> tuple[str, int, int]:
    """Hash normalized visible pixels so harmless file metadata/encoding changes are ignored."""
    image = Image.open(io.BytesIO(blob))
    image.load()
    image = _remove_outer_white_space(image.copy())
    if "A" in image.getbands():
        rgba = image.convert("RGBA")
        background = Image.new("RGBA", rgba.size, (255, 255, 255, 255))
        background.alpha_composite(rgba)
        image = background.convert("RGB")
    else:
        image = image.convert("RGB")
    width, height = image.size
    digest = hashlib.sha256(f"{width}x{height}:RGB:".encode("ascii") + image.tobytes()).hexdigest()
    return digest, width, height


def _deduplication_bucket(occurrence: _Occurrence) -> str | None:
    if occurrence.section in {"question", "choices"} and occurrence.question is not None:
        return "question_gifs"
    if occurrence.section == "solution" and occurrence.question is not None:
        return "solution_gifs"
    return None


def _relationship_blob(doc: _Document, relationship_id: str) -> tuple[bytes, str] | None:
    part = doc.part.related_parts.get(relationship_id)
    if part is None or not hasattr(part, "blob"):
        return None
    return part.blob, str(getattr(part, "partname", relationship_id))


def _collect_occurrences(doc: _Document) -> list[_Occurrence]:
    occurrences: list[_Occurrence] = []
    question: int | None = None
    section = "unclassified"
    choice: int | None = None

    for paragraph in doc.element.xpath(".//w:body//w:p"):
        text = "".join(paragraph.xpath(".//w:t/text() | .//m:t/text()"))
        stripped = text.strip()
        question_match = QUESTION_RE.fullmatch(stripped)
        if question_match:
            question = int(question_match.group(1))
            section = "metadata"
            choice = None
        elif stripped in SECTION_MARKERS:
            section = {
                "@Question:@": "question",
                "@Choices:@": "choices",
                "@Answers:@": "answers",
                "@Solution:@": "solution",
            }[stripped]
            choice = None
        elif stripped == "@e@":
            section = "between_sections"
            choice = None
        elif section == "choices":
            choice_match = CHOICE_RE.match(stripped)
            if choice_match:
                choice = int(choice_match.group(1))

        for drawing in paragraph.xpath(".//w:drawing"):
            doc_properties = drawing.xpath(".//wp:docPr")
            alt_node = doc_properties[0] if doc_properties else None
            for blip in drawing.xpath(".//a:blip"):
                rid = blip.get(qn("r:embed")) or blip.get(qn("r:link"))
                if not rid:
                    continue
                payload = _relationship_blob(doc, rid)
                if payload:
                    occurrences.append(_Occurrence(question, section, choice, payload[0], payload[1], alt_node, "descr"))

        for pict in paragraph.xpath(".//w:pict"):
            shapes = pict.xpath(".//v:shape")
            alt_node = shapes[0] if shapes else None
            for image_data in pict.xpath(".//v:imagedata"):
                rid = image_data.get(qn("r:id"))
                if not rid:
                    continue
                payload = _relationship_blob(doc, rid)
                if payload:
                    occurrences.append(_Occurrence(question, section, choice, payload[0], payload[1], alt_node, "alt"))

    return occurrences


def _assign_names(occurrences: list[_Occurrence], project_id: str) -> list[str]:
    warnings: list[str] = []
    grouped: dict[tuple[int | None, str, int | None], list[_Occurrence]] = {}
    for occurrence in occurrences:
        grouped.setdefault((occurrence.question, occurrence.section, occurrence.choice), []).append(occurrence)

    for (question, section, choice), items in grouped.items():
        for index, occurrence in enumerate(items, 1):
            occurrence.section_index = index
            if question is None:
                occurrence.filename = f"REVIEW_{project_id}_unassigned_{index}.gif"
                occurrence.warning = "Image occurs before an identifiable question."
            elif section == "question":
                occurrence.filename = f"{project_id}_q{question}_{index}.gif"
            elif section == "choices" and choice is not None:
                if len(items) == 1:
                    occurrence.filename = f"{project_id}_q{question}_c{choice}.gif"
                else:
                    occurrence.filename = f"{project_id}_q{question}_c{choice}_{index}.gif"
            elif section == "solution":
                occurrence.filename = f"{project_id}_a{question}.gif" if len(items) == 1 else f"{project_id}_a{question}_{index}.gif"
            elif section == "answers":
                occurrence.filename = f"REVIEW_{project_id}_q{question}_answer_{index}.gif"
                occurrence.warning = f"Question {question} contains an image in the FIB Answers block. CMS does not accept answer-block images; remove it before verification."
            else:
                occurrence.filename = f"REVIEW_{project_id}_q{question}_{section}_{index}.gif"
                occurrence.warning = f"Question {question} contains an image outside Question, Choices or Solution content."
            if occurrence.warning:
                warnings.append(occurrence.warning)
    return warnings


def process_document_images(doc: _Document, project_id: str, output_dir: Path, zip_path: Path) -> ImagePipelineResult:
    output_dir.mkdir(parents=True, exist_ok=True)
    occurrences = _collect_occurrences(doc)
    warnings = _assign_names(occurrences, project_id)
    artifacts: list[ImageArtifact] = []

    canonical_by_visual: dict[tuple[str, str], _Occurrence] = {}
    saved_files: dict[str, tuple[int | None, int | None, str]] = {}

    for occurrence in occurrences:
        visual_sha = ""
        duplicate_of: str | None = None
        try:
            visual_sha, _, _ = _normalised_visual_signature(occurrence.blob)
        except Exception as exc:
            warnings.append(f"Could not calculate a visual fingerprint for {occurrence.original_part}: {exc}")

        bucket = _deduplication_bucket(occurrence)
        dedupe_key = (bucket, visual_sha) if bucket and visual_sha else None
        if dedupe_key and dedupe_key in canonical_by_visual:
            canonical = canonical_by_visual[dedupe_key]
            occurrence.filename = canonical.filename
            duplicate_of = canonical.filename
        elif dedupe_key:
            canonical_by_visual[dedupe_key] = occurrence

        alt_written = not occurrence.filename.startswith("REVIEW_") and occurrence.alt_node is not None
        if alt_written:
            occurrence.alt_node.set(occurrence.alt_attribute, occurrence.filename)
        output = output_dir / occurrence.filename
        if occurrence.filename in saved_files:
            width, height, file_sha = saved_files[occurrence.filename]
        else:
            width = height = None
            try:
                width, height = _save_gif(occurrence.blob, output)
            except Exception as exc:
                warnings.append(f"Could not convert {occurrence.original_part} to GIF: {exc}")
                output.write_bytes(occurrence.blob)
            file_sha = hashlib.sha256(output.read_bytes()).hexdigest()
            saved_files[occurrence.filename] = (width, height, file_sha)
        artifacts.append(
            ImageArtifact(
                question=occurrence.question,
                section=occurrence.section,
                choice=occurrence.choice,
                section_index=occurrence.section_index,
                filename=occurrence.filename,
                original_part=occurrence.original_part,
                width_px=width,
                height_px=height,
                sha256=file_sha,
                visual_sha256=visual_sha,
                duplicate_of=duplicate_of,
                alt_text_written=alt_written,
            )
        )

    manifest_json = output_dir / "image_manifest.json"
    manifest_csv = output_dir / "image_manifest.csv"
    manifest_json.write_text(
        json.dumps({"project_id": project_id, "images": [asdict(item) for item in artifacts], "warnings": warnings}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    with manifest_csv.open("w", newline="", encoding="utf-8-sig") as handle:
        fieldnames = list(ImageArtifact.__dataclass_fields__)
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(asdict(item) for item in artifacts)

    zip_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as archive:
        for file in sorted(output_dir.iterdir()):
            if file.is_file():
                archive.write(file, arcname=file.name)

    return ImagePipelineResult(artifacts, warnings, manifest_json, manifest_csv, zip_path)
