from pathlib import Path
from zipfile import ZipFile

from docx import Document
from lxml import etree
from PIL import Image

from cms_processor import ProcessorOptions, process_docx


def make_sample(path: Path) -> None:
    picture = path.with_name("sample_image.png")
    Image.new("RGB", (40, 30), "white").save(picture)
    doc = Document()
    doc.add_paragraph("Question 1")
    doc.add_paragraph("Type: FIB")
    doc.add_paragraph("Question:")
    p = doc.add_paragraph()
    run = p.add_run("Find  20=2×2×5,then write 3/10.")
    run.bold = True
    doc.add_picture(str(picture))
    doc.add_paragraph("Answers:")
    doc.add_paragraph("(a) 20  (b) 3/10")
    doc.add_paragraph("Solution:")
    s = doc.add_paragraph()
    sr = s.add_run("20=2×2×5")
    sr.bold = True
    doc.add_picture(str(picture))
    table = doc.add_table(rows=2, cols=2)
    table.cell(0, 0).text = "Number"
    table.cell(0, 1).text = "Factor"
    table.cell(1, 0).text = "20"
    table.cell(1, 1).text = "5"
    doc.save(path)


def test_processing_assigns_tags_and_repairs_safe_formatting(tmp_path: Path):
    source = tmp_path / "sample.docx"
    make_sample(source)
    options = ProcessorOptions(
        project_question_prefix="project10436_q",
        start_question_number=7,
        start_snippet_id=220000,
    )
    result = process_docx(source, source.name, options, tmp_path / "storage")
    out = Document(result.output_path)
    texts = [p.text for p in out.paragraphs]

    assert "@Question: 7@" in texts
    assert "@Question id: project10436_q7 @" in texts
    assert "@New snippet id: 220000 @" in texts
    assert "(a) 20" in texts
    assert "(b) " in texts
    assert result.report_path.exists()
    assert result.manifest_path.exists()
    assert result.images_zip_path.exists()
    assert result.image_manifest_path.exists()
    assert result.image_count == 2
    assert any(f.rule == 7 and f.status == "fixed" for f in result.findings)

    question_index = texts.index("@Question:@")
    answer_index = texts.index("@Answers:@")
    for paragraph in out.paragraphs[question_index + 1 : answer_index]:
        assert all(run.bold is not True for run in paragraph.runs)

    assert all(run.bold is True for cell in out.tables[0].rows[0].cells for p in cell.paragraphs for run in p.runs)

    xml = etree.fromstring(ZipFile(result.output_path).read("word/document.xml"))
    ns = {"m": "http://schemas.openxmlformats.org/officeDocument/2006/math"}
    assert len(xml.xpath(".//m:oMath", namespaces=ns)) >= 2
    assert len(xml.xpath(".//m:f", namespaces=ns)) >= 1
    alt_texts = xml.xpath(".//wp:docPr/@descr", namespaces={"wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"})
    assert "project10436_q7_1.gif" in alt_texts
    assert "project10436_a7.gif" in alt_texts


def test_selective_bold_emphasis_is_preserved(tmp_path: Path):
    source = tmp_path / "emphasis.docx"
    doc = Document()
    doc.add_paragraph("Question 1")
    doc.add_paragraph("Question:")
    question = doc.add_paragraph()
    question.add_run("Read the ")
    emphasis = question.add_run("important")
    emphasis.bold = True
    question.add_run(" word.")
    doc.add_paragraph("Answers:")
    doc.add_paragraph("10")
    doc.add_paragraph("Solution:")
    solution = doc.add_paragraph()
    solution.add_run("Use the ")
    solution_emphasis = solution.add_run("factor")
    solution_emphasis.bold = True
    solution.add_run(" shown.")
    doc.save(source)

    result = process_docx(source, source.name, ProcessorOptions(), tmp_path / "storage")
    output = Document(result.output_path)
    question_out = next(p for p in output.paragraphs if p.text == "Read the important word.")
    solution_out = next(p for p in output.paragraphs if p.text == "Use the factor shown.")
    assert next(r for r in question_out.runs if r.text == "important").bold is True
    assert next(r for r in solution_out.runs if r.text == "factor").bold is True


def test_heading_with_difficulty_and_untagged_mcq_are_structured(tmp_path: Path):
    source = tmp_path / "untagged_mcq.docx"
    doc = Document()
    doc.add_paragraph("Q1) Easy")
    doc.add_paragraph("Which pair has a sum of 55?")
    doc.add_paragraph("a) 30 and 25")
    doc.add_paragraph("b) 35 and 20")
    doc.add_paragraph("20 and 35")
    doc.add_paragraph("25 and 30")
    doc.add_paragraph("Answer: b")
    doc.save(source)

    result = process_docx(source, source.name, ProcessorOptions(), tmp_path / "storage")
    texts = [p.text for p in Document(result.output_path).paragraphs]

    assert result.question_count == 1
    assert "@Question: 1@" in texts
    assert "@Type: MCQ@" in texts
    assert "@Difficulty level: Easy @" in texts
    assert "@Question:@" in texts
    assert "@Choices:@" in texts
    assert "@2@ 35 and 20 @correct answer@" in texts
    assert "@Solution:@" in texts
    assert "The correct answer is option (b)." in texts


def test_untagged_fib_answer_is_structured(tmp_path: Path):
    source = tmp_path / "untagged_fib.docx"
    doc = Document()
    doc.add_paragraph("Q6) Medium")
    doc.add_paragraph("A coin finishes at position P. Find P.")
    doc.add_paragraph("Answer 50")
    doc.save(source)

    result = process_docx(source, source.name, ProcessorOptions(), tmp_path / "storage")
    texts = [p.text for p in Document(result.output_path).paragraphs]

    assert result.question_count == 1
    assert "@Type: FIB@" in texts
    assert "@Difficulty level: Average @" in texts
    assert "@Answers:@" in texts
    assert "50" in texts
    assert "@Solution:@" in texts
    assert "The answer is 50." in texts


def test_images_only_mode_preserves_existing_tags_and_text(tmp_path: Path):
    source = tmp_path / "already_tagged.docx"
    picture = tmp_path / "tagged_image.png"
    Image.new("RGB", (40, 30), "white").save(picture)
    doc = Document()
    for text in [
        "@Question: 7@",
        "@Type: MCQ@",
        "@Question id: existing_q7 @",
        "@New snippet id: 900 @",
        "@Difficulty level: Easy @",
        "@Objective: Application @",
        "@Question:@",
        "Choose the diagram.",
    ]:
        doc.add_paragraph(text)
    doc.add_picture(str(picture))
    for text in ["@e@", "@Choices:@", "@1@ One @correct answer@", "@2@ Two", "@3@ Three", "@4@ Four", "@e@", "@Solution:@", "One is correct.", "@e@"]:
        doc.add_paragraph(text)
    doc.save(source)
    original_texts = [p.text for p in doc.paragraphs]

    options = ProcessorOptions(processing_mode="images_only", project_question_prefix="project10436_q")
    result = process_docx(source, source.name, options, tmp_path / "storage")
    output = Document(result.output_path)

    assert result.question_count == 1
    assert [p.text for p in output.paragraphs] == original_texts
    xml = etree.fromstring(ZipFile(result.output_path).read("word/document.xml"))
    alt_texts = xml.xpath(".//wp:docPr/@descr", namespaces={"wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"})
    assert "project10436_q7_1.gif" in alt_texts


def test_exact_duplicate_question_images_reuse_first_filename(tmp_path: Path):
    source = tmp_path / "duplicate_images.docx"
    picture = tmp_path / "fraction_wall.png"
    Image.new("RGB", (60, 40), "white").save(picture)
    doc = Document()
    for question in (2, 5):
        for text in [f"@Question: {question}@", "@Type: FIB@", "@Question:@", "Use the fraction wall."]:
            doc.add_paragraph(text)
        doc.add_picture(str(picture))
        for text in ["@e@", "@Answers:@", "1/2", "@e@", "@Solution:@", "Use the wall.", "@e@"]:
            doc.add_paragraph(text)
    doc.save(source)

    options = ProcessorOptions(processing_mode="images_only", project_question_prefix="project1000_q")
    result = process_docx(source, source.name, options, tmp_path / "storage")
    xml = etree.fromstring(ZipFile(result.output_path).read("word/document.xml"))
    alt_texts = xml.xpath(".//wp:docPr/@descr", namespaces={"wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"})

    assert alt_texts == ["project1000_q2_1.gif", "project1000_q2_1.gif"]
    assert result.image_count == 2
    with ZipFile(result.images_zip_path) as archive:
        gif_names = [name for name in archive.namelist() if name.lower().endswith(".gif")]
    assert gif_names == ["project1000_q2_1.gif"]
    duplicate_artifacts = [item for item in __import__('json').loads(result.image_manifest_path.read_text(encoding="utf-8"))["images"] if item["duplicate_of"]]
    assert len(duplicate_artifacts) == 1
    assert duplicate_artifacts[0]["duplicate_of"] == "project1000_q2_1.gif"
