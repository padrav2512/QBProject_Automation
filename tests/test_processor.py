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
