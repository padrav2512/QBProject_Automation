# CMS DOCX Verification Processor

A deployable Streamlit application that accepts a Word `.docx`, applies safe CMS formatting repairs, adds or normalises confirmed CMS tags, stores the original and processed files on the server, and produces a JSON verification report.

For the team's Streamlit Community Cloud deployment, follow [COMMUNITY_CLOUD_DEPLOYMENT.md](COMMUNITY_CLOUD_DEPLOYMENT.md). Community Cloud is treated as a document processor: users download the processed DOCX, image ZIP and audit report after each run rather than relying on its temporary filesystem for permanent storage.

The app provides five independent processing sections: CMS tags and question structure; images and Alt Text; safe mathematical formatting; text and document formatting; and Curriculum/Taxonomy mapping-data preparation. All are selected by default and can be turned off separately. Existing native equations are preserved, and expressions with uncertain exponent or radical scope are left unchanged and reported for review.

Mapping preparation reads per-question `Curriculum:` and `Taxonomy:` blocks, removes those author-only lines from the processed DOCX and creates a mapping CSV. After the snippets exist in CMS, that CSV can be uploaded in the app's **Apply mappings to CMS** area. A separate mapping-only task can also use a CSV or XLSX with the columns `Question`, `Snippet ID`, `Mapping Type`, and `Path`; when an XLSX has several worksheets, the user selects the worksheet containing the mapping rows. The mapper only adds missing paths; existing mappings are retained and never removed or replaced.

Full preparation accepts plain headings that include both difficulty and objective, such as `Question 44: Easy, Comprehension`. Image extraction supports modern DrawingML pictures and legacy Word VML picture containers.

## Run locally

```powershell
cd "C:\PlaySheets & Math Bites\Class 5 Ch6 CBSE NCERT Questions\cms_docx_verifier_app"
python -m pip install -r requirements.txt
streamlit run app.py
```

The default local URL is `http://localhost:8501`.

## Server storage

By default, files are written below `data/`:

- `data/uploads/` — immutable uploaded originals
- `data/processed/` — generated CMS verification documents
- `data/reports/` — JSON audit reports
- `data/verification_queue/` — one machine-readable queue manifest per submission, with status `READY_FOR_VERIFICATION` or `PENDING_MANUAL_REVIEW`
- `data/images/<run-id>/` — extracted GIFs plus CSV/JSON image manifests
- `data/packages/` — downloadable image ZIP and mapping CSV packages

Set `CMS_STORAGE_DIR` to a persistent mounted directory on the server. Set `CMS_MAX_UPLOAD_MB` to change the app-level file-size limit.

## Docker deployment

```bash
docker build -t cms-docx-verifier .
docker run --rm -p 8501:8501 -v cms-verifier-data:/data cms-docx-verifier
```

For a public deployment, place the container behind HTTPS and your organisation’s authentication layer. Back up the mounted `/data` volume.

## Input expectations

### Plain-text input (no `@` characters required)

Select **CMS tags and question structure**. Put each line below in its own Word paragraph:

```text
Question 26: Average, Comprehension
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
6 times 7 equals 42.
```

Optional `Question:` and `Choices:` / `Options:` headings are supported. `Question type:` is accepted as an alias for `Type:`, and `Correct answer:` is accepted as an alias for `Answer:`. For MCQs, use four options a)–d). A correct key may be written as `b`, `b)`, `(b)`, `2` or `b) choice text`; supplied choice text must agree with the selected choice. For multiple FIB answers, use `Answers:` followed by one labelled answer per paragraph. For single-letter FIB answers, use `Type: FIB` and an `Answers:` section to avoid ambiguity with an MCQ key. Use Enter between paragraphs. The app generates all CMS markers in the output; existing tagged documents remain supported. Check the audit report for ambiguous or incomplete input.

The intended input is a quality-checked Word document; it does not need to contain CMS tags. Each question must still have an identifiable boundary, such as a standalone `Question 1` or `Q1` heading, and recognisable Question, Answers or Choices, and Solution sections. Existing CMS-style documents containing `@Question: n@` records are also accepted, including harmless spacing variants such as `@Question:10 @`. A heading such as `@Question: 29@ case study` is counted, but is reported for manual review. When CMS tag processing is selected, the suffix is moved to the start of the Question block so it is not lost. The app then creates or repairs:

- `@Question: n@`
- `@Type: FIB@` or `@Type: MCQ@`
- `@Question id: … @`
- `@New snippet id: … @`
- `@Difficulty level: … @`
- `@Objective: … @`
- `@Question:@`, `@Choices:@`, `@Answers:@`, `@Solution:@`
- closing `@e@` markers

The processor removes whole-block bold formatting when an entire Question, Choices or Solution block is bold, while preserving selective bold emphasis. It makes CMS metadata and section tag paragraphs bold, while leaving choice markers, `@correct answer@` and `@e@` in ordinary text. It preserves author-supplied italics and adds italics only for high-confidence variables and geometry labels identified through equations or explicit cues such as `point A`, `line BC`, `radius r` and `let x`. Recognised measurement units and ordinary uses such as `A circle` and `Option A` are protected; uncertain labels remain unchanged and are reported. It converts `₹` before a numeric amount to `Rs `, normalises spacing across adjacent Word runs, and removes spaces before punctuation. Selectively styled ordinary text is retained rather than flattened during optional math conversion. Table formatting is preserved; authors must identify and bold header rows or header columns in the source document. Top-level subparts use `(a), (b)` and nested subparts use `(i), (ii)`. It removes blank equation, exponent and subscript templates, converts safe slash fractions and self-contained radicals into native Word equations, and intentionally reports rather than guesses when an automatic change could alter meaning. Examples include `1/2r`, `√a+b`, `√5/9`, `√r²`, equation screenshots, ambiguous nested subparts, missing answer/solution blocks, and Assertion–Reason items without exactly one marked correct answer.

It also extracts embedded question, choice and solution images, creates CMS GIF filenames, and writes each resolved filename into the image's Word Alt Text description. Question images use `project..._qN_1.gif`, choice images use `project..._qN_c1.gif` (or `..._c1_1.gif`, `..._c1_2.gif` for multiple images in one choice), and solution images use `project..._aN.gif` or `project..._aN_1.gif`, `project..._aN_2.gif` when a solution contains multiple images. An image in an FIB Answers block is reported as a blocking issue because CMS does not accept it.

For mapping preparation, place the paths inside each question. The first path may follow the label, and additional paths can follow as separate Word paragraphs:

```text
Curriculum: CBSE NCERT >> Class 4 >> Mathematics >> Measuring Length >> Metres and Centimetres
India >> Class 4 >> Mathematics >> Measurement
India >> Class 4 >> Mathematics >> Case Study Based Questions
Taxonomy: Mathematics >> Measurement >> Length
```

## Production notes

- The app stores files on the server filesystem; it does not send them to a third party.
- Use a persistent volume in Docker, Kubernetes or your hosting provider.
- Add organisation authentication before exposing the app on the public Internet.
- Malware scanning and retention policies should be added at the infrastructure layer if required by your organisation.
