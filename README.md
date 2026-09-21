# CMS DOCX Verification Processor

A deployable Streamlit application that accepts a Word `.docx`, applies safe CMS formatting repairs, adds or normalises confirmed CMS tags, stores the original and processed files on the server, and produces a JSON verification report.

For the team's Streamlit Community Cloud deployment, follow [COMMUNITY_CLOUD_DEPLOYMENT.md](COMMUNITY_CLOUD_DEPLOYMENT.md). Community Cloud is treated as a document processor: users download the processed DOCX, image ZIP and audit report after each run rather than relying on its temporary filesystem for permanent storage.

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
- `data/packages/` — downloadable image ZIP packages

Set `CMS_STORAGE_DIR` to a persistent mounted directory on the server. Set `CMS_MAX_UPLOAD_MB` to change the app-level file-size limit.

## Docker deployment

```bash
docker build -t cms-docx-verifier .
docker run --rm -p 8501:8501 -v cms-verifier-data:/data cms-docx-verifier
```

For a public deployment, place the container behind HTTPS and your organisation’s authentication layer. Back up the mounted `/data` volume.

## Input expectations

The intended input is a quality-checked Word document; it does not need to contain CMS tags. Each question must still have an identifiable boundary, such as a standalone `Question 1` or `Q1` heading, and recognisable Question, Answers or Choices, and Solution sections. Existing CMS-style documents containing `@Question: n@` records are also accepted. The app then creates or repairs:

- `@Question: n@`
- `@Type: FIB@` or `@Type: MCQ@`
- `@Question id: … @`
- `@New snippet id: … @`
- `@Difficulty level: … @`
- `@Objective: … @`
- `@Question:@`, `@Choices:@`, `@Answers:@`, `@Solution:@`
- closing `@e@` markers

The processor removes whole-block bold formatting when an entire Question or Solution is bold, while preserving selective bold emphasis. It makes CMS metadata and section tag paragraphs bold, while leaving choice markers, `@correct answer@` and `@e@` in ordinary text. It enforces the variable-italics convention across Questions, Answers, Choices and Solutions. Table header rows and first columns are always bold. Top-level subparts use `(a), (b)` and nested subparts use `(i), (ii)`. It removes blank equation, exponent and subscript templates, converts slash fractions into native stacked Word fractions, and converts unambiguous standalone mathematical expressions to native Word equations. It intentionally reports rather than guesses when an automatic change could alter meaning. Examples include equation screenshots, ambiguous nested subparts, missing answer/solution blocks, and Assertion–Reason items without exactly one marked correct answer.

It also extracts embedded question, choice and solution images, creates CMS GIF filenames, and writes each resolved filename into the image's Word Alt Text description. Question images use `project..._qN_1.gif`, choice images use `project..._qN_c1.gif` (or `..._c1_1.gif`, `..._c1_2.gif` for multiple images in one choice), and solution images use `project..._aN.gif` or `project..._aN_1.gif`, `project..._aN_2.gif` when a solution contains multiple images. An image in an FIB Answers block is reported as a blocking issue because CMS does not accept it.

## Production notes

- The app stores files on the server filesystem; it does not send them to a third party.
- Use a persistent volume in Docker, Kubernetes or your hosting provider.
- Add organisation authentication before exposing the app on the public Internet.
- Malware scanning and retention policies should be added at the infrastructure layer if required by your organisation.
