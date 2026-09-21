# Shared CMS verification workflow

## Recommendation

Use one application with a staged workflow and one immutable submission folder per run. Do not ask users to run the six scripts separately. Each run should produce a complete package containing the original DOCX, processed DOCX, extracted GIFs, image manifest, validation report and a machine-readable verification manifest.

For three users, a centrally hosted Streamlit app is the simplest option. Each user downloads an image ZIP, extracts it to any local folder and uploads the GIFs manually. The final DOCX does not need that local path: Anand's importer resolves each image by the GIF filename stored in the image's Word Alt Text description.

## Deterministic steps

The following steps can be implemented reliably in code:

1. Read every question record, including paragraphs in tables.
2. Detect missing, duplicated or out-of-order question numbers.
3. Infer the intended FIB/MCQ record structure from the quality-checked Word document, using question headings and section labels such as Question, Answers, Choices and Solution.
4. Accept only the project ID and first snippet ID from the user.
5. Calculate the complete ranges automatically:
   - `project10436_q1` through `project10436_qN`
   - starting snippet ID through `start + N - 1`
6. Add the complete CMS record structure, including question numbers, type, question ID, snippet ID, difficulty, objective, content-section markers, choice or answer markers, solution marker and every required `@e@` closure.
7. Insert or replace the question and snippet IDs.
8. Validate the newly generated CMS tags, FIB/MCQ structure, metadata values, choice order, one correct-answer marker and block-closing `@e@` markers.
9. Check both ID sets for uniqueness and sequence continuity.
10. Extract embedded raster images in document order and record whether each image occurs in the Question, Choices, Answers or Solution section.
11. Include question, choice and solution images. Convert supported raster images to GIF, flatten transparency onto white, remove only empty outer white space and name each occurrence according to the CMS convention:
    - Question images: `project10427_q19_1.gif`, `project10427_q19_2.gif`, and so on.
    - Choice images: `project10427_q15_c1.gif`, `project10427_q15_c2.gif`, and so on.
    - One solution image: `project10250_a153.gif`.
    - Multiple images in one solution: `project10250_a154_1.gif`, `project10250_a154_2.gif`, and so on.
12. Write the generated GIF filename into the image's Word Alt Text description field in the final DOCX.
13. Produce an image manifest containing question number, question ID, section, image occurrence number, original media filename, generated GIF filename, pixel dimensions and file hash.
14. Store the original, processed document, images and reports under a unique run ID so concurrent users cannot overwrite one another.
15. Run the complete validation again after all transformations.

## Steps requiring a controlled decision or human review

- Mathematical and pedagogical correctness remains human-approved input.
- Word drawings and shapes are not always embedded raster images. They need rendering or manual review before becoming GIF files.
- GIF crops should be displayed in an image-review gallery because some diagrams intentionally contain surrounding white space.
- An embedded picture cannot be classified reliably as a diagram versus an equation screenshot from DOCX structure alone. Its Question/Choices/Answers/Solution location can still be identified deterministically.
- Selective bold emphasis inside a question or solution must be preserved. Bold is removed only when all substantive text in the entire Question or Solution block is bold.
- Completely blank Word equation objects, including blank equation, exponent and subscript templates, must be removed.
- Slash-style fractions must be converted into stacked fractions inside native Word equations.
- Multiple images inside one choice use `..._c1_1.gif`, `..._c1_2.gif`, and so on.
- Images inside an FIB Answers block are not supported by CMS and must be reported as a blocking issue for removal.

## Proposed application stages

### Stage 1 — Upload and preflight

The user uploads the quality-checked DOCX. CMS tags are not required in the input. The app first performs a read-only structural audit: it identifies question boundaries, question numbers, FIB/MCQ evidence, Question/Answers/Choices/Solution sections and subparts. It stops only when the intended structure cannot be inferred safely. The preflight report lists missing or duplicated question numbers, ambiguous section boundaries, incomplete choice sets and content requiring manual review.

### Stage 2 — Project configuration

The user enters:

- Project ID, for example `project10436`
- First snippet ID, for example `218989`
- Optional image-export preference

The app detects the question count and displays the calculated final values. The user does not enter the last question ID or last snippet ID; those are derived values.

### Stage 3 — Transform and extract

The app works on a copy. It first creates the complete CMS structure and adds all required tags and `@e@` closures. It then inserts IDs, validates the generated structure, applies safe formatting repairs, extracts images from the processed copy and writes each resolved GIF filename into the corresponding image's Alt Text description. Extracted image names include the project ID from the outset, so a separate rename script is unnecessary.

### Stage 4 — Review package

The app presents:

- Processed DOCX download
- Image gallery grouped by question
- Images ZIP download
- CSV/JSON image manifest
- Rule-by-rule validation report

Each colleague extracts the downloaded ZIP into any local folder. The folder's absolute path does not have to be written into the DOCX.

### Stage 5 — CMS image upload

The colleague uploads the GIFs manually. The deterministic filenames and manifest provide a checklist and prevent question/image mismatches.

### Stage 6 — Alt Text and postflight

Before delivering the final DOCX, the app maps every extracted image back to its original occurrence and writes the generated GIF filename into Word's Alt Text description field. The embedded image remains visible in Word. The app then verifies that every resolved image has matching Alt Text and a matching file in the image ZIP.

Anand's tool subsequently reads the tagged DOCX, uses the Alt Text filename to locate the corresponding manually uploaded CMS image and populates the CMS editor. Our application does not upload questions or images to CMS.

## Review of `StaticQuestionsPythonCode.rar`

Useful components:

- `Geethavalidate_questions.py` performs pre-ID template checks.
- `InsertIds.py` creates sequential IDs and backs up the document.
- `image_extraction.py` extracts embedded images, trims outside white space and creates GIFs named by question and occurrence.
- `AppendProjectIdGif.py` prefixes GIF filenames with the project ID.
- `JayaUniqueness_Validation.py` repeats validation and adds uniqueness checks.

Changes required before shared use:

- Consolidate the two validators into one rules engine and one report format.
- Read paragraphs in tables as well as the main document body.
- Use the current `(a), (b)` top-level subpart rule; the supplied validator checks Roman-numeral subparts only.
- Remove the requirement for users to type final IDs. Calculate them from the detected question count.
- Make question-marker handling consistent. `InsertIds.py` counts an empty `@Question: @` form while the validators expect numbered `@Question: n@` markers.
- Prefix images during extraction instead of renaming them in place later.
- Never place all users' images into a shared folder named only `images`; use a run ID and user/project namespace.
- Create a manifest before any CMS upload.
- Do not claim a backup during image renaming unless one is actually created; `AppendProjectIdGif.py` currently renames files in place without making the backup described in the instructions.
- Add explicit handling for unsupported image formats and Word shapes.
- Add a post-transformation validation pass before a run is marked ready.

## Confirmed filename decisions

1. The number after `a` is the question number, so the solution for Question 154 uses `project10250_a154.gif`.
2. Multiple images inside Choice 1 use `..._c1_1.gif`, `..._c1_2.gif`, and so on.
3. FIB Answers blocks cannot contain images; any such occurrence blocks verification until the image is removed.
