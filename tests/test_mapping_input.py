from io import BytesIO

import pandas as pd

from mapping_input import normalize_mapping_table, read_mapping_table, xlsx_sheet_names


MAPPING_COLUMNS = ["Question", "Snippet ID", "Mapping Type", "Path"]


def _xlsx_bytes() -> bytes:
    output = BytesIO()
    instructions = pd.DataFrame({"Read me": ["Choose the Mappings worksheet"]})
    mappings = pd.DataFrame(
        [
            ["1", "219178", "Curriculum", "CBSE NCERT>>Class 4>>Mathematics"],
            ["1", "219178", "Taxonomy", "India>>Class 4>>Mathematics"],
        ],
        columns=MAPPING_COLUMNS,
    )
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        instructions.to_excel(writer, sheet_name="Instructions", index=False)
        mappings.to_excel(writer, sheet_name="Mappings", index=False)
    return output.getvalue()


def test_reads_selected_xlsx_worksheet_as_text():
    data = _xlsx_bytes()

    assert xlsx_sheet_names(data) == ["Instructions", "Mappings"]
    frame = read_mapping_table(data, "separate_mapping_task.xlsx", "Mappings")

    assert list(frame.columns) == MAPPING_COLUMNS
    assert frame.to_dict("records") == [
        {
            "Question": "1",
            "Snippet ID": "219178",
            "Mapping Type": "Curriculum",
            "Path": "CBSE NCERT>>Class 4>>Mathematics",
        },
        {
            "Question": "1",
            "Snippet ID": "219178",
            "Mapping Type": "Taxonomy",
            "Path": "India>>Class 4>>Mathematics",
        },
    ]


def test_csv_still_works_and_blank_rows_are_removed():
    data = (
        "Question,Snippet ID,Mapping Type,Path\n"
        "1,219178,Curriculum,CBSE NCERT>>Class 4>>Mathematics\n"
        ",,,\n"
    ).encode("utf-8")

    frame = read_mapping_table(data, "mapping.csv")

    assert len(frame) == 1
    assert frame.iloc[0]["Snippet ID"] == "219178"


def test_normalizes_wide_project_snippet_taxonomy_curriculum_sheet():
    frame = pd.DataFrame(
        [
            {
                "Project with qno": "project10433_q71",
                "Snippet ID": "219263",
                "Taxonomy": "Mathematics >> Measurement >> Measuring Temperature",
                "Curriculum": "CBSE NCERT >> Class 4 >> Mathematics\nIndia >> Class 4 >> Mathematics",
            }
        ]
    )

    normalized = normalize_mapping_table(frame)

    assert normalized.to_dict("records") == [
        {
            "Question": "project10433_q71",
            "Snippet ID": "219263",
            "Mapping Type": "Curriculum",
            "Path": "CBSE NCERT >> Class 4 >> Mathematics",
        },
        {
            "Question": "project10433_q71",
            "Snippet ID": "219263",
            "Mapping Type": "Curriculum",
            "Path": "India >> Class 4 >> Mathematics",
        },
        {
            "Question": "project10433_q71",
            "Snippet ID": "219263",
            "Mapping Type": "Taxonomy",
            "Path": "Mathematics >> Measurement >> Measuring Temperature",
        },
    ]


def test_normalizes_wide_sheet_with_snippet_only():
    frame = pd.DataFrame(
        [{"Snippet ids": "219263", "Taxonomy": "Mathematics >> Measurement", "Curriculum": ""}]
    )

    normalized = normalize_mapping_table(frame)

    assert normalized.to_dict("records") == [
        {
            "Question": "",
            "Snippet ID": "219263",
            "Mapping Type": "Taxonomy",
            "Path": "Mathematics >> Measurement",
        }
    ]


def test_normalizes_wide_sheet_with_project_only_for_cms_resolution():
    frame = pd.DataFrame(
        [{"Project": "project10433_q71", "Taxonomy": "Mathematics >> Measurement", "Curriculum": ""}]
    )

    normalized = normalize_mapping_table(frame)

    assert normalized.iloc[0].to_dict() == {
        "Question": "project10433_q71",
        "Snippet ID": "",
        "Mapping Type": "Taxonomy",
        "Path": "Mathematics >> Measurement",
    }


def test_normalizes_long_section_alias_and_multiline_paths():
    frame = pd.DataFrame(
        [
            {
                "Question": "project10436_q51",
                "Snippet ID": "219039",
                "Section": "Curriculum",
                "Path": "India >> Class 6\nCBSE NCERT >> Class 7",
            }
        ]
    )

    normalized = normalize_mapping_table(frame)

    assert list(normalized["Path"]) == ["India >> Class 6", "CBSE NCERT >> Class 7"]
