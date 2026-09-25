from io import BytesIO

import pandas as pd

from mapping_input import read_mapping_table, xlsx_sheet_names


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
