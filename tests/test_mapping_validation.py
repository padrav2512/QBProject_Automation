import pandas as pd

from app import validate_mapping_plan


class ReadOnlyMapper:
    def __init__(self):
        self.reads = []
        self.action_checks = []

    def get_mappings(self, snippet_id, *, question_type="static"):
        self.reads.append((snippet_id, question_type))
        return {
            "name": f"{question_type}-{snippet_id}",
            "curriculum": ["Existing >> Curriculum"],
            "taxonomy": [],
        }

    def validate_question_action(
        self,
        snippet_id,
        curriculum_paths,
        taxonomy_paths,
        *,
        action,
        question_type="static",
    ):
        self.action_checks.append((snippet_id, question_type, action))
        return None

    def validate_path(self, path, *, taxonomy=False):
        return None

    def resolve_snippet_id(self, question_name):
        raise AssertionError("The test supplies snippet IDs")

    def map_question(self, *args, **kwargs):
        raise AssertionError("Read-only validation must never map a question")


def test_read_only_validation_keeps_static_and_dynamic_same_id_separate():
    frame = pd.DataFrame(
        [
            {
                "Question": "",
                "Snippet ID": "1121",
                "Snippet Type": "Static",
                "Mapping Type": "Curriculum",
                "Path": "CBSE NCERT >> Class 7",
            },
            {
                "Question": "",
                "Snippet ID": "1121",
                "Snippet Type": "Dynamic",
                "Mapping Type": "Curriculum",
                "Path": "CBSE NCERT >> Class 7",
            },
        ]
    )
    mapper = ReadOnlyMapper()

    resolved, errors, preview = validate_mapping_plan(mapper, frame, "add")

    assert errors == []
    assert resolved.equals(frame)
    assert mapper.reads == [(1121, "Static"), (1121, "Dynamic")]
    assert mapper.action_checks == [
        (1121, "Static", "add"),
        (1121, "Dynamic", "add"),
    ]
    assert [row["CMS question"] for row in preview] == [
        "Static-1121",
        "Dynamic-1121",
    ]
    assert [row["Route"] for row in preview] == [
        "/question/1121/html_curriculum",
        "/snippet/1121/html_curriculum",
    ]
