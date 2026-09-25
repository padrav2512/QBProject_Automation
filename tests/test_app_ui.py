from pathlib import Path

from streamlit.testing.v1 import AppTest


def test_app_renders_five_processing_sections_and_mapping_stage():
    app_path = Path(__file__).parents[1] / "app.py"
    app = AppTest.from_file(str(app_path)).run(timeout=30)

    assert not app.exception
    assert [checkbox.label for checkbox in app.sidebar.checkbox[:5]] == [
        "1. CMS tags and question structure",
        "2. Handle images and Alt Text",
        "3. Apply safe mathematical formatting",
        "4. Normalize text and document formatting",
        "5. Prepare curriculum and taxonomy mapping data",
    ]
    assert all(checkbox.value for checkbox in app.sidebar.checkbox[:5])
    assert [header.value for header in app.sidebar.header[:2]] == [
        "1. Choose processing",
        "2. CMS and image naming",
    ]
    assert any(
        "Question: 3 Case study" in caption.value and "@Question: 3@ Case study" in caption.value
        for caption in app.sidebar.caption
    )
    assert "Validate and apply mappings to CMS" in [button.label for button in app.button]
