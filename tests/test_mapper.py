from heymath_cms_mapper import HeyMathCmsMapper


def test_mapper_adds_missing_paths_and_retains_existing_mappings(monkeypatch):
    mapper = HeyMathCmsMapper("https://cms.example", "shared", "secret")
    mapper._curriculum_index = {
        "Existing >> Curriculum": ["10"],
        "New >> Curriculum": ["11"],
    }
    mapper._taxonomy_index = {
        "Existing >> Taxonomy": ["20"],
        "New >> Taxonomy": ["21"],
    }
    mapper._trees_loaded_at = 10**12
    mapper._logged_in = True

    before = {
        "name": "Question 1",
        "curriculum_token": "cur-token",
        "taxonomy_token": "tax-token",
        "curriculum": ["Existing >> Curriculum"],
        "taxonomy": ["Existing >> Taxonomy"],
    }
    after = {
        **before,
        "curriculum": ["Existing >> Curriculum", "New >> Curriculum"],
        "taxonomy": ["Existing >> Taxonomy", "New >> Taxonomy"],
    }
    pages = iter([before, before, after])
    monkeypatch.setattr(mapper, "_load_mapping_page", lambda snippet_id: next(pages))
    posts = []
    monkeypatch.setattr(mapper, "_post_form", lambda *args: posts.append(args))

    result = mapper.map_question(
        1001,
        ["Existing>>Curriculum", "New >> Curriculum"],
        ["Existing >> Taxonomy", "New>>Taxonomy"],
    )

    assert result.success
    assert result.curriculum_after == after["curriculum"]
    assert result.taxonomy_after == after["taxonomy"]
    assert result.already_present == [
        "Curriculum: Existing >> Curriculum",
        "Taxonomy: Existing >> Taxonomy",
    ]
    assert result.added == [
        "Curriculum: New >> Curriculum",
        "Taxonomy: New >> Taxonomy",
    ]
    assert [post[4] for post in posts] == ["11", "21"]


def test_mapper_saves_nothing_when_any_path_is_invalid(monkeypatch):
    mapper = HeyMathCmsMapper("https://cms.example", "shared", "secret")
    mapper._curriculum_index = {"Known >> Path": ["10"]}
    mapper._taxonomy_index = {"Known >> Taxonomy": ["20"]}
    mapper._trees_loaded_at = 10**12
    mapper._logged_in = True
    page = {
        "name": "Question 2",
        "curriculum_token": "cur-token",
        "taxonomy_token": "tax-token",
        "curriculum": [],
        "taxonomy": [],
    }
    monkeypatch.setattr(mapper, "_load_mapping_page", lambda snippet_id: page)
    posts = []
    monkeypatch.setattr(mapper, "_post_form", lambda *args: posts.append(args))

    result = mapper.map_question(1002, ["Unknown >> Path"], ["Known >> Taxonomy"])

    assert not result.success
    assert "not found" in result.errors[0]
    assert posts == []
