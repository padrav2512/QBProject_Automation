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


def test_validate_path_suggests_a_close_cms_path():
    mapper = HeyMathCmsMapper("https://cms.example", "shared", "secret")
    mapper._curriculum_index = {"Known >> Path": ["10"]}
    mapper._taxonomy_index = {
        "Mathematics >> Measurement >> Measuring Temperature": ["20"],
        "Mathematics >> Geometry >> Circles": ["21"],
    }
    mapper._trees_loaded_at = 10**12
    mapper._logged_in = True

    error = mapper.validate_path(
        "Mathematics >> Measurement >> Measuring Temparature",
        taxonomy=True,
    )

    assert error is not None
    assert "Path not found" in error
    assert "Did you mean: Mathematics >> Measurement >> Measuring Temperature?" in error


def test_resolve_snippet_id_requires_exact_question_name(monkeypatch):
    mapper = HeyMathCmsMapper("https://cms.example", "shared", "secret")
    mapper._logged_in = True
    monkeypatch.setattr(mapper, "_site_authenticity_token", lambda: "search-token")

    class Response:
        text = '<a href="/question/219263/edit">project10433_q71</a><a href="/question/219264/edit">project10433_q710</a>'
        url = "https://cms.example/site/search"

        @staticmethod
        def raise_for_status():
            return None

    monkeypatch.setattr(mapper._session, "post", lambda *args, **kwargs: Response())
    monkeypatch.setattr(
        mapper,
        "_load_mapping_page",
        lambda snippet_id: {"name": "project10433_q71" if snippet_id == 219263 else "project10433_q710"},
    )

    snippet_id, error = mapper.resolve_snippet_id("project10433_q71")

    assert snippet_id == 219263
    assert error is None


def test_resolve_snippet_id_reports_missing_exact_match(monkeypatch):
    mapper = HeyMathCmsMapper("https://cms.example", "shared", "secret")
    mapper._logged_in = True
    monkeypatch.setattr(mapper, "_site_authenticity_token", lambda: "search-token")

    class Response:
        text = '<a href="/question/219264/edit">project10433_q710</a>'
        url = "https://cms.example/site/search"

        @staticmethod
        def raise_for_status():
            return None

    monkeypatch.setattr(mapper._session, "post", lambda *args, **kwargs: Response())
    monkeypatch.setattr(mapper, "_load_mapping_page", lambda snippet_id: {"name": "project10433_q710"})

    snippet_id, error = mapper.resolve_snippet_id("project10433_q71")

    assert snippet_id is None
    assert "No exact CMS question" in error


def test_mapper_accepts_realistic_form_spacing_and_single_quotes():
    page = """
    <form action = '/curriculum_trees/add_mapping' method = 'post'>
      <input value = 'token&amp;value' name = 'authenticity_token' type = 'hidden'>
    </form>
    """

    assert HeyMathCmsMapper._form_token(page, "/curriculum_trees/add_mapping") == "token&value"


def test_mapper_posts_with_mapping_page_referer(monkeypatch):
    mapper = HeyMathCmsMapper("https://cms.example", "shared", "secret")
    captured = {}

    class Response:
        status_code = 200

    def fake_post(url, **kwargs):
        captured["url"] = url
        captured.update(kwargs)
        return Response()

    monkeypatch.setattr(mapper._session, "post", fake_post)
    mapper._post_form("/curriculum_trees/add_mapping", "token", 123, "node_id", "10", "Add Mapping Form")

    assert captured["headers"] == {"Referer": "https://cms.example/question/123/html_curriculum"}
    assert captured["data"]["node_id"] == "10"


def test_connection_is_read_only_and_reports_tree_sizes(monkeypatch):
    mapper = HeyMathCmsMapper("https://cms.example", "shared", "secret")
    calls = []

    monkeypatch.setattr(mapper, "_ensure_logged_in", lambda: calls.append("login"))

    def load_trees():
        calls.append("trees")
        mapper._curriculum_index = {"A >> B": ["1"]}
        mapper._taxonomy_index = {"X >> Y": ["2"], "X >> Z": ["3"]}

    monkeypatch.setattr(mapper, "_ensure_trees", load_trees)

    assert mapper.test_connection() == {"curriculum_paths": 1, "taxonomy_paths": 2}
    assert calls == ["login", "trees"]
