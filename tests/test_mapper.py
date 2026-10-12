from heymath_cms_mapper import HeyMathCmsMapper
import pytest


def test_q71_parent_level_mapping_has_unmap_control():
    paths = ["CBSE NCERT >> Class 4 >> Mathematics >> Measuring Length",
             "India >> Class 4 >> Mathematics >> Measurement"]
    text = '''<form action="/curriculum_trees/remove_mapping" method="post">
      <input name="authenticity_token" type="hidden" value="test">
      <input name="project_id" type="hidden" value="219263">
      <span><input id="curriculum_tree_ids_716223" name="curriculum_tree_ids[716223]" type="checkbox" value="716223">
      CBSE NCERT &gt;&gt; Class 4 &gt;&gt; Mathematics &gt;&gt; Measuring Length <br></span>
      <span><input id="curriculum_tree_ids_716224" name="curriculum_tree_ids[716224]" type="checkbox" value="716224">
      India &gt;&gt; Class 4 &gt;&gt; Mathematics &gt;&gt; Measurement <br></span>
      <input type="submit" name="UnMap" value="UnMap"></form>'''
    mapper = HeyMathCmsMapper("http://cms.heymath.com", "unused", "unused")
    controls = mapper._parse_removal_controls(text, paths, [])
    assert len(controls) == 2
    page = {"curriculum": paths, "taxonomy": [], "removal_controls": controls}
    assert mapper.validate_action(page, paths, [], action="replace") is None
    assert controls[0]["data"]["curriculum_tree_ids[716223]"] == "716223"


def test_real_cms_unmap_forms_select_only_the_requested_mapping_record(monkeypatch):
    # Structure observed on project10433_q103: curriculum spans, taxonomy
    # checkbox directly in form, shared UnMap submit, mapping-record IDs.
    curriculum = [
        "CBSE NCERT >> Class 4 >> Mathematics >> Measuring Length >> Perimeter",
        "India >> Class 4 >> Mathematics >> Area and Perimeter >> Perimeter of Simple Figures",
    ]
    taxonomy = ["Mathematics >> Measurement >> Length >> Perimeter"]
    text = '''<html><form action="/curriculum_trees/remove_mapping" method="post">
      <input name="authenticity_token" type="hidden" value="c-token">
      <input name="project_id" type="hidden" value="219295">
      <input name="project_type" type="hidden" value="question">
      <input name="parent_project_type" type="hidden" value="">
      <input name="parent_project_id" type="hidden" value="">
      <span><input type="checkbox" name="curriculum_tree_ids[716309]" value="716309">
      CBSE NCERT &gt;&gt; Class 4 &gt;&gt; Mathematics &gt;&gt; Measuring Length &gt;&gt; Perimeter<br></span>
      <span><input type="checkbox" name="curriculum_tree_ids[716310]" value="716310" checked>
      India &gt;&gt; Class 4 &gt;&gt; Mathematics &gt;&gt; Area and Perimeter &gt;&gt; Perimeter of Simple Figures<br></span>
      <input type="submit" name="UnMap" value="UnMap"></form>
      <form action="/heymath_taxonomy/remove_mapping" method="post">
      <input name="authenticity_token" type="hidden" value="t-token">
      <input name="project_id" type="hidden" value="219295">
      <input name="project_type" type="hidden" value="question">
      <input name="parent_project_type" type="hidden" value="">
      <input name="parent_project_id" type="hidden" value="">
      <input type="checkbox" name="heymath_taxonomy_ids[346644]" value="346644">
      Mathematics &gt;&gt; Measurement &gt;&gt; Length &gt;&gt; Perimeter<br>
      <input type="submit" name="UnMap" value="UnMap"></form></html>'''
    mapper = HeyMathCmsMapper("http://cms.heymath.com", "unused", "unused")
    controls = mapper._parse_removal_controls(text, curriculum, taxonomy)
    assert len(controls) == 3
    page = {"curriculum": curriculum, "taxonomy": taxonomy, "removal_controls": controls}
    assert mapper.validate_action(page, curriculum, taxonomy, action="replace") is None
    assert mapper.validate_action(page, curriculum[:1], [], action="remove") is None
    captured = []
    class Response:
        url = "http://cms.heymath.com/question/219295/html_curriculum"
        text = "Browse Mapping Content"
        def raise_for_status(self):
            pass
    monkeypatch.setattr(mapper._session, "request", lambda *args, **kwargs: captured.append((args, kwargs)) or Response())
    mapper._remove_mapping(page, "curriculum", curriculum[0], 219295)
    mapper._remove_mapping(page, "taxonomy", taxonomy[0], 219295)
    data = captured[0][1]["data"]
    assert data["curriculum_tree_ids[716309]"] == "716309"
    assert "curriculum_tree_ids[716310]" not in data
    assert data["project_id"] == "219295"
    assert data["authenticity_token"] == "c-token"
    assert data["UnMap"] == "UnMap"
    assert captured[1][1]["data"]["heymath_taxonomy_ids[346644]"] == "346644"
    assert captured[0][0] == ("POST", "http://cms.heymath.com/curriculum_trees/remove_mapping")


@pytest.mark.parametrize("action", ["replace", "remove"])
def test_destructive_actions_validate_before_removing(monkeypatch, action):
    mapper = HeyMathCmsMapper("https://cms.example", "shared", "secret")
    mapper._logged_in = True
    mapper._curriculum_index = {"Existing >> Path": ["10"]}
    mapper._taxonomy_index = {}
    mapper._trees_loaded_at = 10**12
    page = {"name": "Q", "curriculum": ["Existing >> Path"], "taxonomy": []}
    monkeypatch.setattr(mapper, "_load_mapping_page", lambda _: page)
    monkeypatch.setattr(mapper, "_remove_mapping", lambda *args: pytest.fail("Must not remove"))
    result = mapper.map_question(123, ["Unknown >> Path"], action=action)
    assert not result.success
    result = mapper.map_question(123, ["Existing >> Path"], action=action)
    assert not result.success
    assert "removal control not found" in result.errors[0]


@pytest.mark.parametrize("action", ["add", "replace", "remove"])
def test_blank_paths_skip_question(monkeypatch, action):
    mapper = HeyMathCmsMapper("https://cms.example", "shared", "secret")
    monkeypatch.setattr(mapper, "_ensure_logged_in", lambda: pytest.fail("Blank input must be skipped"))
    assert mapper.map_question(123, ["", " "], [], action=action).success


@pytest.mark.parametrize("action", ["replace", "remove"])
def test_replace_and_remove_verify_final_state(monkeypatch, action):
    mapper = HeyMathCmsMapper("https://cms.example", "shared", "secret")
    mapper._logged_in = True
    mapper._curriculum_index = {"Old >> Curriculum": ["10"], "New >> Curriculum": ["11"]}
    mapper._taxonomy_index = {"Old >> Taxonomy": ["20"]}
    mapper._trees_loaded_at = 10**12
    state = {"name": "Q", "curriculum_token": "c", "taxonomy_token": "t",
             "curriculum": ["Old >> Curriculum"], "taxonomy": ["Old >> Taxonomy"]}
    def load(_):
        return {**state, "curriculum": list(state["curriculum"]), "taxonomy": list(state["taxonomy"]),
                "removal_controls": [{"type": key, "path": p} for key in ("curriculum", "taxonomy") for p in state[key]]}
    monkeypatch.setattr(mapper, "_load_mapping_page", load)
    monkeypatch.setattr(mapper, "_remove_mapping", lambda page, key, path, _: state[key].remove(path))
    monkeypatch.setattr(mapper, "_post_form", lambda *args: state["curriculum"].append("New >> Curriculum"))
    paths = ["New >> Curriculum"] if action == "replace" else ["Old >> Curriculum"]
    result = mapper.map_question(123, paths, action=action)
    assert result.success, result.errors
    assert result.curriculum_after == (["New >> Curriculum"] if action == "replace" else [])
    assert result.taxonomy_after == ([] if action == "replace" else ["Old >> Taxonomy"])
    assert len(result.removed) == (2 if action == "replace" else 1)


def test_removal_controls_are_read_from_explicit_per_path_controls():
    text = '''<html><meta name="csrf-token" content="token"><table>
    <tr><td>Old &gt;&gt; Curriculum</td><td><a href="/curriculum_trees/remove_mapping/7" data-method="delete">Remove</a></td></tr>
    <tr><td>Old &gt;&gt; Taxonomy</td><td><form action="/heymath_taxonomy/remove_mapping" method="post">
    <input name="authenticity_token" value="tax"><input name="id" value="9"><button>Remove</button></form></td></tr>
    </table></html>'''
    controls = HeyMathCmsMapper._parse_removal_controls(text, ["Old >> Curriculum"], ["Old >> Taxonomy"])
    assert len(controls) == 2
    assert controls[0]["method"] == "DELETE"
    assert controls[0]["data"] == {"authenticity_token": "token"}
    assert controls[1]["data"]["id"] == "9"


def test_failed_removal_stops_before_adding(monkeypatch):
    mapper = HeyMathCmsMapper("https://cms.example", "shared", "secret")
    mapper._logged_in = True
    mapper._trees_loaded_at = 10**12
    mapper._curriculum_index = {"Old >> Path": ["1"], "New >> Path": ["2"]}
    mapper._taxonomy_index = {}
    page = {"name": "Q", "curriculum": ["Old >> Path"], "taxonomy": [],
            "removal_controls": [{"type": "curriculum", "path": "Old >> Path"}]}
    monkeypatch.setattr(mapper, "_load_mapping_page", lambda _: page)
    monkeypatch.setattr(mapper, "_remove_mapping", lambda *args: None)
    monkeypatch.setattr(mapper, "_post_form", lambda *args: pytest.fail("Must not add after failed removal"))
    result = mapper.map_question(123, ["New >> Path"], action="replace")
    assert not result.success
    assert "not removed" in result.errors[0]


def test_unmap_absent_path_is_a_verified_noop(monkeypatch):
    mapper = HeyMathCmsMapper("https://cms.example", "shared", "secret")
    mapper._logged_in = True
    mapper._trees_loaded_at = 10**12
    mapper._curriculum_index = {"Absent >> Path": ["1"]}
    mapper._taxonomy_index = {}
    page = {"name": "Q", "curriculum": [], "taxonomy": []}
    monkeypatch.setattr(mapper, "_load_mapping_page", lambda _: page)
    result = mapper.map_question(123, ["Absent >> Path"], action="remove")
    assert result.success
    assert result.already_absent == ["Curriculum: Absent >> Path"]
    assert result.added == result.removed == []


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
        ["existing>>curriculum", "NEW >> curriculum"],
        ["EXISTING >> taxonomy", "new>>TAXONOMY"],
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
    assert "Closest CMS matches: Mathematics >> Measurement >> Measuring Temperature." in error


def test_validate_path_is_case_insensitive():
    mapper = HeyMathCmsMapper("https://cms.example", "shared", "secret")
    mapper._curriculum_index = {
        "India >> Class 4 >> Mathematics >> Measurement": ["10"]
    }
    mapper._taxonomy_index = {"Mathematics >> Measurement >> Lengths": ["20"]}
    mapper._trees_loaded_at = 10**12
    mapper._logged_in = True

    assert mapper.validate_path(
        "mathematics >> MEASUREMENT >> lengths", taxonomy=True
    ) is None
    assert mapper.validate_path(
        "india >> class 4 >> mathematics >> measurement"
    ) is None


def test_validate_path_suggests_same_branch_when_leaf_is_different():
    mapper = HeyMathCmsMapper("https://cms.example", "shared", "secret")
    mapper._curriculum_index = {"Known >> Path": ["10"]}
    mapper._taxonomy_index = {
        "Mathematics >> Measurement >> Length >> Measuring Distance": ["20"],
        "Mathematics >> Measurement >> Length >> Perimeter": ["21"],
        "Mathematics >> Geometry >> Circles": ["22"],
    }
    mapper._trees_loaded_at = 10**12
    mapper._logged_in = True

    error = mapper.validate_path(
        "Mathematics >> Measurement >> Lengths >> Distance",
        taxonomy=True,
    )

    assert error is not None
    assert "Mathematics >> Measurement >> Length >> Measuring Distance" in error
    assert "Mathematics >> Geometry >> Circles" not in error


def test_validate_path_omits_unreliable_suggestions():
    mapper = HeyMathCmsMapper("https://cms.example", "shared", "secret")
    mapper._curriculum_index = {"CBSE NCERT >> Class 4 >> Mathematics": ["10"]}
    mapper._taxonomy_index = {"Language >> Grammar >> Nouns": ["20"]}
    mapper._trees_loaded_at = 10**12
    mapper._logged_in = True

    error = mapper.validate_path("Completely unrelated value", taxonomy=True)

    assert error is not None
    assert "Closest CMS matches" not in error


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

    search_terms = []

    def fake_post(*args, **kwargs):
        search_terms.append(kwargs["data"]["searchterms"])
        return Response()

    monkeypatch.setattr(mapper._session, "post", fake_post)
    monkeypatch.setattr(
        mapper,
        "_load_mapping_page",
        lambda snippet_id: {"name": "project10433_q71" if snippet_id == 219263 else "project10433_q710"},
    )

    snippet_id, error = mapper.resolve_snippet_id("PROJECT10433_Q71")

    assert snippet_id == 219263
    assert error is None
    assert "project10433_q71" in search_terms


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
