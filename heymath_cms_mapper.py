"""Add curriculum and taxonomy paths to existing HeyMath CMS snippets.

This integration uses the CMS mapping webpages. It only adds missing mappings;
it never removes or replaces mappings already present in CMS.
"""

from __future__ import annotations

import html as _html
import logging
import re
import threading
import time
from dataclasses import dataclass, field
from typing import Iterable

import requests


log = logging.getLogger("heymath_cms_mapper")

SEP = " >> "
TREE_CACHE_SECONDS = 6 * 60 * 60
TIMEOUT = 60


def norm(path: str) -> str:
    parts = [re.sub(r"\s+", " ", part).strip() for part in str(path).split(">>")]
    return SEP.join(part for part in parts if part)


def _norm_all(paths: Iterable[str] | None) -> list[str]:
    return [norm(path) for path in (paths or []) if path and str(path).strip()]


@dataclass
class MappingResult:
    snippet_id: int
    question_name: str | None = None
    added: list[str] = field(default_factory=list)
    already_present: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)
    curriculum_before: list[str] = field(default_factory=list)
    taxonomy_before: list[str] = field(default_factory=list)
    curriculum_after: list[str] = field(default_factory=list)
    taxonomy_after: list[str] = field(default_factory=list)

    @property
    def success(self) -> bool:
        return not self.errors

    def summary(self) -> str:
        name = f" ({self.question_name})" if self.question_name else ""
        status = "OK" if self.success else "FAILED: " + "; ".join(self.errors)
        return f"Snippet {self.snippet_id}{name} {status} | added={self.added} | already={self.already_present}"


class HeyMathCmsMapper:
    """Thread-safe CMS mapper using one shared service account."""

    def __init__(self, base_url: str, login: str, password: str):
        self.base_url = base_url.rstrip("/")
        self._login = login
        self._password = password
        self._session = requests.Session()
        self._session.headers["User-Agent"] = "HeyMathCmsMapper/1.0 (python)"
        self._lock = threading.RLock()
        self._logged_in = False
        self._curriculum_index: dict[str, list[str]] | None = None
        self._taxonomy_index: dict[str, list[str]] | None = None
        self._trees_loaded_at = 0.0

    def map_question(
        self,
        snippet_id: int,
        curriculum_paths: Iterable[str] | None = None,
        taxonomy_paths: Iterable[str] | None = None,
    ) -> MappingResult:
        curriculum_paths = list(curriculum_paths or [])
        taxonomy_paths = list(taxonomy_paths or [])
        with self._lock:
            result = MappingResult(int(snippet_id))
            self._ensure_logged_in()
            self._ensure_trees()
            page = self._load_mapping_page(snippet_id)
            if page is None:
                result.errors.append(f"Question {snippet_id} was not found in CMS.")
                return result

            result.question_name = page["name"]
            result.curriculum_before = page["curriculum"]
            result.taxonomy_before = page["taxonomy"]
            curriculum_ids = self._resolve(
                curriculum_paths,
                self._curriculum_index,
                page["curriculum"],
                "Curriculum",
                result,
            )
            taxonomy_ids = self._resolve(
                taxonomy_paths,
                self._taxonomy_index,
                page["taxonomy"],
                "Taxonomy",
                result,
            )
            if result.errors:
                return result

            if curriculum_ids:
                self._post_form(
                    "/curriculum_trees/add_mapping",
                    page["curriculum_token"],
                    snippet_id,
                    "node_id",
                    ",".join(curriculum_ids),
                    "Add Mapping Form",
                )
            if taxonomy_ids:
                fresh = page if not curriculum_ids else self._load_mapping_page(snippet_id)
                if fresh is None:
                    result.errors.append(f"Question {snippet_id} could not be reloaded before taxonomy mapping.")
                    return result
                self._post_form(
                    "/heymath_taxonomy/add_mapping",
                    fresh["taxonomy_token"],
                    snippet_id,
                    "tax_node_id",
                    ",".join(taxonomy_ids),
                    "Add Taxonomy Form",
                )

            after = self._load_mapping_page(snippet_id)
            if after is None:
                result.errors.append(f"Question {snippet_id} could not be verified after mapping.")
                return result
            result.curriculum_after = after["curriculum"]
            result.taxonomy_after = after["taxonomy"]
            for path in _norm_all(curriculum_paths):
                if path not in after["curriculum"]:
                    result.errors.append(f"Curriculum mapping was not saved: {path}")
            for path in _norm_all(taxonomy_paths):
                if path not in after["taxonomy"]:
                    result.errors.append(f"Taxonomy mapping was not saved: {path}")
            for path in page["curriculum"]:
                if path not in after["curriculum"]:
                    result.errors.append(f"Existing curriculum mapping disappeared: {path}")
            for path in page["taxonomy"]:
                if path not in after["taxonomy"]:
                    result.errors.append(f"Existing taxonomy mapping disappeared: {path}")
            log.info(result.summary())
            return result

    def validate_path(self, path: str, *, taxonomy: bool = False) -> str | None:
        with self._lock:
            self._ensure_logged_in()
            self._ensure_trees()
            index = self._taxonomy_index if taxonomy else self._curriculum_index
            matches = index.get(norm(path)) if index else None
            if not matches:
                return f"Path not found: {path}"
            if len(matches) > 1:
                return f"Path is ambiguous ({len(matches)} nodes): {path}"
            return None

    def get_mappings(self, snippet_id: int) -> dict | None:
        """Return the current mappings without changing the CMS."""
        with self._lock:
            self._ensure_logged_in()
            page = self._load_mapping_page(snippet_id)
            if page is None:
                return None
            return {
                "name": page["name"],
                "curriculum": page["curriculum"],
                "taxonomy": page["taxonomy"],
            }

    def test_connection(self, *, refresh_trees: bool = True) -> dict[str, int]:
        """Log in and load both mapping trees without changing CMS content."""
        with self._lock:
            self._ensure_logged_in()
            if refresh_trees:
                self._trees_loaded_at = 0.0
            self._ensure_trees()
            return {
                "curriculum_paths": len(self._curriculum_index or {}),
                "taxonomy_paths": len(self._taxonomy_index or {}),
            }

    def refresh_trees(self) -> None:
        """Reload node trees on the next validation or mapping operation."""
        self._trees_loaded_at = 0.0

    def _load_mapping_page(self, snippet_id: int) -> dict | None:
        text = self._get(f"/question/{int(snippet_id)}/html_curriculum")
        if len(text) < 100 or "Browse Mapping Content" not in text:
            return None
        title = re.search(r"<title[^>]*>\s*Html Mapping-\s*(.*?)\s*</title>", text, re.S | re.I)
        curriculum_token = self._form_token(text, "/curriculum_trees/add_mapping")
        taxonomy_token = self._form_token(text, "/heymath_taxonomy/add_mapping")
        if not curriculum_token or not taxonomy_token:
            raise RuntimeError(f"Mapping forms were not found for snippet {snippet_id}; the CMS page may have changed.")

        body = re.sub(r"(?is)<script.*?</script>|<style.*?</style>", "", text)
        body = re.sub(r"<[^>]+>", "\n", body)
        lines = [re.sub(r"\s+", " ", _html.unescape(line)).strip() for line in body.split("\n")]
        lines = [line for line in lines if line]
        curriculum_heading = next((i for i, line in enumerate(lines) if line.endswith("Curriculum Mapping")), -1)
        taxonomy_heading = next((i for i, line in enumerate(lines) if line.endswith("Taxonomy Mapping")), -1)
        curriculum: list[str] = []
        taxonomy: list[str] = []
        for index, line in enumerate(lines):
            if ">>" not in line:
                continue
            if curriculum_heading >= 0 and index > curriculum_heading and (taxonomy_heading < 0 or index < taxonomy_heading):
                curriculum.append(norm(line))
            elif taxonomy_heading >= 0 and index > taxonomy_heading:
                taxonomy.append(norm(line))
        return {
            "name": _html.unescape(title.group(1)) if title else None,
            "curriculum_token": curriculum_token,
            "taxonomy_token": taxonomy_token,
            "curriculum": curriculum,
            "taxonomy": taxonomy,
        }

    @staticmethod
    def _form_token(text: str, action: str) -> str | None:
        form = re.search(
            r"""(?is)<form[^>]*\baction\s*=\s*["'][^"']*"""
            + re.escape(action)
            + r"""["'][^>]*>(.*?)</form>""",
            text,
        )
        if not form:
            return None
        token_input = re.search(
            r"""(?is)<input[^>]*\bname\s*=\s*["']authenticity_token["'][^>]*>""",
            form.group(1),
        )
        if not token_input:
            return None
        value = re.search(r"""\bvalue\s*=\s*["']([^"']*)["']""", token_input.group(0))
        return _html.unescape(value.group(1)) if value else None

    @staticmethod
    def _resolve(
        paths: Iterable[str],
        index: dict[str, list[str]] | None,
        current: list[str],
        label: str,
        result: MappingResult,
    ) -> list[str]:
        ids: list[str] = []
        index = index or {}
        for raw in paths:
            if not raw or not str(raw).strip():
                continue
            path = norm(raw)
            matches = index.get(path)
            if not matches:
                result.errors.append(f"{label} path not found in the CMS tree: {path}")
            elif len(matches) > 1:
                result.errors.append(f"{label} path matches {len(matches)} CMS nodes: {path}")
            elif path in current:
                result.already_present.append(f"{label}: {path}")
            elif matches[0] not in ids:
                ids.append(matches[0])
                result.added.append(f"{label}: {path}")
        return ids

    def _post_form(self, action: str, token: str, snippet_id: int, id_field: str, ids: str, button: str) -> None:
        referer = f"{self.base_url}/question/{int(snippet_id)}/html_curriculum"
        response = self._session.post(
            self.base_url + action,
            data={
                "authenticity_token": token,
                "project_id": str(int(snippet_id)),
                "project_type": "question",
                "parent_project_type": "",
                "parent_project_id": "",
                id_field: ids,
                "add_mapping": button,
            },
            headers={"Referer": referer},
            timeout=TIMEOUT,
            allow_redirects=True,
        )
        if response.status_code >= 400:
            raise RuntimeError(f"Save failed: HTTP {response.status_code} from {action}")
        log.info("CMS %s snippet=%s ids=%s -> HTTP %s", action, snippet_id, ids, response.status_code)

    def _ensure_trees(self) -> None:
        if self._curriculum_index is not None and time.time() - self._trees_loaded_at < TREE_CACHE_SECONDS:
            return
        curriculum_tree = self._get_json("/curriculum_trees/get_tree_json.json")
        taxonomy_tree = self._get_json("/heymath_taxonomy/get_tree_json.json")
        curriculum: dict[str, list[str]] = {}
        taxonomy: dict[str, list[str]] = {}
        curriculum_tops = curriculum_tree if isinstance(curriculum_tree, list) else [curriculum_tree]
        for top in curriculum_tops:
            for child in top.get("children") or []:
                self._index(child, "", curriculum)
        taxonomy_tops = taxonomy_tree if isinstance(taxonomy_tree, list) else (taxonomy_tree.get("children") or [])
        for top in taxonomy_tops:
            self._index(top, "", taxonomy)
        if not curriculum or not taxonomy:
            raise RuntimeError("The CMS curriculum or taxonomy tree could not be loaded.")
        self._curriculum_index = curriculum
        self._taxonomy_index = taxonomy
        self._trees_loaded_at = time.time()
        log.info("CMS trees loaded: %d curriculum paths, %d taxonomy paths", len(curriculum), len(taxonomy))

    @staticmethod
    def _index(node: dict, prefix: str, output: dict[str, list[str]]) -> None:
        stack = [(node, prefix)]
        while stack:
            current, current_prefix = stack.pop()
            text = _html.unescape(str(current.get("text", ""))).strip()
            path = f"{current_prefix}{SEP}{text}" if current_prefix else text
            output.setdefault(norm(path), []).append(str(current.get("id")))
            for child in current.get("children") or []:
                stack.append((child, path))

    def _ensure_logged_in(self) -> None:
        if self._logged_in:
            return
        self._session.cookies.clear()
        self._session.get(self.base_url + "/account/login", timeout=TIMEOUT)
        self._session.post(
            self.base_url + "/account/login",
            data={"login": self._login, "password": self._password},
            timeout=TIMEOUT,
        )
        home = self._session.get(self.base_url + "/user", timeout=TIMEOUT).text
        self._logged_in = "/account/logout" in home and 'name="password"' not in home
        if not self._logged_in:
            raise RuntimeError("CMS login failed; check the shared CMS account in Streamlit secrets.")

    @staticmethod
    def _looks_like_login(response: requests.Response) -> bool:
        return "/account/login" in response.url or ('name="password"' in response.text and "/account/login" in response.text)

    def _get_response(self, path: str) -> requests.Response:
        response = self._session.get(self.base_url + path, timeout=TIMEOUT)
        if self._looks_like_login(response):
            self._logged_in = False
            self._ensure_logged_in()
            response = self._session.get(self.base_url + path, timeout=TIMEOUT)
        response.raise_for_status()
        return response

    def _get(self, path: str) -> str:
        return self._get_response(path).text

    def _get_json(self, path: str):
        return self._get_response(path).json()
