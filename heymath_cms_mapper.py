"""Add, replace, or remove sheet-specified CMS mappings."""

from __future__ import annotations

import html as _html
from difflib import SequenceMatcher
import logging
import re
import threading
import time
from dataclasses import dataclass, field
from typing import Iterable
from urllib.parse import urljoin, urlparse

from lxml import html

import requests


log = logging.getLogger("heymath_cms_mapper")

SEP = " >> "
TREE_CACHE_SECONDS = 6 * 60 * 60
TIMEOUT = 60
MAPPER_BUILD_ID = "2026.10.12-checkbox-unmap-v4"


def norm(path: str) -> str:
    parts = [re.sub(r"\s+", " ", part).strip() for part in str(path).split(">>")]
    return SEP.join(part for part in parts if part)


def _norm_all(paths: Iterable[str] | None) -> list[str]:
    return [norm(path) for path in (paths or []) if path and str(path).strip()]


def _lookup_path(index: dict[str, list[str]], path: str) -> tuple[str, list[str]] | None:
    """Find a canonical CMS path without making letter case significant."""
    requested = norm(path).casefold()
    matching_keys = [candidate for candidate in index if candidate.casefold() == requested]
    if not matching_keys:
        return None
    ids: list[str] = []
    for key in matching_keys:
        for node_id in index[key]:
            if node_id not in ids:
                ids.append(node_id)
    return matching_keys[0], ids


def _contains_path(paths: Iterable[str], path: str) -> bool:
    requested = norm(path).casefold()
    return any(norm(candidate).casefold() == requested for candidate in paths)


@dataclass
class MappingResult:
    snippet_id: int
    question_name: str | None = None
    added: list[str] = field(default_factory=list)
    already_present: list[str] = field(default_factory=list)
    removed: list[str] = field(default_factory=list)
    already_absent: list[str] = field(default_factory=list)
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
        *,
        action: str = "add",
    ) -> MappingResult:
        if action not in {"add", "replace", "remove"}:
            raise ValueError("Mapping action must be add, replace, or remove.")
        curriculum_paths = _norm_all(curriculum_paths)
        taxonomy_paths = _norm_all(taxonomy_paths)
        if not curriculum_paths and not taxonomy_paths:
            return MappingResult(int(snippet_id))
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
                page["curriculum"] if action == "add" else [],
                "Curriculum",
                result,
            )
            taxonomy_ids = self._resolve(
                taxonomy_paths,
                self._taxonomy_index,
                page["taxonomy"] if action == "add" else [],
                "Taxonomy",
                result,
            )
            if result.errors:
                return result

            error = self.validate_action(page, curriculum_paths, taxonomy_paths, action=action)
            if error:
                result.errors.append(error)
                return result
            if action != "add":
                result.added.clear()
                for label, key, requested in (
                    ("Curriculum", "curriculum", curriculum_paths),
                    ("Taxonomy", "taxonomy", taxonomy_paths),
                ):
                    targets = page[key] if action == "replace" else requested
                    for path in targets:
                        if not _contains_path(page[key], path):
                            result.already_absent.append(f"{label}: {path}")
                            continue
                        fresh = self._load_mapping_page(snippet_id)
                        if fresh is None:
                            raise RuntimeError("Question could not be reloaded before removal.")
                        self._remove_mapping(fresh, key, path, snippet_id)
                        verified = self._load_mapping_page(snippet_id)
                        if verified is None or _contains_path(verified[key], path):
                            result.errors.append(f"{label} mapping was not removed: {path}")
                            return result
                        result.removed.append(f"{label}: {path}")
                page = self._load_mapping_page(snippet_id)
                if page is None:
                    raise RuntimeError("Question could not be reloaded after removal.")
                if action == "remove":
                    curriculum_ids = taxonomy_ids = []
                else:
                    result.added = [f"Curriculum: {path}" for path in curriculum_paths] + [
                        f"Taxonomy: {path}" for path in taxonomy_paths
                    ]

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
            for path in curriculum_paths:
                if action != "remove" and not _contains_path(after["curriculum"], path):
                    result.errors.append(f"Curriculum mapping was not saved: {path}")
            for path in taxonomy_paths:
                if action != "remove" and not _contains_path(after["taxonomy"], path):
                    result.errors.append(f"Taxonomy mapping was not saved: {path}")
            for path in result.curriculum_before:
                if action == "replace" or (action == "remove" and _contains_path(curriculum_paths, path)):
                    continue
                if not _contains_path(after["curriculum"], path):
                    result.errors.append(f"Existing curriculum mapping disappeared: {path}")
            for path in result.taxonomy_before:
                if action == "replace" or (action == "remove" and _contains_path(taxonomy_paths, path)):
                    continue
                if not _contains_path(after["taxonomy"], path):
                    result.errors.append(f"Existing taxonomy mapping disappeared: {path}")
            for key, requested in (("curriculum", curriculum_paths), ("taxonomy", taxonomy_paths)):
                if action == "replace" and {norm(p).casefold() for p in after[key]} != {p.casefold() for p in requested}:
                    result.errors.append(f"Final {key} mappings do not match the sheet.")
                if action == "remove" and any(_contains_path(after[key], p) for p in requested):
                    result.errors.append(f"Some requested {key} mappings remain.")
            log.info(result.summary())
            return result

    def validate_action(self, page: dict, curriculum_paths: Iterable[str], taxonomy_paths: Iterable[str], *, action: str) -> str | None:
        """Check every required removal control before changing any mappings."""
        if action == "add":
            return None
        for key, requested in (("curriculum", curriculum_paths), ("taxonomy", taxonomy_paths)):
            targets = page[key] if action == "replace" else _norm_all(requested)
            for path in targets:
                if _contains_path(page[key], path) and not self._removal_control(page, key, path):
                    count = sum(c["type"] == key for c in page.get("removal_controls", []))
                    return (f"CMS removal control not found for {key}: {path}. "
                            f"Found {count} {key} removal controls for {len(page[key])} existing mappings "
                            f"(mapping build {MAPPER_BUILD_ID}). No changes were made to this question.")
        return None

    def validate_question_action(self, snippet_id: int, curriculum_paths: Iterable[str], taxonomy_paths: Iterable[str], *, action: str) -> str | None:
        """Read and validate full CMS forms without exposing form tokens to the UI."""
        if action not in {"add", "replace", "remove"}:
            raise ValueError("Mapping action must be add, replace, or remove.")
        if action == "add":
            return None
        with self._lock:
            self._ensure_logged_in()
            page = self._load_mapping_page(snippet_id)
            if page is None:
                return f"Question {snippet_id} was not found in CMS."
            return self.validate_action(page, curriculum_paths, taxonomy_paths, action=action)

    @staticmethod
    def _removal_control(page: dict, key: str, path: str) -> dict | None:
        matches = [c for c in page.get("removal_controls", []) if c["type"] == key and _contains_path([c["path"]], path)]
        return matches[0] if len(matches) == 1 else None

    def _remove_mapping(self, page: dict, key: str, path: str, snippet_id: int) -> None:
        control = self._removal_control(page, key, path)
        if control is None:
            raise RuntimeError(f"CMS removal control not found for {key}: {path}")
        url = urljoin(self.base_url + "/", control["url"])
        if (urlparse(url).scheme, urlparse(url).netloc) != (urlparse(self.base_url).scheme, urlparse(self.base_url).netloc):
            raise RuntimeError("CMS removal control points outside the configured CMS.")
        response = self._session.request(control["method"], url, data=control["data"],
            headers={"Referer": f"{self.base_url}/question/{int(snippet_id)}/html_curriculum"},
            timeout=TIMEOUT, allow_redirects=True)
        response.raise_for_status()
        if self._looks_like_login(response):
            raise RuntimeError("CMS session expired during removal. Review current mappings before retrying.")

    def validate_path(self, path: str, *, taxonomy: bool = False) -> str | None:
        with self._lock:
            self._ensure_logged_in()
            self._ensure_trees()
            index = self._taxonomy_index if taxonomy else self._curriculum_index
            resolved = _lookup_path(index or {}, path)
            if resolved is None:
                candidates = self._suggest_paths(path, index or {})
                suggestion = f" Closest CMS matches: {'; '.join(candidates)}." if candidates else ""
                return f"Path not found: {path}.{suggestion}"
            _, matches = resolved
            if len(matches) > 1:
                return f"Path is ambiguous ({len(matches)} nodes): {path}"
            return None

    @staticmethod
    def _suggest_paths(path: str, index: dict[str, list[str]], limit: int = 3) -> list[str]:
        """Return hierarchy-aware suggestions without treating them as approved mappings."""
        requested = norm(path)
        requested_parts = [part.casefold() for part in requested.split(SEP)]
        if not requested_parts:
            return []

        ranked: list[tuple[float, int, str]] = []
        for candidate in index:
            candidate_parts = [part.casefold() for part in candidate.split(SEP)]
            common_prefix = 0
            for requested_part, candidate_part in zip(requested_parts, candidate_parts):
                if requested_part != candidate_part:
                    break
                common_prefix += 1

            whole_similarity = SequenceMatcher(
                None, requested.casefold(), candidate.casefold()
            ).ratio()
            leaf_similarity = SequenceMatcher(
                None, requested_parts[-1], candidate_parts[-1]
            ).ratio()
            shared_branch = common_prefix >= min(2, len(requested_parts), len(candidate_parts))

            # Prefer paths in the same CMS branch, while still allowing a very close
            # spelling correction when the mismatch occurs near the root.
            score = (
                whole_similarity
                + (0.20 if shared_branch else 0.0)
                + 0.10 * leaf_similarity
                - 0.025 * abs(len(requested_parts) - len(candidate_parts))
            )
            if shared_branch or whole_similarity >= 0.62:
                ranked.append((score, common_prefix, candidate))

        ranked.sort(key=lambda item: (-item[0], -item[1], item[2].casefold()))
        return [candidate for _, _, candidate in ranked[:limit]]

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

    def resolve_snippet_id(self, question_name: str) -> tuple[int | None, str | None]:
        """Resolve an exact CMS question name such as project10433_q71 to its snippet ID."""
        requested = re.sub(r"\s+", " ", str(question_name)).strip()
        if not requested:
            return None, "A Project with qno value is blank."
        with self._lock:
            self._ensure_logged_in()
            candidates: list[str] = []
            for search_term in dict.fromkeys((requested, requested.casefold())):
                token = self._site_authenticity_token()
                response = self._session.post(
                    self.base_url + "/site/search",
                    data={
                        "authenticity_token": token,
                        "searchterms": search_term,
                        "content_type": "Question",
                    },
                    timeout=TIMEOUT,
                )
                if self._looks_like_login(response):
                    self._logged_in = False
                    self._ensure_logged_in()
                    token = self._site_authenticity_token()
                    response = self._session.post(
                        self.base_url + "/site/search",
                        data={
                            "authenticity_token": token,
                            "searchterms": search_term,
                            "content_type": "Question",
                        },
                        timeout=TIMEOUT,
                    )
                response.raise_for_status()
                for candidate in re.findall(r'''href=["']/question/(\d+)/''', response.text, re.I):
                    if candidate not in candidates:
                        candidates.append(candidate)
            exact: list[int] = []
            for candidate in candidates[:50]:
                page = self._load_mapping_page(int(candidate))
                candidate_name = re.sub(r"\s+", " ", str((page or {}).get("name") or "")).strip()
                if candidate_name.casefold() == requested.casefold():
                    exact.append(int(candidate))
            if not exact:
                return None, f"No exact CMS question named {requested!r} was found."
            if len(exact) > 1:
                return None, f"CMS contains more than one exact question named {requested!r}: {', '.join(map(str, exact))}."
            return exact[0], None

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
            "removal_controls": self._parse_removal_controls(text, curriculum, taxonomy),
        }

    @staticmethod
    def _parse_removal_controls(text: str, curriculum: list[str], taxonomy: list[str]) -> list[dict]:
        """Read explicit per-path removal forms/links; never invent CMS endpoints."""
        root = html.fromstring(text)
        meta = root.xpath('//meta[@name="csrf-token"]/@content')
        token = meta[0] if meta else None
        controls = []
        # CMS UnMap submits mapping-record IDs, not curriculum/taxonomy tree IDs.
        # Curriculum checkboxes sit inside spans; taxonomy checkboxes are direct
        # children of the form. In both cases the path follows up to the next br.
        # Parse each source form independently. Repairing the entire legacy CMS
        # document can discard nested form tags or move inputs outside a form.
        source_forms = re.findall(r"(?is)<form\b[^>]*>.*?</form\s*>", text)
        for source_form in source_forms:
            # Start at the innermost opening form if the source nests forms.
            openings = list(re.finditer(r"(?is)<form\b", source_form))
            form = html.fromstring(source_form[openings[-1].start():])
            endpoint = urlparse(form.get("action", "")).path
            key = {"/curriculum_trees/remove_mapping": "curriculum",
                   "/heymath_taxonomy/remove_mapping": "taxonomy"}.get(endpoint)
            if key is None:
                continue
            base_data = {field.get("name"): field.get("value", "")
                         for field in form.xpath('.//input[@name][@type="hidden"]')}
            for submit in form.xpath('.//input[@type="submit"][@name]'):
                if submit.get("value", "").casefold() == "unmap":
                    base_data[submit.get("name")] = submit.get("value")
            expected_prefix = "curriculum_tree_ids[" if key == "curriculum" else "heymath_taxonomy_ids["
            for checkbox in form.xpath('.//input[@type="checkbox"][@name]'):
                if not checkbox.get("name", "").startswith(expected_prefix):
                    continue
                path_text = checkbox.tail or ""
                for sibling in checkbox.itersiblings():
                    if sibling.tag in {"br", "input"}:
                        break
                    path_text += sibling.text_content() + (sibling.tail or "")
                path = norm(path_text)
                paths = curriculum if key == "curriculum" else taxonomy
                if not _contains_path(paths, path):
                    continue
                controls.append({"type": key, "path": path,
                    "method": form.get("method", "post").upper(), "url": form.get("action"),
                    "data": {**base_data, checkbox.get("name"): checkbox.get("value", "")}})
        for element in root.xpath('//a[@href] | //form[@action]'):
            if element.tag == "form" and element.xpath('.//input[@type="checkbox"]'):
                continue
            marker = " ".join([element.text_content(), element.get("title", ""),
                element.get("aria-label", ""), element.get("href", ""), element.get("action", "")])
            if not re.search(r"\b(remove|delete|unmap|destroy)\b|remove_mapping|delete_mapping", marker, re.I):
                continue
            container = element
            matches = []
            # Require exactly one known mapped path within the nearest enclosing row/item.
            for _ in range(5):
                visible = norm(container.text_content())
                matches = [(key, path) for key, paths in (("curriculum", curriculum), ("taxonomy", taxonomy))
                    for path in paths if norm(path).casefold() in visible.casefold()]
                if matches:
                    break
                container = container.getparent()
                if container is None:
                    break
            if len(matches) != 1:
                continue
            data = {}
            if element.tag == "form":
                for field in element.xpath('.//input[@name] | .//button[@name]'):
                    if field.get("type", "").lower() not in {"checkbox", "radio"} or field.get("checked") is not None:
                        data[field.get("name")] = field.get("value", "")
                method = element.get("method", "get").upper()
                url = element.get("action")
            else:
                method = element.get("data-method", element.get("data-turbo-method", "get")).upper()
                url = element.get("href")
                if method != "GET":
                    if not token:
                        continue
                    data = {"authenticity_token": token}
            if not url or url.startswith(("#", "javascript:")) or method not in {"GET", "POST", "DELETE", "PATCH"}:
                continue
            endpoint = urlparse(url).path
            if not endpoint.startswith(("/curriculum_trees/", "/heymath_taxonomy/")):
                continue
            if method == "GET" and data:
                continue
            controls.append({"type": matches[0][0], "path": matches[0][1], "method": method, "url": url, "data": data})
        return controls

    def _site_authenticity_token(self) -> str:
        text = self._get("/user")
        input_match = re.search(
            r'''(?is)<input[^>]*name=["']authenticity_token["'][^>]*>''',
            text,
        )
        if input_match:
            value = re.search(r'''value=["']([^"']+)["']''', input_match.group(0), re.I)
            if value:
                return _html.unescape(value.group(1))
        meta = re.search(
            r'''(?is)<meta[^>]*name=["']csrf-token["'][^>]*content=["']([^"']+)["']''',
            text,
        )
        if meta:
            return _html.unescape(meta.group(1))
        raise RuntimeError("The CMS search token could not be found; the CMS page may have changed.")

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
            requested_path = norm(raw)
            resolved = _lookup_path(index, requested_path)
            if resolved is None:
                path = requested_path
                result.errors.append(f"{label} path not found in the CMS tree: {path}")
                continue
            path, matches = resolved
            if len(matches) > 1:
                result.errors.append(f"{label} path matches {len(matches)} CMS nodes: {path}")
            elif _contains_path(current, path):
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
