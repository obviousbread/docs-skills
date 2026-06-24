"""Static quality checks for portable skill prompts."""

import re
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
LOCAL_CONFIG_REFS = {"org_details.md", "org_notes.md"}
INTENTIONAL_FUTURE_REFS = {"knowledge_base_medical.md"}


def _read(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def _markdown_files():
    yield from sorted((REPO_ROOT / "skills").glob("**/*.md"))
    yield from sorted((REPO_ROOT / "references").glob("*.md"))


def test_web_search_protocol_defines_fast_reliable_evidence_rules():
    text = _read(REPO_ROOT / "references" / "web-search.md")
    required = [
        "web-search.md",
        "WebSearch",
        "WebFetch",
        "не более 3 queries",
        "не более 2 fetched pages",
        "publication.pravo.gov.ru",
        "profstandart.rosmintrud.ru",
        "consultant.ru",
        "garant.ru",
        "source_url",
        "checked_at",
        "confidence",
        "не найден",
        "Не выдумывать",
    ]
    missing = [token for token in required if token not in text]
    assert not missing


def test_every_websearch_instruction_points_to_shared_protocol():
    offenders = []
    for path in _markdown_files():
        text = _read(path)
        if "WebSearch" in text and "web-search.md" not in text:
            offenders.append(str(path.relative_to(REPO_ROOT)))
    assert offenders == []


def test_runtime_reference_paths_exist_in_packaged_root_references():
    offenders = []
    for path in _markdown_files():
        text = _read(path)
        for rel in re.findall(r"~/.docs-plugin/runtime/(references/[^\s`)]+\.md(?:\.example)?)", text):
            if not (REPO_ROOT / rel).exists():
                offenders.append(f"{path.relative_to(REPO_ROOT)} -> {rel}")
    assert offenders == []


def test_backticked_markdown_references_resolve():
    offenders = []
    for path in _markdown_files():
        text = _read(path)
        base = path.parent
        rel_path = path.relative_to(REPO_ROOT)
        skill_root = REPO_ROOT / rel_path.parts[0] / rel_path.parts[1] if rel_path.parts[0] == "skills" else base
        for ref in re.findall(r"`([^`]+?\.md(?:\.example)?)`", text):
            if ref.startswith(("~", "http", "/")):
                continue
            if " " in ref:
                continue
            if Path(ref).name in LOCAL_CONFIG_REFS or Path(ref).name in INTENTIONAL_FUTURE_REFS:
                continue
            candidates = [
                base / ref,
                skill_root / "references" / ref,
                skill_root / ref,
                REPO_ROOT / "references" / ref,
                REPO_ROOT / ref,
            ]
            if not any(candidate.exists() for candidate in candidates):
                offenders.append(f"{path.relative_to(REPO_ROOT)} -> {ref}")
    assert offenders == []


def test_all_generators_use_shared_docx_meta_new_document():
    offenders = []
    for path in sorted((REPO_ROOT / "skills").glob("docs-*/generate.py")):
        text = _read(path)
        if "from docx_meta import new_document" not in text:
            offenders.append(str(path.relative_to(REPO_ROOT)))
    assert offenders == []
