"""Static quality checks for portable skill prompts."""

import re
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
LOCAL_CONFIG_REFS = {"org_details.md"}
INTENTIONAL_FUTURE_REFS = {"knowledge_base_medical.md"}


def _read(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def _markdown_files():
    yield from sorted((REPO_ROOT / "skills").glob("**/*.md"))
    yield from sorted((REPO_ROOT / "references").glob("*.md"))


def test_every_websearch_instruction_points_to_local_protocol():
    offenders = []
    for path in _markdown_files():
        text = _read(path)
        if "WebSearch" in text and "web-search.md" not in text:
            offenders.append(str(path.relative_to(REPO_ROOT)))
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


def test_all_generators_are_self_contained():
    offenders = []
    for path in sorted((REPO_ROOT / "skills").glob("docs-*/generate.py")):
        text = _read(path)
        if "def new_document" not in text or any(token in text for token in ("../../lib", ".docs-plugin/runtime", "from db import", "from docx_meta import")):
            offenders.append(str(path.relative_to(REPO_ROOT)))
    assert offenders == []


def test_document_skills_use_the_knowledge_base_contract():
    for name in ("docs-ord", "docs-letter", "docs-memo", "docs-di", "docs-protocol"):
        text = _read(REPO_ROOT / "skills" / name / "SKILL.md")
        assert "knowledge_base_path" in text
        assert "Читай столько" in text
        assert "корневые инструкции" in text


def test_org_template_contains_knowledge_base_path():
    template = REPO_ROOT / "skills" / "docs-init" / "references" / "org_details.md.example"
    assert "knowledge_base_path:" in _read(template)


def test_user_config_has_no_legacy_context_stores():
    offenders = []
    forbidden = (
        "~/.docs-plugin/di/approvers.md",
        "~/.docs-plugin/ord/scripts",
        "/examples/",
        "docs.db",
        "generations.jsonl",
    )
    for path in _markdown_files():
        text = _read(path)
        if any(token in text for token in forbidden):
            offenders.append(str(path.relative_to(REPO_ROOT)))
    assert offenders == []
