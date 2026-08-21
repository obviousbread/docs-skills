"""Checks for the two-command proofread workflow."""

import importlib.util
import json
import re
import sys
from pathlib import Path
from types import SimpleNamespace

import pytest
from docx import Document


REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT = REPO_ROOT / "skills" / "proofread" / "scripts" / "proofread_fast.py"


def _load_workflow():
    spec = importlib.util.spec_from_file_location("proofread_fast", SCRIPT)
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def _docx(path, text="Опечаткка."):
    document = Document()
    document.add_paragraph(text)
    document.save(path)


def _edits(path, old="Опечаткка", new="Опечатка"):
    path.write_text(json.dumps({
        "author": "Reviewer",
        "edits": [{"block": 1, "old": old, "new": new}],
    }), encoding="utf-8")


def test_intake_moves_external_docx_without_overwriting(tmp_path):
    workflow = _load_workflow()
    root = tmp_path / "work"
    (root / "ПРОЕКТЫ").mkdir(parents=True)
    source = tmp_path / "input.docx"
    _docx(source)

    moved, created = workflow.intake(source, root, "Proofread task")

    assert moved == root / "ПРОЕКТЫ" / "Proofread task" / "input.docx"
    assert created == moved.parent
    assert moved.is_file()
    assert not source.exists()
    duplicate = tmp_path / "input.docx"
    _docx(duplicate)
    with pytest.raises(SystemExit, match="Refusing to overwrite"):
        workflow.intake(duplicate, root, "Proofread task")
    assert duplicate.is_file()


def test_intake_without_projects_contract_leaves_source_in_place(tmp_path):
    workflow = _load_workflow()
    source = tmp_path / "input.docx"
    _docx(source)

    prepared, created = workflow.intake(source, None)

    assert prepared == source
    assert created is None
    assert source.is_file()


def test_finish_runs_visual_check_before_replacing_output(
    tmp_path, monkeypatch, capsys
):
    workflow = _load_workflow()
    source = tmp_path / "input.docx"
    output = tmp_path / "output.docx"
    edits = tmp_path / "edits.json"
    _docx(source)
    _edits(edits)

    def fake_render(render_source, temporary_output, final_output):
        assert render_source == source
        assert temporary_output.is_file()
        assert not final_output.exists()
        return tmp_path / ".proofread-qa", {"source": 1, "output": 1}

    monkeypatch.setattr(workflow, "render_pair", fake_render)
    workflow.finish(SimpleNamespace(
        input=str(source), edits=str(edits), output=str(output), visual=True
    ))

    result = json.loads(capsys.readouterr().out)
    assert output.is_file()
    assert result["visual_check"] is True
    assert result["pages"] == {"source": 1, "output": 1}
    assert result["edits"] == 1
    revisions = re.search(
        r"revision_runs=(\d+) yellow_runs=(\d+)", result["audit_result"]
    )
    assert revisions and revisions.group(1) == revisions.group(2) != "0"
    assert not list(tmp_path.glob(".*.proofread-fast.tmp.docx"))


def test_finish_skips_visual_check_by_default(tmp_path, monkeypatch, capsys):
    workflow = _load_workflow()
    source = tmp_path / "input.docx"
    output = tmp_path / "output.docx"
    edits = tmp_path / "edits.json"
    _docx(source)
    _edits(edits)

    def unexpected_render(*_args):
        raise AssertionError("render_pair must not run without --visual")

    monkeypatch.setattr(workflow, "render_pair", unexpected_render)
    workflow.finish(SimpleNamespace(
        input=str(source), edits=str(edits), output=str(output), visual=False
    ))

    result = json.loads(capsys.readouterr().out)
    assert output.is_file()
    assert result["visual_check"] is False
    assert result["qa_output"] is None
    assert result["pages"] is None
    assert result["page_count_changed"] is None
    assert result["seconds"]["render"] == 0.0


def test_finish_preserves_existing_output_when_preflight_fails(tmp_path):
    workflow = _load_workflow()
    source = tmp_path / "input.docx"
    output = tmp_path / "output.docx"
    edits = tmp_path / "edits.json"
    _docx(source)
    output.write_bytes(b"existing result")
    _edits(edits, old="missing text", new="replacement")

    with pytest.raises(SystemExit, match="output not written"):
        workflow.finish(SimpleNamespace(
            input=str(source), edits=str(edits), output=str(output), visual=False
        ))

    assert output.read_bytes() == b"existing result"
    assert not list(tmp_path.glob(".*.proofread-fast.tmp.docx"))
