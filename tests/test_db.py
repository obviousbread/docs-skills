"""Tests for lib/db.py - generation logging (JSONL)."""

import json
import os
import sys

import pytest

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(REPO_ROOT, "lib"))

from db import log_generation, _ensure_dir, DATA_DIR, LOG_PATH, EXAMPLES_DIR


class TestConstants:
    def test_data_dir_is_hidden(self):
        assert "/.docs-plugin" in DATA_DIR

    def test_examples_dir_inside_data_dir(self):
        assert EXAMPLES_DIR.startswith(DATA_DIR)

    def test_log_path_is_jsonl(self):
        assert LOG_PATH.endswith(".jsonl")


class TestEnsureDir:
    def test_creates_directory(self, tmp_path, monkeypatch):
        target = str(tmp_path / "sub" / "dir")
        monkeypatch.setattr("db.DATA_DIR", target)
        _ensure_dir()
        assert os.path.isdir(target)

    def test_idempotent(self, tmp_path, monkeypatch):
        monkeypatch.setattr("db.DATA_DIR", str(tmp_path))
        _ensure_dir()
        _ensure_dir()


class TestLogGeneration:
    def test_appends_line(self, tmp_path, monkeypatch):
        log_file = str(tmp_path / "generations.jsonl")
        monkeypatch.setattr("db.LOG_PATH", log_file)
        monkeypatch.setattr("db.DATA_DIR", str(tmp_path))
        log_generation("prikaz", "О тестировании", "/tmp/test.docx", {"points_count": 3})
        with open(log_file, encoding="utf-8") as f:
            lines = f.readlines()
        assert len(lines) == 1
        entry = json.loads(lines[0])
        assert entry["doc_type"] == "prikaz"
        assert entry["title"] == "О тестировании"
        assert entry["output_path"] == "/tmp/test.docx"
        assert entry["params"]["points_count"] == 3

    def test_multiple_appends(self, tmp_path, monkeypatch):
        log_file = str(tmp_path / "generations.jsonl")
        monkeypatch.setattr("db.LOG_PATH", log_file)
        monkeypatch.setattr("db.DATA_DIR", str(tmp_path))
        log_generation("prikaz", "First", "/tmp/1.docx")
        log_generation("memo", "Second", "/tmp/2.docx")
        with open(log_file, encoding="utf-8") as f:
            lines = f.readlines()
        assert len(lines) == 2
        assert json.loads(lines[0])["doc_type"] == "prikaz"
        assert json.loads(lines[1])["doc_type"] == "memo"

    def test_fail_safe(self, monkeypatch):
        monkeypatch.setattr("db.LOG_PATH", "/nonexistent/dir/gen.jsonl")
        monkeypatch.setattr("db.DATA_DIR", "/nonexistent/dir")
        log_generation("prikaz", "Тест", "/tmp/t.docx")

    def test_params_none(self, tmp_path, monkeypatch):
        log_file = str(tmp_path / "generations.jsonl")
        monkeypatch.setattr("db.LOG_PATH", log_file)
        monkeypatch.setattr("db.DATA_DIR", str(tmp_path))
        log_generation("memo", "Записка", "/tmp/m.docx")
        with open(log_file, encoding="utf-8") as f:
            entry = json.loads(f.readline())
        assert entry["params"] == {}

    def test_created_at_auto(self, tmp_path, monkeypatch):
        log_file = str(tmp_path / "generations.jsonl")
        monkeypatch.setattr("db.LOG_PATH", log_file)
        monkeypatch.setattr("db.DATA_DIR", str(tmp_path))
        log_generation("letter", "Письмо", "/tmp/l.docx")
        with open(log_file, encoding="utf-8") as f:
            entry = json.loads(f.readline())
        assert "created_at" in entry
        assert len(entry["created_at"]) >= 10

    def test_each_line_is_valid_json(self, tmp_path, monkeypatch):
        log_file = str(tmp_path / "generations.jsonl")
        monkeypatch.setattr("db.LOG_PATH", log_file)
        monkeypatch.setattr("db.DATA_DIR", str(tmp_path))
        log_generation("ord", "A", "/tmp/a.docx", {"k": "v"})
        log_generation("memo", "B", "/tmp/b.docx")
        with open(log_file, encoding="utf-8") as f:
            for line in f:
                json.loads(line)
