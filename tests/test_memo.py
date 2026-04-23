"""Tests for skills/docs-memo/generate.py - create_sluzhebka()."""

import os

from conftest import memo_generate

create_sluzhebka = memo_generate.create_sluzhebka


class TestCreateSluzhebkaSmoke:
    """Smoke tests for memo generation."""

    def test_minimal(self, tmp_path):
        path = create_sluzhebka(
            body_paragraphs=["Прошу рассмотреть вопрос."],
            output_path=str(tmp_path / "memo.docx"),
        )
        assert os.path.isfile(path)
        assert path.endswith(".docx")

    def test_default_output_path(self, tmp_path, monkeypatch):
        monkeypatch.chdir(tmp_path)
        path = create_sluzhebka(
            body_paragraphs=["Текст записки."],
        )
        assert os.path.isfile(path)


class TestCreateSluzhebkaParams:
    """Test various parameter combinations."""

    def test_custom_date(self, tmp_path):
        path = create_sluzhebka(
            body_paragraphs=["Текст."],
            doc_date="15.03.2025",
            output_path=str(tmp_path / "date.docx"),
        )
        assert os.path.isfile(path)

    def test_custom_addressee_and_sender(self, tmp_path):
        path = create_sluzhebka(
            addressee="Заместителю директора\nТУ ЦТ\nПримерову П.П.",
            sender="от начальника отдела\nТестова Т.Т.",
            body_paragraphs=["Прошу согласовать."],
            signer="Тестов Т.Т.",
            output_path=str(tmp_path / "custom.docx"),
        )
        assert os.path.isfile(path)

    def test_multiple_body_paragraphs(self, tmp_path):
        path = create_sluzhebka(
            body_paragraphs=[
                "Довожу до Вашего сведения.",
                "Необходимо принять меры.",
                "Прошу рассмотреть и дать указания.",
            ],
            output_path=str(tmp_path / "multi.docx"),
        )
        assert os.path.isfile(path)

    def test_empty_org_config(self, tmp_path, monkeypatch):
        """Generator should work even with empty org config."""
        monkeypatch.setattr(memo_generate, "_load_org_details", lambda: {})
        path = create_sluzhebka(
            body_paragraphs=["Текст без конфига."],
            output_path=str(tmp_path / "empty_cfg.docx"),
        )
        assert os.path.isfile(path)
