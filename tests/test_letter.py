"""Tests for skills/docs-letter/generate.py - create_letter()."""

import os

from conftest import letter_generate

create_letter = letter_generate.create_letter


class TestCreateLetterSmoke:
    """Smoke tests for letter generation."""

    def test_minimal(self, tmp_path):
        path = create_letter(
            addressee_lines=["Директору", "ООО «Тест»", "", "Тестову Т.Т."],
            body_paragraphs=["Направляем информацию."],
            output_path=str(tmp_path / "letter.docx"),
        )
        assert os.path.isfile(path)
        assert path.endswith(".docx")

    def test_default_output_path(self, tmp_path, monkeypatch):
        monkeypatch.chdir(tmp_path)
        path = create_letter(
            addressee_lines=["Директору", "ООО «Тест»"],
            body_paragraphs=["Текст письма."],
        )
        assert os.path.isfile(path)


class TestCreateLetterParams:
    """Test various parameter combinations."""

    def test_all_params(self, tmp_path):
        path = create_letter(
            addressee_lines=["Генеральному директору", "ФГБУ «Тест»", "", "Примерову П.П."],
            greeting="Уважаемый Тест Тестович!",
            body_paragraphs=[
                "Направляем документы.",
                "Просим рассмотреть в установленном порядке.",
            ],
            appendix_text="Документы на 5 л. в 1 экз.",
            signer_title="Генеральный директор",
            signer_name="Т.Т. Тестов",
            doc_date="15.03.2025",
            doc_number="01-01/123",
            on_number="N 32-024/13-3/1217 от 18.07.2025",
            executor_name="Тестов Тест Тестович",
            executor_phone="+7 (000) 000-00-00",
            output_path=str(tmp_path / "full.docx"),
        )
        assert os.path.isfile(path)

    def test_body_dict_format(self, tmp_path):
        path = create_letter(
            addressee_lines=["Директору", "ООО «Тест»"],
            body_paragraphs=[
                "Обычный абзац.",
                {"text": "Жирный заголовок", "bold": True, "indent": False, "center": True},
                {"text": "Нумерованный пункт 1.", "numbered": True},
                {"text": "Нумерованный пункт 2.", "numbered": True},
                {"text": "Буллет с тире.", "dashed": True},
            ],
            output_path=str(tmp_path / "dict_body.docx"),
        )
        assert os.path.isfile(path)

    def test_executor_in_body(self, tmp_path):
        path = create_letter(
            addressee_lines=["Директору", "ООО «Тест»"],
            body_paragraphs=["Текст письма."],
            executor_in_body=True,
            executor_name="Тестов Тест Тестович",
            executor_title="Начальник отдела",
            executor_phone="+7 (000) 000-00-00",
            output_path=str(tmp_path / "exec_body.docx"),
        )
        assert os.path.isfile(path)

    def test_empty_date_placeholder(self, tmp_path):
        """Empty doc_date should produce underscored placeholder."""
        path = create_letter(
            addressee_lines=["Директору", "ООО «Тест»"],
            body_paragraphs=["Текст."],
            doc_date="",
            output_path=str(tmp_path / "no_date.docx"),
        )
        assert os.path.isfile(path)

    def test_with_on_number(self, tmp_path):
        path = create_letter(
            addressee_lines=["Директору", "ООО «Тест»"],
            body_paragraphs=["Ответ на ваше обращение."],
            on_number="N 123 от 01.01.2025",
            output_path=str(tmp_path / "reply.docx"),
        )
        assert os.path.isfile(path)
