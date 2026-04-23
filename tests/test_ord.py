"""Tests for skills/docs-ord/generate.py - create_ord()."""

import os

from conftest import ord_generate

create_ord = ord_generate.create_ord


class TestCreateOrdSmoke:
    """Smoke tests: each doc type produces a valid .docx file."""

    def test_prikaz(self, tmp_path):
        path = create_ord(
            doc_type="prikaz",
            title="О тестировании",
            points=["Провести тестирование."],
            control_person="self",
            output_path=str(tmp_path / "prikaz.docx"),
        )
        assert os.path.isfile(path)
        assert path.endswith(".docx")

    def test_rasporyazhenie(self, tmp_path):
        path = create_ord(
            doc_type="rasporyazhenie",
            title="О проведении проверки",
            points=["Провести проверку."],
            control_person="self",
            output_path=str(tmp_path / "rasp.docx"),
        )
        assert os.path.isfile(path)

    def test_ukazanie(self, tmp_path):
        path = create_ord(
            doc_type="ukazanie",
            title="О направлении документов",
            points=["Направить документы."],
            control_person="self",
            output_path=str(tmp_path / "ukaz.docx"),
        )
        assert os.path.isfile(path)


class TestCreateOrdParams:
    """Test various parameter combinations."""

    def test_points_with_subpoints(self, tmp_path):
        path = create_ord(
            doc_type="prikaz",
            title="О назначении ответственных",
            points=[
                "Назначить ответственных:",
                ("Тестов Т.Т. - отдел тестирования.", 1),
                ("Примеров П.П. - отдел контроля.", 1),
                "Контроль исполнения возложить на заместителя директора.",
            ],
            control_person="self",
            output_path=str(tmp_path / "subpoints.docx"),
        )
        assert os.path.isfile(path)

    def test_with_basis(self, tmp_path):
        path = create_ord(
            doc_type="prikaz",
            title="Об утверждении положения",
            basis="В соответствии с Трудовым кодексом Российской Федерации",
            points=["Утвердить положение."],
            control_person="self",
            output_path=str(tmp_path / "basis.docx"),
        )
        assert os.path.isfile(path)

    def test_with_text_attachment(self, tmp_path):
        path = create_ord(
            doc_type="prikaz",
            title="Об утверждении плана",
            points=["Утвердить план (приложение)."],
            control_person="self",
            attachments=[{
                "title": "План мероприятий",
                "type": "text",
                "content": "Содержимое плана мероприятий.",
            }],
            output_path=str(tmp_path / "attach.docx"),
        )
        assert os.path.isfile(path)

    def test_with_commission_attachment(self, tmp_path):
        path = create_ord(
            doc_type="prikaz",
            title="О создании комиссии",
            points=["Создать комиссию (приложение)."],
            control_person="self",
            attachments=[{
                "title": "Состав комиссии",
                "type": "commission",
                "members": [
                    {"fio": "Тестов\nТест Тестович", "position": "Директор", "role": "Председатель"},
                    {"fio": "Примеров\nПример Примерович", "position": "Зам. директора", "role": None},
                ],
            }],
            output_path=str(tmp_path / "commission.docx"),
        )
        assert os.path.isfile(path)

    def test_with_notify_persons(self, tmp_path):
        path = create_ord(
            doc_type="prikaz",
            title="О тестировании уведомлений",
            points=["Провести тестирование."],
            control_person="self",
            notify_persons=[
                {"name": "Тестов Т.Т.", "position": "Начальник отдела"},
                {"name": "Примеров П.П.", "position": "Заместитель"},
            ],
            output_path=str(tmp_path / "notify.docx"),
        )
        assert os.path.isfile(path)

    def test_with_approved_by(self, tmp_path):
        path = create_ord(
            doc_type="prikaz",
            title="О согласовании",
            points=["Согласовать документ."],
            control_person="self",
            approved_by=[
                {"name": "Юристов Ю.Ю.", "position": "Главный юрисконсульт"},
            ],
            output_path=str(tmp_path / "approved.docx"),
        )
        assert os.path.isfile(path)

    def test_custom_date(self, tmp_path):
        path = create_ord(
            doc_type="prikaz",
            title="О тесте даты",
            points=["Тест."],
            control_person="self",
            doc_date="01.01.2025",
            output_path=str(tmp_path / "date.docx"),
        )
        assert os.path.isfile(path)

    def test_default_output_path(self, tmp_path, monkeypatch):
        monkeypatch.chdir(tmp_path)
        path = create_ord(
            doc_type="prikaz",
            title="Тестовый приказ",
            points=["Тест."],
            control_person="self",
        )
        assert os.path.isfile(path)
        assert path.endswith(".docx")
