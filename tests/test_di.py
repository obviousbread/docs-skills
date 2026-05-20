"""Tests for skills/docs-di/generate.py - create_di()."""

import os

from conftest import di_generate

create_di = di_generate.create_di


SAMPLE_SECTIONS = [
    {
        "title": "Общие положения",
        "points": [
            "Настоящая должностная инструкция определяет функциональные "
            "обязанности, права и ответственность специалиста.",
        ],
    },
    {
        "title": "Трудовая функция",
        "intro": "Специалист осуществляет следующие функции:",
        "points": ["Выполнение поручений руководителя."],
    },
    {
        "title": "Должностные обязанности",
        "points": ["Выполняет поручения непосредственного руководителя."],
    },
    {
        "title": "Права",
        "points": ["Запрашивать информационные материалы."],
    },
    {
        "title": "Ответственность",
        "points": [
            "неисполнение обязанностей в пределах трудового "
            "законодательства Российской Федерации.",
        ],
    },
    {
        "title": "Заключительные положения",
        "points": ["Должностная инструкция действует бессрочно."],
    },
]


class TestCreateDiSmoke:
    def test_minimal(self, tmp_path):
        path = create_di(
            position_full="специалиста тестового отдела",
            position_full_genitive="специалиста тестового отдела",
            sections=SAMPLE_SECTIONS,
            output_path=str(tmp_path / "di.docx"),
        )
        assert os.path.isfile(path)
        assert path.endswith(".docx")

    def test_default_output_path(self, tmp_path, monkeypatch):
        monkeypatch.chdir(tmp_path)
        path = create_di(
            position_full="специалиста тестового отдела",
            position_full_genitive="специалиста тестового отдела",
            sections=SAMPLE_SECTIONS,
        )
        assert os.path.isfile(path)


class TestCreateDiParams:
    def test_with_approvers(self, tmp_path):
        path = create_di(
            position_full="начальника тестового отдела",
            position_full_genitive="начальника тестового отдела",
            sections=SAMPLE_SECTIONS,
            output_path=str(tmp_path / "dept_head.docx"),
            approvers=[
                {"position": "Заместитель директора", "name": "Примеров П.П."},
                {"position": "Главный юрисконсульт", "name": "Иванов И.И."},
            ],
        )
        assert os.path.isfile(path)

    def test_with_dash_items(self, tmp_path):
        sections = SAMPLE_SECTIONS[:1] + [
            {
                "title": "Трудовая функция",
                "points": [
                    {
                        "text": "На должность назначается лицо, имеющее:",
                        "dash_items": [
                            "высшее образование – специалитет;",
                            "опыт работы не менее 1 года.",
                        ],
                    },
                ],
            },
        ] + SAMPLE_SECTIONS[2:]
        path = create_di(
            position_full="специалиста тестового отдела",
            position_full_genitive="специалиста тестового отдела",
            sections=sections,
            output_path=str(tmp_path / "dashes.docx"),
        )
        assert os.path.isfile(path)

    def test_empty_org_config(self, tmp_path, monkeypatch):
        monkeypatch.setattr(di_generate, "_load_org_details", lambda: {})
        path = create_di(
            position_full="специалиста",
            position_full_genitive="специалиста",
            sections=SAMPLE_SECTIONS,
            output_path=str(tmp_path / "empty_cfg.docx"),
        )
        assert os.path.isfile(path)

    def test_custom_acquaintance_rows(self, tmp_path):
        path = create_di(
            position_full="специалиста",
            position_full_genitive="специалиста",
            sections=SAMPLE_SECTIONS,
            output_path=str(tmp_path / "custom_rows.docx"),
            acquaintance_rows=5,
        )
        assert os.path.isfile(path)
