"""Тесты для skills/docs-protocol/generate.py — create_protocol() и edit_protocol()."""

import os
import warnings

from conftest import protocol_generate

create_protocol = protocol_generate.create_protocol


class TestSmoke:
    """Скилл собирает .docx без падений на минимальном валидном входе."""

    def test_minimal(self, tmp_path):
        path = create_protocol(
            subtype="оперативного совещания",
            chair={"lastname": "Алмазов", "initials": "А.А.",
                   "position": "и.о. генерального директора"},
            attendees=[
                {"lastname": "Бирюзов", "initials": "Б.Б.",
                 "position": "начальник отдела кадров"},
            ],
            items=[
                {"text": "Подготовить отчёт по обращениям граждан.",
                 "responsible": ["Бирюзов Б.Б."],
                 "deadline": "01.06.2026",
                 "subitems": None},
            ],
            venue="г. Москва, конференц-зал",
            doc_date="12.05.2026",
            output_path=str(tmp_path / "protocol.docx"),
        )
        assert os.path.isfile(path)
        assert path.endswith(".docx")


class TestHelpers:
    """Базовые хелперы корректно создают элементы."""

    def test_warn_slash_warns(self):
        with warnings.catch_warnings(record=True) as w:
            warnings.simplefilter("always")
            protocol_generate._warn_slash("первое/второе")
            assert any("Косая черта" in str(x.message) for x in w)

    def test_warn_slash_skips_legitimate(self):
        with warnings.catch_warnings(record=True) as w:
            warnings.simplefilter("always")
            protocol_generate._warn_slash("ИНН/КПП и п/п")
            assert not any("Косая черта" in str(x.message) for x in w)
