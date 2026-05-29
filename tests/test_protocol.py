"""Тесты для skills/docs-protocol/generate.py — create_protocol() и edit_protocol()."""

import os
import warnings

from docx.enum.text import WD_ALIGN_PARAGRAPH

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


class TestOrgConfig:
    def test_output_dir_protocol_loaded(self):
        cfg = protocol_generate._build_org_config()
        assert cfg["output_dir_protocol"] == "/tmp"

    def test_header_lines_have_parent_and_full_name(self):
        cfg = protocol_generate._build_org_config()
        assert any("Тестовое агентство" in l for l in cfg["header_lines"])
        assert any("Центр тестирования" in l for l in cfg["header_lines"])


class TestFormatters:
    def test_format_ru_date(self):
        assert protocol_generate._format_ru_date("12.05.2026") == "12 мая 2026 г."

    def test_format_ru_date_january(self):
        assert protocol_generate._format_ru_date("01.01.2026") == "1 января 2026 г."

    def test_format_fio_initials_first(self):
        assert protocol_generate._format_fio_initials_first("Иванов И.И.") == "И.И. Иванов"

    def test_format_fio_already_initials_first(self):
        # Если уже «И.О. Фамилия» — вернуть как есть.
        assert protocol_generate._format_fio_initials_first("И.И. Иванов") == "И.И. Иванов"


class TestHeaderBlock:
    def test_header_has_4_centered_bold_lines(self):
        from docx import Document
        doc = Document()
        cfg = protocol_generate._build_org_config()
        protocol_generate._make_header_block(doc, cfg)
        bold_centered = [p for p in doc.paragraphs
                         if p.alignment == WD_ALIGN_PARAGRAPH.CENTER
                         and any(r.bold for r in p.runs)]
        assert len(bold_centered) >= 2  # parent_org + full_name + (short_name)


class TestTitleBlock:
    def test_two_centered_bold_lines(self):
        from docx import Document
        doc = Document()
        protocol_generate._make_title_block(doc, "оперативного совещания")
        ps = [p for p in doc.paragraphs
              if p.alignment == WD_ALIGN_PARAGRAPH.CENTER and p.runs]
        assert any(r.text == "ПРОТОКОЛ" and r.bold for p in ps for r in p.runs)
        assert any(r.text == "оперативного совещания" and r.bold for p in ps for r in p.runs)


class TestDateNumberTable:
    def test_default_number_placeholder(self):
        from docx import Document
        doc = Document()
        protocol_generate._make_date_number_table(doc, "12.05.2026", "")
        cells = doc.tables[0].rows[0].cells
        assert "12 мая 2026 г." in cells[0].text
        assert cells[1].text.strip() == "№ _________"

    def test_explicit_number(self):
        from docx import Document
        doc = Document()
        protocol_generate._make_date_number_table(doc, "12.05.2026", "42")
        assert doc.tables[0].rows[0].cells[1].text.strip() == "№ 42"


class TestVenueChairBlock:
    def test_labels_bold(self):
        from docx import Document
        doc = Document()
        chair = {"lastname": "Алмазов", "initials": "А.А.",
                 "position": "и.о. директора"}
        protocol_generate._make_venue_chair_block(doc, "Москва", chair)
        labels = []
        for p in doc.paragraphs:
            for r in p.runs:
                if r.bold and r.text in ("Место проведения: ", "Председатель: "):
                    labels.append(r.text)
        assert "Место проведения: " in labels
        assert "Председатель: " in labels


class TestAttendeesTable:
    def test_chair_filtered_with_warning(self):
        from docx import Document
        import warnings as _w
        doc = Document()
        chair = {"lastname": "Алмазов", "initials": "А.А.", "position": "и.о. директора"}
        attendees = [
            {"lastname": "Алмазов", "initials": "А.А.", "position": "и.о. директора"},
            {"lastname": "Бирюзов", "initials": "Б.Б.", "position": "начальник кадров"},
        ]
        with _w.catch_warnings(record=True) as warn_log:
            _w.simplefilter("always")
            protocol_generate._make_attendees_table(doc, attendees, chair)
        body = doc.tables[0].rows
        names = [r.cells[0].text for r in body]
        assert all("Алмазов" not in n for n in names)
        assert any("Бирюзов" in n for n in names)
        assert any("Председатель" in str(x.message) for x in warn_log)

    def test_three_columns(self):
        from docx import Document
        doc = Document()
        chair = {"lastname": "Алмазов", "initials": "А.А.", "position": "и.о."}
        attendees = [{"lastname": "Бирюзов", "initials": "Б.Б.", "position": "кадры"}]
        protocol_generate._make_attendees_table(doc, attendees, chair)
        assert len(doc.tables[0].rows[0].cells) == 3
