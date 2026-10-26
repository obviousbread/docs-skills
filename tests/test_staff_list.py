"""Tests for docs-staff-list and its docs-protocol consumer contract."""

import importlib.util
import subprocess
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import Workbook, load_workbook

ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / "skills" / "docs-staff-list" / "scripts" / "normalize.py"
SPEC = importlib.util.spec_from_file_location("docs_staff_list_normalize", SCRIPT)
normalizer = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(normalizer)
PROTOCOL_SCRIPT = ROOT / "skills" / "docs-protocol" / "generate.py"
PROTOCOL_SPEC = importlib.util.spec_from_file_location("docs_protocol_staff_consumer", PROTOCOL_SCRIPT)
staff_consumer = importlib.util.module_from_spec(PROTOCOL_SPEC)
PROTOCOL_SPEC.loader.exec_module(staff_consumer)


def _source_workbook(path, name="Тестов Тест Тестович"):
    wb = Workbook()
    ws = wb.active
    ws.append(["Сотрудник", "Дата рождения", "Должность", "Подразделение", "Филиал"])
    ws.append([name, "01.02.1980", "специалист", "Отдел А", "Филиал А"])
    wb.save(path)


def test_clean_fio():
    cases = {
        "Тестов Тест Тестович (внеш. совм.)": "Тестов Тест Тестович",
        "Примеров Пример Примерович, Основное место работы, Работа": "Примеров Пример Примерович",
        "Условная Анна АлексеевнаРабота": "Условная Анна Алексеевна",
        "Учебная Наталья Викторовна уборщица": "Учебная Наталья Викторовна",
        "Тестов Тимур Вагиф оглы": "Тестов Тимур Вагиф оглы",
        "Примеров Пример Примерович (метка)": "Примеров Пример Примерович",
        "Чистая Фамилия Отчество": "Чистая Фамилия Отчество",
        "Условный Пример осн": "Условный Пример",
        "Условный Пример внешн.": "Условный Пример",
    }
    for raw, expected in cases.items():
        assert normalizer.clean_fio(raw) == expected


def test_format_date():
    assert normalizer.format_date("21.02.1967") == "21.02.1967"
    assert normalizer.format_date("25.04.82") == "25.04.1982"
    assert normalizer.format_date("30.07.99") == "30.07.1999"
    assert normalizer.format_date("23.10.01") == "23.10.2001"
    assert normalizer.format_date(pd.Timestamp("1972-11-14")) == "14.11.1972"
    assert normalizer.format_date("1972-11-14 00:00:00") == "14.11.1972"
    assert normalizer.format_date("1946", year_only=True) == "1946"
    assert normalizer.format_date(pd.Timestamp("1946-03-08"), year_only=True) == "1946"
    assert normalizer.format_date("") == ""


def test_detect_format():
    hierarchical = pd.DataFrame([
        ["Штатная расстановка"], ["Подразделение"], ["Позиция"], ["Сотрудник, Состояние"],
    ])
    assert normalizer.detect_format(hierarchical)[0] == "shtatnaya"

    table = pd.DataFrame([
        ["№", "Сотрудник", "Должность", "Дата рождения"],
        ["1", "Тестов Тест Тестович", "специалист", "01.01.1980"],
    ])
    assert normalizer.detect_format(table)[0] == "table"

    split = [["1", "Тестова", "Теста", "Тестовна", "23.05.75", "специалист"]] * 5
    assert normalizer.detect_format(pd.DataFrame(split))[0] == "split_fio"


def test_cli_output_is_consumed_by_protocol(tmp_path):
    input_dir = tmp_path / "input"
    older_dir = tmp_path / "staff" / "09.08.2026"
    output_dir = tmp_path / "staff" / "10.08.2026"
    input_dir.mkdir()
    older_dir.mkdir(parents=True)
    output_dir.mkdir(parents=True)
    _source_workbook(input_dir / "branch-a.xlsx")
    _source_workbook(
        older_dir / "Сводный список сотрудников 09.08.2026.xlsx",
        "Устаревший Пример Примерович",
    )
    output = output_dir / "Сводный список сотрудников 10.08.2026.xlsx"

    result = subprocess.run(
        [
            sys.executable, str(SCRIPT),
            "--input-dir", str(input_dir),
            "--title", "Сводный список сотрудников 10.08.2026",
            "--output", str(output),
        ],
        capture_output=True,
        text=True,
        check=False,
    )
    assert result.returncode == 0, result.stderr
    assert "Верификация: OK" in result.stdout

    wb = load_workbook(output, read_only=True, data_only=True)
    ws = wb.active
    assert [cell.value for cell in ws[1]] == [
        "№", "Сотрудник", "Дата рождения", "Должность", "Подразделение", "Филиал",
    ]
    assert ws["B2"].value == "Тестов Тест Тестович"
    wb.close()

    staff = staff_consumer._load_staff(str(tmp_path / "staff"))
    assert staff == [{"lastname": "Тестов", "initials": "Т.Т.", "position": "специалист"}]


def test_protocol_loader_accepts_split_name_columns(tmp_path):
    path = tmp_path / "staff.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.append(["Список сотрудников"])
    ws.append([])
    ws.append(["Фамилия", "Имя", "Отчество", "Должность"])
    ws.append(["Примеров", "Пример", "Примерович", "руководитель"])
    wb.save(path)

    assert staff_consumer._load_staff(str(path)) == [
        {"lastname": "Примеров", "initials": "П.П.", "position": "руководитель"},
    ]


def test_cli_requires_output_after_probe(tmp_path):
    _source_workbook(tmp_path / "source.xlsx")
    result = subprocess.run(
        [sys.executable, str(SCRIPT), "--input-dir", str(tmp_path)],
        capture_output=True,
        text=True,
        check=False,
    )
    assert result.returncode == 2
    assert "--output обязателен" in result.stderr


def _run_cli(input_dir, output, *files):
    result = subprocess.run(
        [sys.executable, str(SCRIPT), "--input-dir", str(input_dir),
         "--files", *files, "--output", str(output)],
        capture_output=True, text=True, check=False,
    )
    assert result.returncode == 0, result.stderr
    return result


def test_rich_merged_export_preserves_all_fields_and_flattens_header(tmp_path):
    source = tmp_path / "source.xlsx"
    wb = Workbook(); ws = wb.active
    # Wide merged date/name cells also occur in exports, not just grouped headers.
    ws.append(["Сотрудник", None, None, "Дата рождения"] + [None] * 15 + [
        "СНИЛС", None, None, None, "Должность", "Подразделение", "Филиал",
        "Вид занятости", "Удостоверение личности", None, None, None, None,
        None, None, None, "Адрес места проживания", "Страна гражданства", "Состояние",
    ])
    ws.append([None] * 27 + ["Вид", "Серия", "Номер", "Дата выдачи", "Кем выдано",
                             "Код подразделения", "Место рождения", "Физическое лицо.Мобильный телефон"])
    for region in ("A1:C2", "D1:S2", "T1:W2", "AB1:AI1"):
        ws.merge_cells(region)
    data = ["Условный Пример Примерович (внеш. совм.)", None, None, "02.03.1980"] + [None] * 15 + [
        "000-000-000 00", None, None, None, "специалист ", "\nОтдел А", "Филиал А",
        "Внешнее совместительство", "Документ-образец", "00 01", "000123",
        datetime(2020, 1, 2), "Условное ведомство", "000-000", "Условный населенный пункт",
        "[телефон]", "[адрес]", "[страна]", "Работа",
    ]
    ws.append(data)
    for region in ("A3:C3", "D3:S3", "T3:W3"):
        ws.merge_cells(region)
    wb.save(source)
    output = tmp_path / "result.xlsx"
    _run_cli(tmp_path, output, source.name)
    book = load_workbook(output); result = book.active
    headers = [c.value for c in result[1]]
    assert headers == ["№", "Сотрудник", "Дата рождения", "СНИЛС", "Должность",
                       "Подразделение", "Филиал", "Вид занятости", "Вид", "Серия",
                       "Номер", "Дата выдачи", "Кем выдано", "Код подразделения",
                       "Место рождения", "Физическое лицо.Мобильный телефон",
                       "Адрес места проживания", "Страна гражданства", "Состояние"]
    assert result.max_row == 2 and result.max_column == 19
    values = [c.value for c in result[2]]
    assert values == [1, "Условный Пример Примерович", "02.03.1980"] + [data[19]] + data[23:]
    assert result["J2"].value == "00 01" and result["K2"].value == "000123"
    assert result["L2"].number_format == "dd.mm.yyyy"
    assert not result.merged_cells.ranges
    assert result.freeze_panes == "B2" and result.auto_filter.ref == "A1:S2"
    assert result.sheet_view.showGridLines is True
    for row in result:
        for cell in row:
            assert cell.font.name == "Arial" and cell.font.sz == 14
            assert cell.font.color.rgb == "00000000"
            assert cell.font.bold == (cell.row == 1)
            assert cell.fill.patternType is None
            assert cell.alignment.wrap_text and cell.alignment.vertical == "center"
            assert all(side.style == "thin" for side in (
                cell.border.left, cell.border.right, cell.border.top, cell.border.bottom,
            ))
    book.close()


def test_union_keeps_extra_columns_values_and_duplicates(tmp_path):
    for filename, extra in (("a.xlsx", "СНИЛС"), ("b.xlsx", "Примечание")):
        wb = Workbook(); ws = wb.active
        ws.append(["Сотрудник", "Должность", "Подразделение", "Филиал", extra])
        ws.append(["Условный Пример Примерович", "специалист", "Отдел А", "Филиал А",
                   "0000123" if extra == "СНИЛС" else "=служебная пометка"])
        if extra == "Примечание":
            ws["E2"].data_type = "s"
        wb.save(tmp_path / filename)
    output = tmp_path / "result.xlsx"
    _run_cli(tmp_path, output, "a.xlsx", "b.xlsx")
    wb = load_workbook(output); ws = wb.active
    headers = [c.value for c in ws[1]]
    assert headers == ["№", "Сотрудник", "Должность", "Подразделение", "Филиал",
                       "СНИЛС", "Дата рождения", "Примечание"]
    assert ws.max_row == 3
    assert ws["F2"].value == "0000123" and ws["H2"].value is None
    assert ws["F3"].value is None and ws["H3"].value == "=служебная пометка"
    assert ws["H3"].data_type == "s"
    assert [ws.cell(r, 1).value for r in (2, 3)] == [1, 2]
    wb.close()


def test_empty_department_and_branch_are_not_invented(tmp_path):
    wb = Workbook(); ws = wb.active
    ws.append(["Сотрудник", "Должность", "Подразделение", "Филиал", "Код подразделения"])
    ws.append(["Условный Пример Примерович", "специалист", "Отдел А", "Филиал А", "000-000"])
    ws.append(["Тестовый Пример Примерович", "специалист", None, None, "000-001"])
    source = tmp_path / "branch-a.xlsx"; wb.save(source)
    output = tmp_path / "result.xlsx"
    cli = _run_cli(tmp_path, output, source.name)
    result = load_workbook(output).active
    assert result["D3"].value is None and result["E3"].value is None
    assert result["F3"].value == "000-001"
    assert "Строка 3: пустое подразделение" in cli.stdout
    assert "Строка 3: пустой филиал" in cli.stdout


def test_flat_table_without_department_keeps_personal_fields():
    df = pd.DataFrame([
        ["Сотрудник", "Должность", "Код подразделения", "Место рождения"],
        ["Условный Пример Примерович", "специалист", "000-001", "[место]"],
    ])
    rows, _ = normalizer.parse_table(df)
    assert len(rows) == 1
    assert rows[0]["Код подразделения"] == "000-001"
    assert rows[0]["Место рождения"] == "[место]"
    assert rows[0]["Подразделение"] == rows[0]["Дата рождения"] == ""


def test_duplicate_subheaders_and_unlabelled_data_are_preserved():
    df = pd.DataFrame([
        ["Сотрудник", "Должность", "Подразделение", "Документ А", "Документ Б", None],
        [None, None, None, "Номер", "Номер", None],
        ["Условный Пример Примерович", "специалист", "Отдел А", "001", "002", "[данные]"],
    ])
    rows, _ = normalizer.parse_table(df)
    assert rows[0]["Документ А: Номер"] == "001"
    assert rows[0]["Документ Б: Номер"] == "002"
    assert rows[0]["Колонка 6"] == "[данные]"


def test_split_names_keep_source_columns_and_additional_fields(tmp_path):
    wb = Workbook(); ws = wb.active
    ws.append(["СНИЛС", "Фамилия", "Имя", "Отчество", "Дата рождения", "Должность",
               "Подразделение", "Филиал"])
    ws.append(["0000123", "Условный", "Пример", "Примерович", "02.03.1980",
               "специалист", "Отдел А", "Филиал А"])
    source = tmp_path / "source.xlsx"; wb.save(source)
    output = tmp_path / "result.xlsx"
    _run_cli(tmp_path, output, source.name)
    result = load_workbook(output).active
    assert [c.value for c in result[1]] == ["№", "СНИЛС", "Фамилия", "Имя", "Отчество",
                                          "Дата рождения", "Должность", "Подразделение",
                                          "Филиал", "Сотрудник"]
    assert result["B2"].value == "0000123"
    assert result["J2"].value == "Условный Пример Примерович"
