#!/usr/bin/env python3
"""Генератор протоколов совещаний (.docx).

Конфигурация организации загружается из ~/.docs-plugin/org_details.md.
При отсутствии файла запустите docs-init для первичной настройки.

Использование:
    python3 generate.py  # создаёт пример протокола
"""

import os
import shutil
import warnings as _warnings
import re as _re
from datetime import date

from docx import Document
from docx.shared import Pt, Cm, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

import sys as _sys

_HERE = os.path.dirname(os.path.abspath(__file__))
for _lib_path in (
    os.path.join(_HERE, "..", "..", "lib"),
    os.path.expanduser("~/.docs-plugin/runtime/lib"),
):
    if os.path.isdir(_lib_path) and _lib_path not in _sys.path:
        _sys.path.insert(0, _lib_path)
from db import log_generation, ORG_DETAILS_PATH
from docx_meta import new_document


_SLASH_OK = _re.compile(r'п/п|ИНН/КПП|[0-9]/[0-9]|[A-Za-z]/[A-Za-z]')


def _warn_slash(text):
    """Предупреждение при обнаружении косой черты между кириллическими словами."""
    for m in _re.finditer(r'[а-яА-ЯёЁ]+\s*/\s*[а-яА-ЯёЁ]+', text):
        if not _SLASH_OK.search(m.group()):
            _warnings.warn(
                f"Косая черта в тексте документа: «{m.group()}». "
                f"Используйте запятую, «или», «и» или скобки.",
                stacklevel=3,
            )


def _load_org_details():
    """Загрузить реквизиты из ~/.docs-plugin/org_details.md."""
    config = {}
    if not os.path.exists(ORG_DETAILS_PATH):
        print("WARNING: ~/.docs-plugin/org_details.md not found. Run docs-init to configure.")
        return config
    with open(ORG_DETAILS_PATH, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if ": " in line and not line.startswith("#"):
                key, _, value = line.partition(": ")
                config[key.strip()] = value.strip()
    return config


def _load_staff():
    """Загрузить штатный список из staff_file (xlsx).

    Возвращает list[{"lastname", "initials", "position"}].
    При отсутствии staff_file возвращает [] (сверка станет no-op).
    """
    cfg = _load_org_details()
    path = os.path.expanduser(cfg.get("staff_file", "") or "")
    if not path or not os.path.exists(path):
        return []
    try:
        from openpyxl import load_workbook
    except ImportError:
        return []
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    header = [str(c or "").strip().lower() for c in rows[0]]
    def idx(name_options):
        for opt in name_options:
            if opt in header:
                return header.index(opt)
        return None
    i_fio = idx(["фио", "ф.и.о.", "сотрудник"])
    i_pos = idx(["должность"])
    if i_fio is None:
        return []
    out = []
    for row in rows[1:]:
        cell = row[i_fio] if i_fio < len(row) else None
        if not cell:
            continue
        s = str(cell).strip()
        # «Фамилия Имя Отчество» или «Фамилия И.О.»
        parts = s.split()
        if len(parts) >= 2:
            lastname = parts[0]
            if len(parts) == 2 and "." in parts[1]:
                initials = parts[1]
            else:
                initials = ".".join(p[0] for p in parts[1:3]) + "."
            position = ""
            if i_pos is not None and i_pos < len(row) and row[i_pos]:
                position = str(row[i_pos]).strip().lower()
            out.append({"lastname": lastname, "initials": initials, "position": position})
    return out


def create_protocol(*args, **kwargs):
    """Заглушка — реализация в Task B15."""
    raise NotImplementedError("create_protocol будет реализован в Task B15")
