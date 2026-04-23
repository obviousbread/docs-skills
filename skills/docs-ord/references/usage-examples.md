# Примеры использования generate.py

## Основной вызов через `create_ord()`

```python
import os, sys
runtime_dir = os.path.expanduser("~/.docs-plugin/runtime")
sys.path.insert(0, os.path.join(runtime_dir, "skills", "docs-ord"))
from generate import create_ord

create_ord(
    doc_type="prikaz",             # "prikaz" | "rasporyazhenie" | "ukazanie"
    title="О назначении ответственных за ...",
    basis="На основании приказа ... от ДД.ММ.ГГГГ № ...",
    # points: str → уровень 0; (str, level) → явный уровень (0, 1, ...)
    # Нумерация автоматическая через встроенные списки Word (_list_para).
    # НЕ вписывать номера вручную – Word проставит их сам.
    points=[
        "Назначить ответственным ...",              # → 1.
        ("[Фамилия И.О.]:", 0),                     # → 2.
        ("обеспечить ...", 1),                       # → 2.1.
        ("организовать ...", 1),                     # → 2.2.
        "Возложить персональную ответственность ...",# → 3.
    ],
    control_person="заместителя генерального директора по стратегическому развитию [Фамилия И.О.]",
    notify_persons=[
        {"name": "[Фамилия И.О.]", "position": "Заведующий отделением"},
    ],
    approved_by=[
        {"name": "[Фамилия И.О.]", "position": "Заместитель генерального директора по стратегическому развитию"},
    ],
    has_attachment=False,
    doc_date="18.02.2026",
    output_path="~/Documents/1_ОРД/новый_приказ.docx",
)
```

## Приложение с заполненной таблицей (form_table)

```python
from docx.shared import Cm

create_ord(
    doc_type="prikaz",
    title="Об утверждении плана мероприятий по ...",
    basis="В целях ...",
    points=["Утвердить план мероприятий ... (приложение)."],
    control_person="заместителя генерального директора ... [Фамилия И.О.]",
    notify_persons=[
        {"name": "[Фамилия И.О.]", "position": "Начальник отдела"},
    ],
    attachments=[{
        "title": "План мероприятий по ... на 2026 год",
        "type": "form_table",
        "columns": ["№ п/п", "Мероприятие", "Сроки", "Ответственные", "Примечание"],
        "col_widths": [Cm(1.2), Cm(7.5), Cm(3.0), Cm(3.3), Cm(1.5)],
        "rows": [
            # Ячейка с \n разбивается на строки внутри ячейки
            ["1", "Провести анализ ...", "до 25.03.2026", "[Фамилия И.О.]\nНачальник отдела", ""],
            ["2", "Подготовить отчет ...", "Ежеквартально", "[Фамилия И.О.]\nНачальник отдела", ""],
        ],
    }],
    output_path="~/Documents/1_ОРД/план.docx",
)
```

Если `rows` - число (int), создается пустая таблица-шаблон:

```python
attachments=[{
    "title": "Форма отчета",
    "type": "form_table",
    "columns": ["№", "Показатель", "Значение"],
    "rows": 10,  # 10 пустых строк
}]
```

## Нумерация в специфичных скриптах (не через `create_ord`)

```python
from generate import _setup_ord_numbering, _add_num_instance, _list_para

num_id = _setup_ord_numbering(doc)           # инициализация (один раз на документ)

_list_para(doc, "Пункт первый", level=0, num_id=num_id)      # → 1.
_list_para(doc, "Подпункт", level=1, num_id=num_id)          # → 1.1.

app_num_id = _add_num_instance(doc)          # независимый счётчик для приложения
_list_para(doc, "Раздел", level=0, num_id=app_num_id, bold=True)  # → 1.
_list_para(doc, "Пункт", level=1, num_id=app_num_id)              # → 1.1.
```

## Запуск из Bash

```bash
python3 skills/docs-ord/generate.py
```
