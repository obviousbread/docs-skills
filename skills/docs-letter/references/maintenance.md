# Обслуживание скилла letter

## Синхронизация

`generate.py` и reference-файлы должны быть синхронизированы:

| Изменение | Где обновить |
|-----------|-------------|
| Параметры `create_letter()` | `helpers.md` (таблица параметров) |
| Реквизиты организации | `~/.docs-plugin/org_details.md` (загружается автоматически при вызове) |
| Новые типы абзацев в body_paragraphs | `helpers.md` (формат body_paragraphs) |
| Новый тип письма | `letter-patterns.md` + `usage-examples.md` |
| Новый хелпер / API | `helpers.md` |
| Workflow / стиль | `SKILL.md` |
| Локальный сторонний бланк | `~/.docs-plugin/org_notes.md` + `~/.docs-plugin/letter/scripts/*.py` (не коммитить) |

## Обновление реквизитов организации

При смене руководителя или контактных данных обновить `~/.docs-plugin/org_details.md` - поля `leader_title`, `leader_name_nom`, `author_name_short`, `author_phone`. Генератор загружает реквизиты автоматически при каждом вызове.

## Структура файлов

```
letter/
├── SKILL.md                    # Workflow + стиль + файловая политика
├── generate.py                 # Генератор писем организации
├── scripts/                    # Одноразовые скрипты (gitignored)
│   └── *.py
└── references/
    ├── letter-patterns.md      # Паттерны деловой переписки (10+ типов)
    ├── helpers.md              # API generate.py
    ├── usage-examples.md       # 9 примеров кода
    # Примеры из finetune: ~/.docs-plugin/examples/letter.md (вне репо)
    └── maintenance.md          # Этот файл
```
