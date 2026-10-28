# Синхронизация

`new_document()` отключает четыре параметра разбивки абзацев в docDefaults и стилях. Перечень параметров в `SKILL.md`.

При изменении `generate.py` одновременно проверь:

- сигнатуру в `references/helpers.md`;
- workflow и пример в `SKILL.md`;
- структуру и инварианты в `references/structure.md` и `quality-checklist.md`;
- поля конфигурации в `docs-init`.

После изменения запусти `test_generate.py`, валидатор скилла и обезличивающий поиск по всей папке.
