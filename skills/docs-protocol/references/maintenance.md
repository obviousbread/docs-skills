# Правила поддержки скилла docs-protocol

## Синхронизация файлов

Любое изменение в `generate.py` требует проверки и при необходимости обновления смежных файлов:

| Что изменилось | Где обновить |
|---|---|
| Новая или изменённая функция-хелпер | `references/helpers.md` |
| Изменился публичный API `create_protocol()` | `references/helpers.md` + SKILL.md |
| Изменился workflow (порядок блоков, новый блок) | SKILL.md шаг «Описать структуру пользователю» |
| Изменилась схема person | `references/attendees-rules.md` |
| Новое ограничение или запрет | SKILL.md раздел форматирования |

Правило работает в обе стороны: если в SKILL.md появляется новое форматное требование, оно должно быть реализовано в `generate.py`.

## Word-нумерация

`numId` и `abstractNumId` назначаются динамически через `_next_num_id(doc)`. Нельзя хардкодить конкретные числа.

- `_setup_protocol_numbering(doc)` – decimal-нумерация пунктов «РЕШИЛИ» (`%1.`).
- `_setup_dash_numbering(doc)` – en-dash bullet для подпунктов.

При добавлении нового вида нумерации: создать аналогичную функцию `_setup_*_numbering(doc)`, вызывать `_next_num_id` в её начале, не использовать фиксированные идентификаторы.

## Схема person

Все участники (chair, attendees, secretary) передаются как словарь с тремя полями:

```python
{"lastname": "Фамилия", "initials": "И.О.", "position": "должность"}
```

При добавлении новых полей в схему – обновить `references/attendees-rules.md`.

`_verify_fios` сверяет `(lastname.lower(), initials.lower())` со штатным списком и при несовпадении выбрасывает `ValueError` с fuzzy-подсказками (до 5 кандидатов, порог 0.6).

## Импорт shared-кода

Сначала локальный `../../lib`, затем `~/.docs-plugin/runtime/lib` как fallback. Такой порядок обязателен для корректной работы как из dev-дерева, так и в установленном runtime. Не менять порядок без веской причины.

```python
for _lib_path in (
    os.path.join(_HERE, "..", "..", "lib"),
    os.path.expanduser("~/.docs-plugin/runtime/lib"),
):
    if os.path.isdir(_lib_path) and _lib_path not in _sys.path:
        _sys.path.insert(0, _lib_path)
```

## Тесты

- Файл тестов: `tests/test_protocol.py`.
- Общая конфигурация: `tests/conftest.py`.
- Фикстура `mock_staff` (autouse) подменяет `_load_staff` через `monkeypatch`, чтобы тесты не зависели от реального xlsx-файла.
- Все ФИО в тестах и примерах выдуманы. Никогда не использовать имена реальных людей.

При добавлении новых параметров в `create_protocol()` – добавить покрывающий тест в `test_protocol.py`.

## Запреты

- Не хардкодить `numId` и `abstractNumId`.
- Не использовать реальные ФИО в тестах или примерах.
- Не применять косые черты между кириллическими словами в тексте документа (функция `_warn_slash` будет предупреждать при нарушении).
- Не коммитить выходные `.docx`-файлы.
