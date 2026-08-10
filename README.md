<p align="center">
  <img src="assets/logo.png" alt="Логотип docs-skills" width="320">
</p>

<h1 align="center">docs</h1>

<p align="center">
  Набор скиллов для ИИ-агентов для создания официальных документов на русском языке в формате <code>.docx</code>.
</p>

<p align="center">
  <strong>Русский</strong> · <a href="./README.en.md">English</a>
</p>

## Навыки

- `docs-init` — настройка организации и пути к базе знаний.
- `docs-ord` — приказы, распоряжения и указания.
- `docs-letter` — официальные письма.
- `docs-memo` — служебные записки.
- `docs-di` — должностные инструкции.
- `docs-polozheniya` — положения о структурных подразделениях.
- `docs-protocol` — протоколы совещаний.
- `proofread` — вычитка DOCX с отслеживаемыми правками и выборочными комментариями.

## Установка

Для установки используется официальный [Vercel Skills CLI](https://github.com/vercel-labs/skills):

```bash
npx skills add obviousbread/docs-skills
```

В интерактивном меню можно выбрать навыки, область установки — текущий проект или глобальное окружение, целевых агентов, а также способ установки — копирование или символические ссылки.

По умолчанию навыки устанавливаются в текущий проект. Стандартный общий каталог — `.agents/skills`. Если навыки также нужны Claude Code, выберите его отдельно: тогда они появятся и в `.claude/skills`.

Примеры установки без интерактивного меню:

```bash
# Codex, текущий проект
npx skills add obviousbread/docs-skills --skill '*' --agent codex --yes --copy

# Codex, глобальная установка
npx skills add obviousbread/docs-skills --skill '*' --agent codex --global --yes --copy

# Codex и Claude Code, текущий проект
npx skills add obviousbread/docs-skills --skill '*' --agent codex --agent claude-code --yes --copy
```

## Первый запуск

Запустите `docs-init`. Команда создаст файл `~/.docs-plugin/org_details.md` с реквизитами организации и путями для готовых документов. При необходимости в нём также можно указать путь к списку сотрудников и `knowledge_base_path` — путь к хранилищу Obsidian или другому каталогу с рабочими материалами.

Если база знаний настроена, навыки читают инструкции из её корня и находят контекст, нужный для подготовки документа. Пользовательские примеры хранятся там же.

## Разработка

```bash
python3 -m pytest tests
npx skills add . --list
```

Требования: Python 3, `python-docx` и `openpyxl`.

## Лицензия

MIT
