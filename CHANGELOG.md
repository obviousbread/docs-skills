# Changelog

## Unreleased

### docs-ord

- Fixed sub-attachment grift: подприложения к Положению/Программе/Порядку теперь содержат полную цепочку - дательный падеж родительского документа, согласованное «утвержденн*» и строку «приказом [орг] от ___ № ___». Поддерживается 14 типов документов.
- Documented "Таблица мониторинга" attachment type: column alignment rules, measurable threshold requirement, optional interpretation sub-table.
- Documented multi-responsible cell encoding (`\n` within, `\n\n` between people) in Plan мероприятий.
- Added sub-attachment title deduplication rule to SKILL.md step 5.

## 0.0.1

- Rebuilt distribution around a single `npx @obviousbread/docs` installer.
- Renamed standalone skills to portable kebab-case names: `docs-init`, `docs-ord`, `docs-letter`, `docs-memo`, `docs-finetune`.
- Added copy-based installs for Claude Code, Codex, and Gemini CLI.