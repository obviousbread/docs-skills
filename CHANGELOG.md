# Changelog

## Unreleased

## 0.0.4

### docs-finetune

- Letters: when extracting an example whose signer is NOT an employee of the user's organization (resolved from `~/.docs-plugin/org_details.md`), omit the `**Подписант:**` line. Addressee, greeting, body and attachments are kept as-is.
- Letters: when the body mentions the source organization's full or short name, replace it with the `[организация]` placeholder so the example can be reused for the current organization without manual rewriting.

## 0.0.3

### docs-di

- Added portable `docs-di` skill for generating Russian job descriptions (должностные инструкции) for specialist, head of department, and head of directorate positions.
- Integrated with shared `lib/db.py`: reads org details from `~/.docs-plugin/org_details.md`, logs generations to `~/.docs-plugin/generations.jsonl`, resolves output dir from `output_dir_di`.
- Separated user data from the package: the согласующие pool now lives in `~/.docs-plugin/di/approvers.md` (user-zone, not shipped). Repo ships `references/approvers.md.example` as a placeholder template. Same for reference example .md files — agent reads them from `~/.docs-plugin/di/examples/` if present.
- Organization-agnostic templates: hardcoded organization name and curator surnames stripped from `SKILL.md`, `tier_rules.md`, `structure.md`, `knowledge_base.md`, `helpers.md`, `quality-checklist.md`. References to specific roles use approver keys (`udkp_head`, `hr_head`, `legal_head`, `curating_dgd`) resolved at generation time from the user-zone approvers file.

### installer

- Added `docs-di` to the SKILLS array in `bin/install.js`.

### docs-init

- Block 6 (output paths) now also asks for the job-description folder (`output_dir_di`).

### references

- `references/org_details.md.example`: added `output_dir_di` field under «Пути вывода».

## 0.0.2

### installer

- Scoped skill cleanup to installer-managed directories instead of deleting every `docs-*` folder by prefix.
- Added installer ownership metadata: per-skill `.docs-managed.json` markers and `~/.docs-plugin/install-state.json`.
- Preserved clean reinstall semantics for managed skills so structure changes replace stale nested files.
- Added coverage for foreign `docs-*` preservation, legacy canonical skill replacement, and state-driven rename migration.

### docs-ord

- Fixed sub-attachment grift: подприложения к Положению/Программе/Порядку теперь содержат полную цепочку - дательный падеж родительского документа, согласованное «утвержденн*» и строку «приказом [орг] от ___ № ___». Поддерживается 14 типов документов.
- Documented "Таблица мониторинга" attachment type: column alignment rules, measurable threshold requirement, optional interpretation sub-table.
- Documented multi-responsible cell encoding (`\n` within, `\n\n` between people) in Plan мероприятий.
- Added sub-attachment title deduplication rule to SKILL.md step 5.

## 0.0.1

- Rebuilt distribution around a single `npx @obviousbread/docs` installer.
- Renamed standalone skills to portable kebab-case names: `docs-init`, `docs-ord`, `docs-letter`, `docs-memo`, `docs-finetune`.
- Added copy-based installs for Claude Code, Codex, and Gemini CLI.
