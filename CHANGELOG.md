# Changelog

## Unreleased

## 0.3.1

### chore

- Added a pre-release content-review step (Releases step 0) and a content-hygiene rule to the Forbidden section of `AGENTS.md`.
- Removed all references to the federal medical system **профильная ИС** from skill references (`docs-di/references/knowledge_base.md`, `docs-di/references/quality-checklist.md`), replaced with the generic «профильные информационные системы».
- Cleaned up the `docs-di` smoke-test sample data and removed stray empty string literals in `generate.py`.
- Reworded the `parent_org_short` placeholder comment in `references/org_details.md.example` to a generic example.
- Fixed whitespace alignment in a `docs-protocol` signature example.

### packaging

- Updated `repository.url` in `package.json` to `github.com/obviousbread/docs-skills` (repository rename).

## 0.3.0

### docs-ord

- Added **review mode**: when an existing order/directive (`.doc`/`.docx`) is supplied with a review/fix request, the skill reviews the document instead of creating one. Mode is detected up front (creation vs review) and announced in `SKILL.md`.
- Two-loop review: a primary **holistic** loop (read the full rule corpus and review as if rebuilding the document from scratch) plus a deterministic **conformance** loop against the formatting invariants of `generate.py`. The checklist is an explicit floor, not a ceiling.
- New reference `references/review-checklist.md`: structural-layer invariants (page margins, requisites/title/signature/approval as tables, section vs page breaks, TNR numbering font, acquaintance-sheet columns, content-table borders/autofit/cantSplit) and text-layer rules (quotes, dashes, slashes, DD.MM.YYYY dates, semicolon-enumeration → list, NPA requisite format, alias consistency), with classification into text-layer vs structural-layer fixes.
- Remediation policy: text-layer fixes → Word track-changes in a `<name>_правки.docx` copy (original untouched); structural fixes (cannot be expressed as track-changes) → rebuild via `generate.py` into a conformant `.docx`. The skill offers three variants — text-only, structural-only, or a combined rebuild that applies everything at once (recommended).
- Track-changes revision author is read from the `revision_author` config key (generic default when absent), not hardcoded in the runtime.

### docs-init

- Block 3 now also asks for `revision_author` (track-changes author used in docs-ord review mode).

### config

- `references/org_details.md.example`: added `revision_author` key.

## 0.2.0

### docs-protocol

- Added portable `docs-protocol` skill for generating meeting minutes (.docx) in the format of the reference protocol from 2026-05-12. Public API: `create_protocol(subtype, chair, attendees, items, venue, doc_date, doc_number, secretary, notify_persons, output_path)`. Generates header, title block («ПРОТОКОЛ\n<subtype>»), date+number table, venue/chair block, attendees table, «РЕШИЛИ» with Word-native decimal numbering, signature block (two-line position), optional secretary block, and a mandatory acquaintance sheet (union of attendees and responsibles, deduplicated and sorted A-Я, chair excluded as signatory).
- Edit sub-workflow: `edit_protocol(src_path, edits)` copies the source with «2» suffix and inserts edits via Word track changes (`w:ins`/`w:del`). Original file is never touched.
- Staff verification: `_verify_fios` hard-fails on unknown FIO with top-3..5 fuzzy candidates via `difflib.get_close_matches` (no external dependency).
- Heuristic input classification (raw notes / structured / ready list) by presence of «Председатель:» / «Решили:» / «Присутствовали:» markers.

### installer

- Added `docs-protocol` to the SKILLS array in `bin/install.js`.

### docs-init

- Block 6 (output paths) now also asks for the protocol folder (`output_dir_protocol`).

### docs-finetune

- Indexes protocol examples from `~/.docs-plugin/protocol/examples/`.

### references

- `references/org_details.md.example`: added `output_dir_protocol` field under «Пути вывода».

## 0.1.0

### docs-letter

- **BREAKING**: dropped `executor_in_body` parameter from `create_letter()`. The executor block is now always placed in the body of the document after the signature, separated by 5 empty paragraphs (gray 8pt). The first-page footer mode and `_build_first_page_footer` are removed.
- Header is now a 2-column table (left = blank requisites, right = addressee). The previously empty middle spacer column is gone; the addressee cell loses its inner left indent and aligns flush-left.
- Lists (`numbered`, `dashed`) are now real Word lists (registered in `numbering.xml`, attached via `<w:numPr>`) instead of plain paragraphs with manual `"1. "` or `"– "` prefixes. Each consecutive `numbered` group claims a fresh `numId` so numbering restarts as expected; `dashed` items share one en-dash bullet list.
- Style: forbid the `«В ответ на Ваше письмо от ДД.ММ.ГГГГ № ХХХ сообщаем/направляем...»` opener — the incoming letter's metadata is already in the header `на №` field. Replaced with `«Сообщаем Вам...»`, `«По Вашему запросу...»`, `«Направляем Вам...»`, `«Во исполнение...»` in `letter-patterns.md` (§1, §2, §13) and the example in `usage-examples.md`.
- Style: codified lowercase/uppercase rules for act references. Lowercase for act types as common nouns (приказ, распоряжение, постановление, протокол, письмо, ...); uppercase for proper names of supreme NPAs (Конституция, Федеральный закон, Указ Президента, Кодекс as part of a name, ...). Full citation format documented with worked examples in `letter-patterns.md`.
- Style: introduced abbreviation `(далее — X)` only if X is used later in the body. Added a pre-flight check before the preview step and a checklist entry; bare "(далее — ФГБУ «Научный центр»)" with no downstream use is forbidden.

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
