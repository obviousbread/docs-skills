# docs

Portable agent skills for official `.docx` document generation.

## Critical Rules

### Never Commit

- `~/.docs-plugin/` - persistent user data: organization details, examples, scripts, logs, and installer runtime.
- Output `.docx` files.

### Skill Sync

When modifying `generate.py` in any document-generation skill, update its `SKILL.md` and `references/maintenance.md` if present. Divergence breaks agent behavior.

### Runtime Shape

- `package.json` is the only version source.
- `bin/install.js` is the only supported distribution path.
- No Claude/Codex/Gemini harness-specific manifests are maintained.
- Installer copies runtime files into `~/.docs-plugin/runtime` and copies `docs-*` skill folders into native harness locations.
- Generators import shared code from local `../../lib` first, then `~/.docs-plugin/runtime/lib` as an installed fallback.

## Python Environment

- Dependencies: `python-docx`, `openpyxl`.
- No project virtualenv is assumed.
- Shared data layer: `lib/db.py`.

## Commit Convention

Conventional Commits, English, lowercase imperative: `type(scope): summary`.

Common types: `feat`, `fix`, `docs`, `refactor`, `chore`.

## Releases

1. Update `package.json` version.
2. Update `CHANGELOG.md`.
3. Tag the release with the same version.
