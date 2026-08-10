<p align="center">
  <img src="assets/logo.png" alt="docs-skills logo" width="320">
</p>

<h1 align="center">docs</h1>

<p align="center">
  Agent skills for generating official Russian administrative documents in <code>.docx</code> format.
</p>

<p align="center">
  <a href="./README.md">Русский</a> · <strong>English</strong>
</p>

## Skills

- `docs-init` — organization settings and knowledge-base path.
- `docs-ord` — orders, directives, and instructions.
- `docs-letter` — official letters.
- `docs-memo` — internal memos.
- `docs-di` — job descriptions.
- `docs-polozheniya` — standalone organizational-unit regulations.
- `docs-protocol` — meeting minutes.
- `proofread` — DOCX proofreading with tracked changes and selective comments.

## Install

Distribution uses the official [Vercel Skills CLI](https://github.com/vercel-labs/skills):

```bash
npx skills add obviousbread/docs-skills
```

The interactive menu lets you choose skills, project or global scope, target agents, and copy or symlink installation. Project scope is the default; the canonical shared location is `.agents/skills`. Select Claude Code explicitly when you also want `.claude/skills`.

Non-interactive examples:

```bash
# Codex, current project
npx skills add obviousbread/docs-skills --skill '*' --agent codex --yes --copy

# Codex, global
npx skills add obviousbread/docs-skills --skill '*' --agent codex --global --yes --copy

# Codex and Claude Code, current project
npx skills add obviousbread/docs-skills --skill '*' --agent codex --agent claude-code --yes --copy
```

## First run

Run `docs-init`. It writes `~/.docs-plugin/org_details.md` with organization details, output paths, an optional staff-list path, and an optional `knowledge_base_path` pointing to an Obsidian vault or equivalent knowledge base.

When the knowledge-base path is configured, document skills read its root agent instructions and retrieve as much relevant source context as needed. User-specific examples are kept there as well.

## Development

```bash
python3 -m pytest tests
npx skills add . --list
```

Requirements: Python 3, `python-docx`, and `openpyxl`.

## License

MIT
