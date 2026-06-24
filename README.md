<p align="center">
  <img src="assets/logo.png" alt="docs-skills logo" width="320">
</p>

<h1 align="center">docs</h1>

<p align="center">
  Portable agent skills for generating official Russian administrative documents in <code>.docx</code> format.
</p>

---

`docs` installs the same `docs-*` skills into native harness skill directories, so Claude Code, Codex, and Gemini CLI can use one maintained version instead of copied, drifting folders.

## Features

- Generates official `.docx` documents in a consistent format.
- Supports orders, directives, instructions, official letters, internal memos, and job descriptions.
- Uses local organization details from `~/.docs-plugin/org_details.md`.
- Checks staff data from a configured `.xlsx` file when needed.
- Learns recurring wording patterns from your existing `.docx`, `.md`, and `.txt` files.

## Skills

- `docs-init` - configure or update organization details.
- `docs-ord` - orders, directives, and instructions.
- `docs-letter` - official letters on organizational letterhead.
- `docs-memo` - internal memos.
- `docs-di` - job descriptions (должностные инструкции).
- `docs-finetune` - extract reusable examples from real documents.

## Install

Install for all supported harnesses:

```bash
npx @obviousbread/docs@latest --all
```

Install for one harness:

```bash
npx @obviousbread/docs@latest --claude
npx @obviousbread/docs@latest --codex
npx @obviousbread/docs@latest --gemini
```

Running the same command again updates the installed skills. The installer copies the current package into `~/.docs-plugin/runtime` and then copies `docs-*` skills into native harness locations:

- Claude Code: `~/.claude/skills/docs-*`
- Codex: `~/.codex/skills/docs-*`
- Gemini CLI: `~/.gemini/skills/docs-*`

For Gemini CLI, use `/skills list` to confirm discovery after install. If Gemini CLI is already running, use `/skills reload` or restart the session.

Uninstall managed files:

```bash
npx @obviousbread/docs@latest --all --uninstall
```

Uninstall does not remove user configuration, examples, logs, or generated documents from `~/.docs-plugin/`.

## First Run

Ask your agent to run `docs-init`. It creates or updates:

- organization details
- author and signer details
- default approvers
- output folders for generated documents
- optional staff list path

Persistent user data lives in `~/.docs-plugin/`:

- `org_details.md` - organization details and output paths
- `org_notes.md` - free-form organization context
- `{category}/examples/{category}.md` - examples added by `docs-finetune`
- `generations.jsonl` - generation log

## Development

Run tests:

```bash
npm test
```

For local installer testing without touching real harness directories:

```bash
DOCS_INSTALL_HOME="$(mktemp -d)" node bin/install.js --all
```

For active Gemini CLI skill development, you can also link a local skill folder manually:

```bash
gemini skills link ./skills/docs-letter
```

## Requirements

- Node.js 18+ for the installer.
- Python 3 with `python-docx` and `openpyxl` for document generation.

## License

MIT
