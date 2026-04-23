# Platform Support

`docs` is distributed as portable skills. The installer copies the same `docs-*` skill folders into native harness locations.

## Claude Code

- Skills are installed to `~/.claude/skills/docs-*`.
- Read `~/.docs-plugin/org_details.md` explicitly before generating documents.
- If subagents are available, use them for verification tasks described in the skill. Otherwise continue inline.

## Codex

- Skills are installed to `~/.codex/skills/docs-*`.
- Read `~/.docs-plugin/org_details.md` explicitly before generating documents.
- Use native Codex file, shell, web, and agent tools when available.

## Gemini CLI

- Skills are installed to `~/.gemini/skills/docs-*`.
- Use `/skills list` to confirm discovery after install. If a Gemini CLI session is already running, use `/skills reload` or restart it.
- During local development, `gemini skills link ./skills/<skill-name>` can link one skill folder without copying it.
- Read `~/.docs-plugin/org_details.md` explicitly before generating documents.
- Gemini does not provide the same subagent model as Claude Code; perform verification inline.

## Shared Runtime

- User data: `~/.docs-plugin/`
- Installed package runtime: `~/.docs-plugin/runtime`
- Organization template: `~/.docs-plugin/runtime/references/org_details.md.example`
