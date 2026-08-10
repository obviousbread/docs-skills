# docs-polozheniya implementation plan

Source: [docs-polozheniya brainstorm](../brainstorms/2026-08-10-docs-polozheniya.md)

## Success criteria

1. `docs-polozheniya` is independently discoverable and generates a standalone DOCX regulation for both supported unit types.
2. Runtime content contains no real personal, organizational, or workstation-identifying data.
3. Organization details, leader forms, approvers, and output directory come from config or explicit arguments.
4. Existing output files are never overwritten.
5. `docs-di`, `docs-ord`, `docs-init`, README, tests, and changelog describe the new boundary consistently.
6. The full test suite and `npx skills add . --list` pass.

## Steps

1. Add the portable skill, generator, references, and focused tests.
2. Extend `docs-init` configuration and update adjacent skill routing.
3. Update README, changelog, and the repository learning produced by this correction.
4. Run targeted tests, full tests, discovery, and the blocking personal-data gate.
5. Present the release diff for explicit confirmation before commits, merge, tag, push, and GitHub release.

## Verification

- `python3 -m pytest tests/test_polozheniya.py`
- `python3 -m pytest tests`
- `npx skills add . --list`
- repository and diff personal-data scans
