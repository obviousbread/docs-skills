const assert = require("node:assert/strict");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");
const test = require("node:test");

const installer = require("../bin/install.js");

const repoRoot = path.resolve(__dirname, "..");

function tempHome() {
  return fs.mkdtempSync(path.join(os.tmpdir(), "docs-installer-"));
}

test("parseArgs defaults to all runtimes", () => {
  assert.deepEqual(installer.parseArgs([]), {
    runtimes: ["claude", "codex", "gemini"],
    uninstall: false,
    help: false,
  });
});

test("SKILLS matches docs skill directories", () => {
  const skillDirs = fs
    .readdirSync(path.join(repoRoot, "skills"), { withFileTypes: true })
    .filter((entry) => entry.isDirectory() && entry.name.startsWith("docs-"))
    .map((entry) => entry.name)
    .sort();

  assert.deepEqual([...installer.SKILLS].sort(), skillDirs);
});

test("install copies managed skills without deleting unrelated skills", () => {
  const home = tempHome();
  const skillsRoot = path.join(home, ".claude", "skills");
  fs.mkdirSync(path.join(skillsRoot, "docs-custom"), { recursive: true });
  fs.mkdirSync(path.join(skillsRoot, "other-skill"), { recursive: true });

  installer.install({
    home,
    sourceRoot: repoRoot,
    runtimes: ["claude"],
    quiet: true,
  });

  assert.equal(fs.existsSync(path.join(skillsRoot, "docs-custom")), true);
  assert.equal(fs.existsSync(path.join(skillsRoot, "other-skill")), true);
  for (const skill of installer.SKILLS) {
    assert.equal(fs.existsSync(path.join(skillsRoot, skill, "SKILL.md")), true);
    const marker = JSON.parse(
      fs.readFileSync(path.join(skillsRoot, skill, ".docs-managed.json"), "utf8"),
    );
    assert.equal(installer.isManagedMarker(marker, skill), true);
  }
});

test("install removes canonical legacy skill names on first managed install", () => {
  const home = tempHome();
  const skillsRoot = path.join(home, ".claude", "skills");
  fs.mkdirSync(path.join(skillsRoot, "docs-letter"), { recursive: true });
  fs.writeFileSync(path.join(skillsRoot, "docs-letter", "stale.txt"), "old");
  fs.mkdirSync(path.join(skillsRoot, "docs-custom"), { recursive: true });

  installer.install({
    home,
    sourceRoot: repoRoot,
    runtimes: ["claude"],
    quiet: true,
  });

  assert.equal(fs.existsSync(path.join(skillsRoot, "docs-letter", "stale.txt")), false);
  assert.equal(fs.existsSync(path.join(skillsRoot, "docs-letter", "SKILL.md")), true);
  assert.equal(fs.existsSync(path.join(skillsRoot, "docs-custom")), true);
});

test("install writes runtime and preserves user config", () => {
  const home = tempHome();
  const dataRoot = path.join(home, ".docs-plugin");
  const orgDetails = path.join(dataRoot, "org_details.md");
  fs.mkdirSync(dataRoot, { recursive: true });
  fs.writeFileSync(orgDetails, "short_name: TEST\n");

  installer.install({
    home,
    sourceRoot: repoRoot,
    runtimes: ["codex"],
    quiet: true,
  });

  assert.equal(fs.existsSync(path.join(dataRoot, "runtime", "skills", "docs-ord", "SKILL.md")), true);
  assert.equal(fs.readFileSync(orgDetails, "utf8"), "short_name: TEST\n");
  const state = installer.readInstallState(installer.defaultPaths({ home, sourceRoot: repoRoot }));
  assert.equal(state.package, "@obviousbread/docs");
  assert.deepEqual(state.runtimes.codex.skills, installer.SKILLS);
});

test("install copies shared web-search.md reference into runtime", () => {
  const home = tempHome();
  const dataRoot = path.join(home, ".docs-plugin");

  installer.install({
    home,
    sourceRoot: repoRoot,
    runtimes: ["codex"],
    quiet: true,
  });

  assert.equal(
    fs.existsSync(path.join(dataRoot, "runtime", "references", "web-search.md")),
    true,
  );
});

test("install removes state-tracked legacy skill names during rename migration", () => {
  const home = tempHome();
  const paths = installer.defaultPaths({ home, sourceRoot: repoRoot });
  const skillsRoot = paths.claudeSkills;
  fs.mkdirSync(path.join(skillsRoot, "docs-old"), { recursive: true });
  fs.mkdirSync(path.dirname(paths.installState), { recursive: true });
  fs.writeFileSync(
    paths.installState,
    JSON.stringify(
      {
        package: "@obviousbread/docs",
        version: "0.0.0",
        updatedAt: "2026-01-01T00:00:00.000Z",
        runtimes: {
          claude: {
            skills: ["docs-old"],
          },
        },
      },
      null,
      2,
    ),
  );

  installer.install({
    home,
    sourceRoot: repoRoot,
    runtimes: ["claude"],
    quiet: true,
  });

  assert.equal(fs.existsSync(path.join(skillsRoot, "docs-old")), false);
  assert.equal(fs.existsSync(path.join(skillsRoot, "docs-letter", "SKILL.md")), true);
});

test("gemini install uses native skills and cleans up legacy docs context", () => {
  const home = tempHome();
  const geminiMd = path.join(home, ".gemini", "GEMINI.md");
  const skillsRoot = path.join(home, ".gemini", "skills");
  const unrelatedSkill = path.join(skillsRoot, "other-skill");
  const oldManagedSkill = path.join(skillsRoot, "docs-custom");
  const legacyDocs = path.join(home, ".gemini", "docs");
  fs.mkdirSync(path.dirname(geminiMd), { recursive: true });
  fs.mkdirSync(unrelatedSkill, { recursive: true });
  fs.mkdirSync(oldManagedSkill, { recursive: true });
  fs.mkdirSync(legacyDocs, { recursive: true });
  fs.writeFileSync(
    geminiMd,
    "# Existing Gemini context\n\n<!-- docs-skills:start -->\n@/tmp/old-docs/GEMINI.md\n<!-- docs-skills:end -->\n",
  );

  installer.install({
    home,
    sourceRoot: repoRoot,
    runtimes: ["gemini"],
    quiet: true,
  });

  const installed = fs.readFileSync(geminiMd, "utf8");
  assert.equal(installed.includes("docs-skills:start"), false);
  assert.match(installed, /Existing Gemini context/);
  assert.equal(fs.existsSync(path.join(skillsRoot, "other-skill")), true);
  assert.equal(fs.existsSync(path.join(skillsRoot, "docs-custom")), true);
  assert.equal(fs.existsSync(path.join(skillsRoot, "docs-letter", "SKILL.md")), true);
  assert.equal(fs.existsSync(legacyDocs), false);

  installer.uninstall(installer.defaultPaths({ home, sourceRoot: repoRoot }), ["gemini"]);

  const uninstalled = fs.readFileSync(geminiMd, "utf8");
  assert.equal(uninstalled.includes("docs-skills:start"), false);
  assert.match(uninstalled, /Existing Gemini context/);
  assert.equal(fs.existsSync(path.join(skillsRoot, "docs-letter")), false);
  assert.equal(fs.existsSync(path.join(skillsRoot, "other-skill")), true);
  assert.equal(fs.existsSync(path.join(skillsRoot, "docs-custom")), true);
});

test("gemini install leaves normal GEMINI.md content untouched", () => {
  const home = tempHome();
  const geminiMd = path.join(home, ".gemini", "GEMINI.md");
  const geminiContent = "# User Gemini context\n\nKeep this trailing whitespace.  \n";
  fs.mkdirSync(path.dirname(geminiMd), { recursive: true });
  fs.writeFileSync(geminiMd, geminiContent);

  installer.install({
    home,
    sourceRoot: repoRoot,
    runtimes: ["gemini"],
    quiet: true,
  });

  assert.equal(fs.readFileSync(geminiMd, "utf8"), geminiContent);
  assert.equal(fs.existsSync(path.join(home, ".gemini", "skills", "docs-init", "SKILL.md")), true);
});

test("uninstall removes managed skill directories without touching foreign docs skills", () => {
  const home = tempHome();
  const paths = installer.defaultPaths({ home, sourceRoot: repoRoot });
  const skillsRoot = paths.codexSkills;

  installer.install({
    home,
    sourceRoot: repoRoot,
    runtimes: ["codex"],
    quiet: true,
  });

  fs.mkdirSync(path.join(skillsRoot, "docs-custom"), { recursive: true });
  installer.uninstall(paths, ["codex"]);

  assert.equal(fs.existsSync(path.join(skillsRoot, "docs-letter")), false);
  assert.equal(fs.existsSync(path.join(skillsRoot, "docs-custom")), true);
  const state = installer.readInstallState(paths);
  assert.equal(state, null);
});
