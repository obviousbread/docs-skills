#!/usr/bin/env node

const fs = require("fs");
const os = require("os");
const path = require("path");

const pkg = require("../package.json");

const SKILLS = [
  "docs-init",
  "docs-ord",
  "docs-letter",
  "docs-memo",
  "docs-finetune",
];

const RUNTIMES = ["claude", "codex", "gemini"];
const GEMINI_START = "<!-- docs-skills:start -->";
const GEMINI_END = "<!-- docs-skills:end -->";
const EXCLUDED_NAMES = new Set([
  ".git",
  ".DS_Store",
  ".pytest_cache",
  "__pycache__",
  "node_modules",
  "tests",
]);

function usage() {
  return `docs ${pkg.version}

Install portable docs skills into native agent harness directories.

Usage:
  npx @obviousbread/docs@latest
  npx @obviousbread/docs@latest --all
  npx @obviousbread/docs@latest --claude
  npx @obviousbread/docs@latest --codex
  npx @obviousbread/docs@latest --gemini
  npx @obviousbread/docs@latest --all --uninstall

Options:
  --all        Install for Claude Code, Codex, and Gemini CLI (default)
  --claude    Install only Claude Code skills
  --codex     Install only Codex skills
  --gemini    Install only Gemini CLI skills
  --uninstall Remove managed docs skills
  --help      Show this help
`;
}

function parseArgs(argv) {
  const selected = new Set();
  let uninstall = false;
  let help = false;

  for (const arg of argv) {
    if (arg === "--help" || arg === "-h") {
      help = true;
    } else if (arg === "--uninstall") {
      uninstall = true;
    } else if (arg === "--all") {
      RUNTIMES.forEach((runtime) => selected.add(runtime));
    } else if (arg === "--claude") {
      selected.add("claude");
    } else if (arg === "--codex") {
      selected.add("codex");
    } else if (arg === "--gemini") {
      selected.add("gemini");
    } else {
      throw new Error(`unknown option: ${arg}`);
    }
  }

  if (!help && selected.size === 0) {
    RUNTIMES.forEach((runtime) => selected.add(runtime));
  }

  return { runtimes: [...selected], uninstall, help };
}

function defaultPaths(options = {}) {
  const home = options.home || process.env.DOCS_INSTALL_HOME || os.homedir();
  const dataRoot =
    options.dataRoot || process.env.DOCS_DATA_DIR || path.join(home, ".docs-plugin");
  return {
    home,
    dataRoot,
    runtimeRoot: path.join(dataRoot, "runtime"),
    sourceRoot: options.sourceRoot || path.resolve(__dirname, ".."),
    claudeSkills: path.join(home, ".claude", "skills"),
    codexSkills: path.join(home, ".codex", "skills"),
    geminiSkills: path.join(home, ".gemini", "skills"),
    legacyGeminiDocs: path.join(home, ".gemini", "docs"),
    geminiContext: path.join(home, ".gemini", "GEMINI.md"),
  };
}

function shouldExclude(name) {
  return EXCLUDED_NAMES.has(name) || name.endsWith(".pyc") || name.endsWith(".pyo");
}

function copyDir(src, dest) {
  const stat = fs.statSync(src);
  if (!stat.isDirectory()) {
    throw new Error(`not a directory: ${src}`);
  }

  fs.rmSync(dest, { recursive: true, force: true });
  fs.mkdirSync(dest, { recursive: true });

  for (const entry of fs.readdirSync(src, { withFileTypes: true })) {
    if (shouldExclude(entry.name)) {
      continue;
    }

    const srcPath = path.join(src, entry.name);
    const destPath = path.join(dest, entry.name);

    if (entry.isDirectory()) {
      copyDir(srcPath, destPath);
    } else if (entry.isSymbolicLink()) {
      const target = fs.readlinkSync(srcPath);
      fs.symlinkSync(target, destPath);
    } else if (entry.isFile()) {
      fs.copyFileSync(srcPath, destPath);
    }
  }
}

function removeManagedSkills(skillsRoot) {
  if (!fs.existsSync(skillsRoot)) {
    return [];
  }

  const removed = [];
  for (const entry of fs.readdirSync(skillsRoot, { withFileTypes: true })) {
    if (!entry.name.startsWith("docs-")) {
      continue;
    }
    const target = path.join(skillsRoot, entry.name);
    fs.rmSync(target, { recursive: true, force: true });
    removed.push(target);
  }
  return removed;
}

function installRuntime(paths) {
  fs.mkdirSync(paths.dataRoot, { recursive: true });
  copyDir(paths.sourceRoot, paths.runtimeRoot);
  fs.writeFileSync(path.join(paths.runtimeRoot, "VERSION"), `${pkg.version}\n`);
  fs.writeFileSync(path.join(paths.dataRoot, "VERSION"), `${pkg.version}\n`);
}

function installSkills(runtimeSkillsRoot, skillsRoot) {
  fs.mkdirSync(skillsRoot, { recursive: true });
  removeManagedSkills(skillsRoot);

  for (const skill of SKILLS) {
    const src = path.join(runtimeSkillsRoot, skill);
    const dest = path.join(skillsRoot, skill);
    copyDir(src, dest);
  }
}

function removeGeminiBlock(content) {
  const pattern = new RegExp(
    `\\n?${escapeRegExp(GEMINI_START)}[\\s\\S]*?${escapeRegExp(GEMINI_END)}\\n?`,
    "g",
  );
  return content.replace(pattern, "\n").replace(/\n{3,}/g, "\n\n").trimEnd();
}

function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function removeLegacyGeminiBlock(paths) {
  if (!fs.existsSync(paths.geminiContext)) {
    return;
  }
  const existing = fs.readFileSync(paths.geminiContext, "utf8");
  if (!existing.includes(GEMINI_START) || !existing.includes(GEMINI_END)) {
    return;
  }
  const next = removeGeminiBlock(existing);
  if (next !== existing) {
    fs.writeFileSync(paths.geminiContext, next ? `${next}\n` : "");
  }
}

function installGemini(paths) {
  installSkills(path.join(paths.runtimeRoot, "skills"), paths.geminiSkills);
  fs.rmSync(paths.legacyGeminiDocs, { recursive: true, force: true });
  removeLegacyGeminiBlock(paths);
}

function uninstall(paths, runtimes) {
  if (runtimes.includes("claude")) {
    removeManagedSkills(paths.claudeSkills);
  }
  if (runtimes.includes("codex")) {
    removeManagedSkills(paths.codexSkills);
  }
  if (runtimes.includes("gemini")) {
    removeManagedSkills(paths.geminiSkills);
    fs.rmSync(paths.legacyGeminiDocs, { recursive: true, force: true });
    removeLegacyGeminiBlock(paths);
  }

  if (runtimes.length === RUNTIMES.length) {
    fs.rmSync(paths.runtimeRoot, { recursive: true, force: true });
    fs.rmSync(path.join(paths.dataRoot, "VERSION"), { force: true });
  }
}

function install(options = {}) {
  const paths = defaultPaths(options);
  const runtimes = options.runtimes || RUNTIMES;
  const log = options.quiet ? () => {} : console.log;

  installRuntime(paths);

  if (runtimes.includes("claude")) {
    installSkills(path.join(paths.runtimeRoot, "skills"), paths.claudeSkills);
    log(`installed Claude skills: ${paths.claudeSkills}`);
  }
  if (runtimes.includes("codex")) {
    installSkills(path.join(paths.runtimeRoot, "skills"), paths.codexSkills);
    log(`installed Codex skills: ${paths.codexSkills}`);
  }
  if (runtimes.includes("gemini")) {
    installGemini(paths);
    log(`installed Gemini skills: ${paths.geminiSkills}`);
  }

  log(`docs runtime: ${paths.runtimeRoot}`);
  log(`version: ${pkg.version}`);
}

function run(argv = process.argv.slice(2), options = {}) {
  const parsed = parseArgs(argv);
  if (parsed.help) {
    console.log(usage());
    return;
  }

  const paths = defaultPaths(options);
  if (parsed.uninstall) {
    uninstall(paths, parsed.runtimes);
    console.log("removed managed docs skills");
    return;
  }

  install({ ...options, runtimes: parsed.runtimes });
}

if (require.main === module) {
  try {
    run();
  } catch (error) {
    console.error(`error: ${error.message}`);
    process.exitCode = 1;
  }
}

module.exports = {
  SKILLS,
  RUNTIMES,
  defaultPaths,
  install,
  parseArgs,
  removeGeminiBlock,
  removeLegacyGeminiBlock,
  removeManagedSkills,
  run,
  uninstall,
};
