"""Shared data layer for docs skills.

Paths:
    ~/.docs-plugin/generations.jsonl              - Generation log (one JSON object per line)
    ~/.docs-plugin/{category}/examples/{cat}.md   - Persistent .md examples (from finetune)
    ~/.docs-plugin/{category}/scripts/            - One-off generation scripts
"""

import json
import os
import shutil
from datetime import datetime

DATA_DIR = os.path.expanduser("~/.docs-plugin")
LOG_PATH = os.path.join(DATA_DIR, "generations.jsonl")
EXAMPLES_DIR = os.path.join(DATA_DIR, "examples")  # deprecated, use examples_file()
ORG_DETAILS_PATH = os.path.join(DATA_DIR, "org_details.md")

_CATEGORIES = ("ord", "letter", "memo", "di")


def scripts_dir(category):
    """Return path to ~/.docs-plugin/{category}/scripts/, creating if needed."""
    d = os.path.join(DATA_DIR, category, "scripts")
    os.makedirs(d, exist_ok=True)
    return d


def examples_file(category):
    """Return path to ~/.docs-plugin/{category}/examples/{category}.md."""
    d = os.path.join(DATA_DIR, category, "examples")
    os.makedirs(d, exist_ok=True)
    return os.path.join(d, f"{category}.md")


def _migrate_examples():
    """Move examples from old flat layout to per-category dirs."""
    old_dir = os.path.join(DATA_DIR, "examples")
    if not os.path.isdir(old_dir):
        return
    for cat in _CATEGORIES:
        old_file = os.path.join(old_dir, f"{cat}.md")
        if os.path.exists(old_file):
            new_file = examples_file(cat)
            if not os.path.exists(new_file):
                shutil.move(old_file, new_file)
    # Remove old dir if empty
    try:
        os.rmdir(old_dir)
    except OSError:
        pass


def _ensure_dir():
    """Create data dir if it doesn't exist. Run migrations."""
    os.makedirs(DATA_DIR, exist_ok=True)
    _migrate_examples()


def log_generation(doc_type, title, output_path, params=None):
    """Log a document generation event. Fail-safe: errors do not propagate."""
    try:
        _ensure_dir()
        entry = {
            "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "doc_type": doc_type,
            "title": title,
            "output_path": output_path,
            "params": params or {},
        }
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(json.dumps(entry, ensure_ascii=False) + "\n")
    except Exception:
        pass
