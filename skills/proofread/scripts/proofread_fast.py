#!/usr/bin/env python3
"""Fast two-command workflow around the proofread redline engine."""

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
from pathlib import Path


def fail(message):
    raise SystemExit(message)


def run(command, env=None):
    result = subprocess.run(command, text=True, capture_output=True, env=env)
    if result.returncode:
        fail((result.stdout + result.stderr).strip())
    return (result.stdout + result.stderr).strip()


def engine():
    return Path(__file__).with_name("redline_docx.py")


def work_root(explicit=None):
    if explicit:
        root = Path(explicit).expanduser().resolve()
        if not root.is_dir():
            fail(f"Work root not found: {root}")
        return root
    current = Path.cwd().resolve()
    for candidate in (current, *current.parents):
        if (candidate / "ПРОЕКТЫ").is_dir():
            return candidate
    return None


def is_inside(path, directory):
    try:
        path.relative_to(directory)
        return True
    except ValueError:
        return False


def intake(source, root, task_name=None):
    source = source.expanduser().resolve()
    if not source.is_file() or source.suffix.lower() != ".docx":
        fail(f"DOCX not found: {source}")
    if root is None or is_inside(source, root):
        return source, None

    name = task_name or source.stem
    if Path(name).name != name or name in {"", ".", ".."}:
        fail("Task name must be one directory name.")
    destination_dir = root / "ПРОЕКТЫ" / name
    destination_dir.mkdir(parents=True, exist_ok=True)
    destination = destination_dir / source.name
    if destination.exists():
        fail(f"Refusing to overwrite: {destination}")
    shutil.move(source, destination)
    return destination, destination_dir


def prepare(args):
    started = time.perf_counter()
    source, created = intake(
        Path(args.input), work_root(args.work_root), args.task_name
    )
    extracted = run([sys.executable, str(engine()), "extract", str(source)])
    print(extracted)
    print(json.dumps({
        "source": str(source),
        "task_dir_created": str(created) if created else None,
        "elapsed_seconds": round(time.perf_counter() - started, 3),
    }, ensure_ascii=False))


def edit_stats(path):
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        fail(f"Cannot read edits.json: {exc}")
    edits = data.get("edits") if isinstance(data, dict) else data
    author = data.get("author") if isinstance(data, dict) else None
    if not isinstance(edits, list):
        fail('edits.json must contain an "edits" list.')
    if author is not None and (not isinstance(author, str) or not author.strip()):
        fail('edits.json "author" must be a non-empty string or omitted.')
    large = []
    for index, edit in enumerate(edits):
        if not isinstance(edit, dict):
            continue
        values = [edit.get("old"), edit.get("new", "")]
        if any(isinstance(value, str) and len(value) > 120 for value in values):
            large.append(index + 1)
    return len(edits), large, author or "Корректор"


def render_pair(source, output, final_output):
    soffice = shutil.which("soffice")
    pdftoppm = shutil.which("pdftoppm")
    if not soffice or not pdftoppm:
        fail("Required commands not found: soffice and pdftoppm.")

    qa = final_output.parent / ".proofread-qa" / final_output.stem
    if qa.exists():
        shutil.rmtree(qa)
    qa.mkdir(parents=True)

    with tempfile.TemporaryDirectory(prefix="proofread-fast-") as temporary:
        temporary = Path(temporary)
        profile = temporary / "profile"
        staged = temporary / "documents"
        profile.mkdir()
        staged.mkdir()
        shutil.copy2(source, staged / "source.docx")
        shutil.copy2(output, staged / "output.docx")
        env = os.environ.copy()
        env["TMPDIR"] = str(temporary)
        run([
            soffice,
            "--headless",
            f"-env:UserInstallation={profile.as_uri()}",
            "--convert-to", "pdf",
            "--outdir", str(qa),
            str(staged / "source.docx"),
            str(staged / "output.docx"),
        ], env=env)

    counts = {}
    for label in ("source", "output"):
        pdf = qa / f"{label}.pdf"
        if not pdf.is_file():
            fail(f"Expected PDF not created: {pdf}")
        run([pdftoppm, "-png", "-r", "150", str(pdf), str(qa / label)])
        pages = sorted(qa.glob(f"{label}-*.png"))
        if not pages or any(page.stat().st_size < 1000 for page in pages):
            fail(f"Invalid PNG render for {label}.")
        counts[label] = len(pages)
    return qa, counts


def finish(args):
    started = time.perf_counter()
    source = Path(args.input).expanduser().resolve()
    edits = Path(args.edits).expanduser().resolve()
    if not source.is_file() or source.suffix.lower() != ".docx":
        fail(f"DOCX not found: {source}")
    if not edits.is_file():
        fail(f"edits.json not found: {edits}")
    output = (
        Path(args.output).expanduser().resolve()
        if args.output
        else source.with_name(f"{source.stem} (с правками).docx")
    )
    if output == source:
        fail("Output must not overwrite the source DOCX.")
    edit_count, large_edits, author = edit_stats(edits)
    temporary_output = output.with_name(f".{output.stem}.proofread-fast.tmp.docx")
    temporary_output.unlink(missing_ok=True)

    try:
        phase = time.perf_counter()
        review = run([
            sys.executable, str(engine()), "review", str(source),
            str(temporary_output), "--edits", str(edits),
        ])
        review_seconds = time.perf_counter() - phase

        phase = time.perf_counter()
        audit = run([
            sys.executable, str(engine()), "audit", str(temporary_output),
            "--author", author,
        ])
        audit_seconds = time.perf_counter() - phase
        match = re.search(
            r"zip_ok=True.+track_changes=True.+revision_runs=(\d+).+yellow_runs=(\d+)",
            audit,
        )
        if (
            not match
            or int(match.group(1)) != int(match.group(2))
            or (edit_count and not int(match.group(1)))
        ):
            fail("Unexpected audit result: " + audit)
        comments_match = re.search(r"comments=(\d+)", audit)
        if not comments_match:
            fail("Unexpected audit result: " + audit)

        phase = time.perf_counter()
        qa, pages = render_pair(source, temporary_output, output)
        render_seconds = time.perf_counter() - phase
        temporary_output.replace(output)
        review = review.replace(str(temporary_output), str(output))
    except BaseException:
        temporary_output.unlink(missing_ok=True)
        raise

    print(json.dumps({
        "output": str(output),
        "qa_output": str(qa),
        "edits": edit_count,
        "large_edits_over_120_chars": large_edits,
        "comments": int(comments_match.group(1)),
        "pages": pages,
        "page_count_changed": pages["source"] != pages["output"],
        "seconds": {
            "review": round(review_seconds, 3),
            "audit": round(audit_seconds, 3),
            "render": round(render_seconds, 3),
            "total": round(time.perf_counter() - started, 3),
        },
        "review_result": review,
        "audit_result": audit,
    }, ensure_ascii=False, indent=2))


def main():
    parser = argparse.ArgumentParser()
    commands = parser.add_subparsers(dest="command", required=True)

    prepare_parser = commands.add_parser("prepare")
    prepare_parser.add_argument("input")
    prepare_parser.add_argument("--task-name")
    prepare_parser.add_argument("--work-root")
    prepare_parser.set_defaults(func=prepare)

    finish_parser = commands.add_parser("finish")
    finish_parser.add_argument("input")
    finish_parser.add_argument("--edits", required=True)
    finish_parser.add_argument("--output")
    finish_parser.set_defaults(func=finish)

    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
