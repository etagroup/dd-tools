#!/usr/bin/env python3
"""Bundle the latest source code into a zip file.

This script tries to:
1) Detect a Git repository (preferred) and bundle tracked + untracked (non-ignored) files.
2) Otherwise, bundle a directory tree while excluding common junk (.git, node_modules, etc.).

Usage:
  python bundle_latest_source.py --path /path/to/project --out source_code.zip

If --path is omitted, it uses the current working directory.
"""

from __future__ import annotations

import argparse
import os
import subprocess
import sys
import zipfile
from pathlib import Path

EXCLUDE_DIRS = {
    ".git",
    ".hg",
    ".svn",
    "node_modules",
    "__pycache__",
    ".pytest_cache",
    ".mypy_cache",
    ".ruff_cache",
    ".venv",
    "venv",
    ".idea",
    ".vscode",
    "dist",
    "build",
    ".DS_Store",
}

EXCLUDE_FILE_SUFFIXES = {
    ".pyc",
    ".pyo",
    ".pyd",
    ".so",
    ".dylib",
    ".dll",
    ".exe",
    ".class",
    ".o",
    ".a",
    ".zip",
}


def run(cmd: list[str], cwd: Path) -> str:
    out = subprocess.check_output(cmd, cwd=str(cwd), stderr=subprocess.STDOUT)
    return out.decode("utf-8", errors="replace").strip()


def is_git_repo(path: Path) -> bool:
    try:
        run(["git", "rev-parse", "--is-inside-work-tree"], cwd=path)
        return True
    except Exception:
        return False


def git_root(path: Path) -> Path:
    return Path(run(["git", "rev-parse", "--show-toplevel"], cwd=path))


def git_file_list(repo_root: Path) -> list[Path]:
    tracked = run(["git", "ls-files", "-z"], cwd=repo_root).split("\x00")
    untracked = run([
        "git",
        "ls-files",
        "--others",
        "--exclude-standard",
        "-z",
    ], cwd=repo_root).split("\x00")

    rels = [p for p in tracked + untracked if p]
    files: list[Path] = []
    for rel in rels:
        fp = repo_root / rel
        if fp.is_file():
            files.append(fp)
    return files


def should_exclude(path: Path, base: Path) -> bool:
    # Exclude by directory name anywhere in the relative path.
    rel_parts = path.relative_to(base).parts
    if any(part in EXCLUDE_DIRS for part in rel_parts):
        return True
    if path.suffix.lower() in EXCLUDE_FILE_SUFFIXES:
        return True
    return False


def walk_files(root: Path) -> list[Path]:
    files: list[Path] = []
    for p in root.rglob("*"):
        if not p.is_file():
            continue
        if should_exclude(p, root):
            continue
        files.append(p)
    return files


def make_zip(files: list[Path], base: Path, out_path: Path) -> None:
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(out_path, "w", compression=zipfile.ZIP_DEFLATED) as z:
        for f in sorted(files):
            arcname = f.relative_to(base).as_posix()
            z.write(str(f), arcname)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--path", default=os.getcwd(), help="Project directory (default: cwd)")
    ap.add_argument("--out", default="source_code_bundle.zip", help="Output zip filename")
    args = ap.parse_args()

    project_path = Path(args.path).expanduser().resolve()
    if not project_path.exists():
        print(f"ERROR: path does not exist: {project_path}", file=sys.stderr)
        return 2

    # Choose output path. If user passed a relative path, write it next to the script's cwd.
    out_path = Path(args.out).expanduser().resolve()

    if is_git_repo(project_path):
        repo = git_root(project_path)
        files = git_file_list(repo)
        if not files:
            print(f"WARNING: No files found in git repo at {repo}")
        make_zip(files, repo, out_path)
        print(f"Created: {out_path}")
        print(f"Bundled {len(files)} files from git repo: {repo}")
        return 0

    # Non-git fallback: zip everything (with exclusions).
    files = walk_files(project_path)
    if not files:
        print(f"WARNING: No files found under {project_path}")
    make_zip(files, project_path, out_path)
    print(f"Created: {out_path}")
    print(f"Bundled {len(files)} files from directory: {project_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
