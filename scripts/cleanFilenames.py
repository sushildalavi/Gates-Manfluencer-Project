"""
Two-pass file cleanup across the full codebase:
  1. Convert every .csv to .xlsx (same folder, same base name)
  2. Rename every data file whose basename contains '_' → replace '_' with ' '
     (skips .py and .ipynb — renaming those breaks imports)

Run from project root.
"""
from __future__ import annotations

import os
import shutil
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]

SKIP_EXTENSIONS = {".py", ".ipynb"}   # never rename code files
SKIP_DIRS = {"__pycache__", ".git", ".ipynb_checkpoints"}


def should_skip(path: Path) -> bool:
    return any(s in path.parts for s in SKIP_DIRS)


# ── Pass 1: CSV → xlsx ────────────────────────────────────────────────────────

def convert_csv_to_xlsx(path: Path) -> Path | None:
    """Convert one CSV to xlsx, return new path or None on failure."""
    new_path = path.with_suffix(".xlsx")
    if new_path.exists():
        # avoid overwriting — suffix with _converted
        new_path = path.with_name(path.stem + "_converted.xlsx")
    try:
        df = pd.read_csv(path, low_memory=False)
        df.to_excel(new_path, index=False)
        path.unlink()          # delete the csv
        return new_path
    except Exception as e:
        print(f"  ✗ CSV convert failed: {path.relative_to(ROOT)} — {e}")
        return None


def pass1_convert_csvs():
    print("\n=== PASS 1: Convert CSV → xlsx ===")
    csvs = [p for p in ROOT.rglob("*.csv") if not should_skip(p)]
    print(f"  Found {len(csvs)} CSV files")
    converted, failed = 0, 0
    for p in csvs:
        result = convert_csv_to_xlsx(p)
        if result:
            print(f"  ✓ {p.relative_to(ROOT)}  →  {result.name}")
            converted += 1
        else:
            failed += 1
    print(f"\n  Done: {converted} converted, {failed} failed")
    return converted


# ── Pass 2: Rename files with underscores ─────────────────────────────────────

def safe_rename(path: Path) -> Path | None:
    """Rename file replacing '_' with ' ' in basename. Returns new path or None."""
    new_name = path.name.replace("_", " ")
    if new_name == path.name:
        return None   # no underscores, skip
    new_path = path.parent / new_name
    if new_path.exists():
        print(f"  ! CONFLICT (target exists): {new_path.relative_to(ROOT)} — skipping")
        return None
    try:
        path.rename(new_path)
        return new_path
    except Exception as e:
        print(f"  ✗ Rename failed: {path.relative_to(ROOT)} — {e}")
        return None


def pass2_rename_underscores():
    print("\n=== PASS 2: Remove underscores from filenames ===")
    # All files except code files
    all_files = [
        p for p in ROOT.rglob("*")
        if p.is_file()
        and not should_skip(p)
        and p.suffix.lower() not in SKIP_EXTENSIONS
        and "_" in p.name
    ]
    print(f"  Found {len(all_files)} files with underscores")
    renamed, skipped = 0, 0
    for p in sorted(all_files):
        result = safe_rename(p)
        if result:
            print(f"  ✓ {p.name}  →  {result.name}")
            renamed += 1
        else:
            skipped += 1
    print(f"\n  Done: {renamed} renamed, {skipped} skipped (conflicts or no underscores)")
    return renamed


# ── Verification ──────────────────────────────────────────────────────────────

def verify():
    print("\n=== VERIFICATION ===")
    remaining_csvs = [p for p in ROOT.rglob("*.csv") if not should_skip(p)]
    remaining_underscores = [
        p for p in ROOT.rglob("*")
        if p.is_file()
        and not should_skip(p)
        and p.suffix.lower() not in SKIP_EXTENSIONS
        and "_" in p.name
    ]
    print(f"  Remaining CSVs:              {len(remaining_csvs)}")
    for p in remaining_csvs:
        print(f"    {p.relative_to(ROOT)}")
    print(f"  Files still with underscores: {len(remaining_underscores)}")
    for p in remaining_underscores[:20]:
        print(f"    {p.relative_to(ROOT)}")
    if len(remaining_underscores) > 20:
        print(f"    ... and {len(remaining_underscores)-20} more")


if __name__ == "__main__":
    print(f"Root: {ROOT}")
    pass1_convert_csvs()
    pass2_rename_underscores()
    verify()
    print("\n=== ALL DONE ===")
