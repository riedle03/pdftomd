"""Interactive batch converter — scan a folder, pick files, convert to Markdown."""

from __future__ import annotations

import argparse
import sys
import time
from pathlib import Path

from convert_to_md import (
    MARKITDOWN_EXTENSIONS,
    LEGACY_OFFICE_EXTENSIONS,
    HWP_EXTENSIONS,
    PDF_EXTENSIONS,
    convert_one,
    is_supported,
)

ALL_SUPPORTED = PDF_EXTENSIONS | HWP_EXTENSIONS | LEGACY_OFFICE_EXTENSIONS | MARKITDOWN_EXTENSIONS

EXT_GROUPS = {
    "pdf": PDF_EXTENSIONS,
    "hwp": HWP_EXTENSIONS,
    "office": {".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx"},
    "image": {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tif", ".tiff"},
    "web": {".html", ".htm", ".csv", ".json", ".xml"},
    "text": {".txt", ".md", ".rtf"},
}


def scan_folder(folder: Path, recursive: bool) -> list[Path]:
    pattern = "**/*" if recursive else "*"
    files = sorted(
        (f for f in folder.glob(pattern) if f.is_file() and is_supported(f)),
        key=lambda p: (p.suffix.lower(), p.name.lower()),
    )
    return files


def ext_label(path: Path) -> str:
    return path.suffix.lower().lstrip(".")


def format_size(size: int) -> str:
    if size < 1024:
        return f"{size} B"
    if size < 1024 * 1024:
        return f"{size / 1024:.1f} KB"
    return f"{size / (1024 * 1024):.1f} MB"


def print_file_list(files: list[Path], folder: Path) -> None:
    if not files:
        print("  (no supported files found)")
        return

    max_idx_width = len(str(len(files)))
    for i, f in enumerate(files, 1):
        rel = f.relative_to(folder) if f.is_relative_to(folder) else f
        size = format_size(f.stat().st_size)
        ext = ext_label(f)
        print(f"  {i:>{max_idx_width}}. [{ext:>4}] {rel}  ({size})")


def print_ext_summary(files: list[Path]) -> None:
    counts: dict[str, int] = {}
    for f in files:
        ext = ext_label(f)
        counts[ext] = counts.get(ext, 0) + 1
    parts = [f".{ext}: {cnt}" for ext, cnt in sorted(counts.items(), key=lambda x: -x[1])]
    print(f"  Total: {len(files)} files  |  {', '.join(parts)}")


def parse_selection(text: str, total: int) -> list[int] | None:
    """Parse user selection into 0-based indices. Returns None for 'all'."""
    text = text.strip().lower()
    if text in ("a", "all", "*", ""):
        return None  # means all

    indices: set[int] = set()
    for part in text.replace(",", " ").split():
        # extension filter: e.g. ".pdf" or "pdf"
        if part.startswith(".") or part in EXT_GROUPS:
            return part  # handled separately

        # range: e.g. "3-7"
        if "-" in part:
            tokens = part.split("-", 1)
            try:
                start, end = int(tokens[0]), int(tokens[1])
                for n in range(start, end + 1):
                    if 1 <= n <= total:
                        indices.add(n - 1)
            except ValueError:
                pass
            continue

        # single number
        try:
            n = int(part)
            if 1 <= n <= total:
                indices.add(n - 1)
        except ValueError:
            pass

    return sorted(indices)


def select_by_ext(files: list[Path], ext_input: str) -> list[int]:
    ext_input = ext_input.lstrip(".").lower()
    if ext_input in EXT_GROUPS:
        target_exts = EXT_GROUPS[ext_input]
    else:
        target_exts = {f".{ext_input}"}
    return [i for i, f in enumerate(files) if f.suffix.lower() in target_exts]


def confirm_selection(files: list[Path], indices: list[int] | None, folder: Path) -> list[Path]:
    if indices is None:
        selected = files
    else:
        selected = [files[i] for i in indices]

    print(f"\n  Selected {len(selected)} file(s) for conversion:")
    for f in selected[:10]:
        rel = f.relative_to(folder) if f.is_relative_to(folder) else f
        print(f"    - {rel}")
    if len(selected) > 10:
        print(f"    ... and {len(selected) - 10} more")

    answer = input("\n  Proceed? (Y/n): ").strip().lower()
    if answer in ("", "y", "yes"):
        return selected
    return []


def run_batch(files: list[Path], output_dir: Path, extract_images: bool) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    total = len(files)
    ok_count = 0
    fail_count = 0
    skip_count = 0
    failed_files: list[tuple[Path, str]] = []

    start_time = time.time()

    for i, path in enumerate(files, 1):
        prefix = f"[{i}/{total}]"
        try:
            result = convert_one(path, output_dir, extract_images=extract_images)
            ok_count += 1
            note = f", {result.note}" if result.note else ""
            print(f"  {prefix} OK   {path.name} -> {result.output.name} ({result.engine}{note})")
        except Exception as exc:
            fail_count += 1
            failed_files.append((path, str(exc)))
            print(f"  {prefix} FAIL {path.name}: {exc}", file=sys.stderr)

    elapsed = time.time() - start_time

    print(f"\n{'=' * 60}")
    print(f"  Batch conversion complete")
    print(f"  Output : {output_dir.resolve()}")
    print(f"  Time   : {elapsed:.1f}s")
    print(f"  Success: {ok_count}  |  Failed: {fail_count}  |  Total: {total}")
    if failed_files:
        print(f"\n  Failed files:")
        for path, err in failed_files:
            print(f"    - {path.name}: {err}")
    print(f"{'=' * 60}")


def interactive_mode(folder: Path, output_dir: Path, recursive: bool, extract_images: bool) -> int:
    print(f"\n{'=' * 60}")
    print(f"  Batch Document-to-Markdown Converter")
    print(f"{'=' * 60}")
    print(f"  Scanning: {folder.resolve()}")
    print(f"  Mode    : {'recursive' if recursive else 'top-level only'}\n")

    files = scan_folder(folder, recursive)
    if not files:
        print("  No supported files found in the specified folder.")
        return 1

    print_ext_summary(files)
    print()
    print_file_list(files, folder)

    print(f"\n  Selection options:")
    print(f"    a / Enter  = all files")
    print(f"    1 3 5      = specific numbers (space or comma separated)")
    print(f"    2-8        = range")
    print(f"    pdf        = all .pdf files")
    print(f"    office     = all Office files (.doc/.docx/.xls/.xlsx/.ppt/.pptx)")
    print(f"    hwp        = all .hwp files")
    print(f"    .docx      = specific extension")

    selection_input = input("\n  Select files to convert: ").strip()

    # Check if it's an extension/group filter
    sel = parse_selection(selection_input, len(files))
    if isinstance(sel, str):
        sel = select_by_ext(files, sel)

    if isinstance(sel, list) and len(sel) == 0:
        print("  No files matched the selection.")
        return 1

    selected = confirm_selection(files, sel, folder)
    if not selected:
        print("  Cancelled.")
        return 0

    print()
    run_batch(selected, output_dir, extract_images)
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Interactive batch converter — scan a folder, pick files, convert to Markdown."
    )
    parser.add_argument("folder", nargs="?", default=".", help="Folder to scan (default: current directory)")
    parser.add_argument("-o", "--output-dir", default="output_md", help="Markdown output directory")
    parser.add_argument("-r", "--recursive", action="store_true", help="Scan subfolders recursively")
    parser.add_argument("--no-images", action="store_true", help="Disable image extraction")
    parser.add_argument("--auto", action="store_true", help="Skip interactive selection — convert all supported files")
    args = parser.parse_args()

    folder = Path(args.folder).resolve()
    if not folder.is_dir():
        print(f"Error: '{args.folder}' is not a directory.", file=sys.stderr)
        return 1

    output_dir = Path(args.output_dir)
    extract_images = not args.no_images

    if args.auto:
        files = scan_folder(folder, args.recursive)
        if not files:
            print("No supported files found.")
            return 1
        print(f"\n  Auto mode: converting {len(files)} file(s)...\n")
        run_batch(files, output_dir, extract_images)
        return 0

    return interactive_mode(folder, output_dir, args.recursive, extract_images)


if __name__ == "__main__":
    raise SystemExit(main())
