from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
import tempfile
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Iterable

from pdf_to_md_ai import convert_pdf as convert_pdf_custom

try:
    from markitdown import MarkItDown
except ImportError:  # pragma: no cover
    MarkItDown = None

try:
    import pythoncom
    import win32com.client  # type: ignore
except ImportError:  # pragma: no cover
    pythoncom = None
    win32com = None


MARKITDOWN_EXTENSIONS = {
    ".docx",
    ".xlsx",
    ".xls",
    ".pptx",
    ".html",
    ".htm",
    ".csv",
    ".json",
    ".xml",
    ".txt",
    ".md",
    ".rtf",
    ".zip",
    ".epub",
    ".jpg",
    ".jpeg",
    ".png",
    ".gif",
    ".bmp",
    ".tif",
    ".tiff",
    ".wav",
    ".mp3",
}
LEGACY_OFFICE_EXTENSIONS = {".doc", ".xls", ".ppt"}
HWP_EXTENSIONS = {".hwp"}
PDF_EXTENSIONS = {".pdf"}

WORD_FORMAT_DOCX = 16
EXCEL_FORMAT_XLSX = 51
POWERPOINT_FORMAT_PPTX = 24


@dataclass
class ConversionResult:
    source: Path
    output: Path
    engine: str
    note: str = ""


def build_frontmatter(source_path: Path, engine: str, note: str = "") -> str:
    tags = ["documents", "markdown", "obsidian", engine.replace(" ", "-").lower()]
    lines = [
        "---",
        f'title: "{source_path.stem}"',
        f'aliases: ["{source_path.stem}"]',
        "tags:",
    ]
    lines.extend(f"  - {tag}" for tag in tags)
    lines.extend(
        [
            f'source_file: "{source_path.name}"',
            f'source_ext: "{source_path.suffix.lower()}"',
            f'converter: "{engine}"',
            f"converted_at: {date.today().isoformat()}",
        ]
    )
    if note:
        lines.append(f'note: "{note}"')
    lines.extend(
        [
            "---",
            "",
            f"# {source_path.stem}",
            "",
            "> [!abstract] Document Overview",
            f"> - Source file: `{source_path.name}`",
            f"> - Converter: `{engine}`",
        ]
    )
    if note:
        lines.append(f"> - Note: {note}")
    lines.extend(
        [
            "",
            "> [!tip] Obsidian Notes",
            "> - Add `[[links]]`, `#tags`, and checklists directly in this note.",
            "> - Review tables, images, and layout-sensitive sections manually when fidelity matters.",
            "",
        ]
    )
    return "\n".join(lines)


def ensure_markitdown() -> None:
    if MarkItDown is None:
        raise RuntimeError(
            "MarkItDown is not installed. Run `py -m pip install -r requirements.txt` first."
        )


def ensure_hwpjs() -> str:
    command = shutil.which("hwpjs")
    if command:
        return command
    # PyInstaller frozen exe may not inherit full PATH; check npm global dir directly
    if sys.platform == "win32":
        appdata = os.environ.get("APPDATA", "")
        if appdata:
            for name in ("hwpjs.cmd", "hwpjs.ps1", "hwpjs"):
                candidate = Path(appdata) / "npm" / name
                if candidate.exists():
                    return str(candidate)
    raise RuntimeError(
        "HWP 변환을 위해 hwpjs 설치가 필요합니다.\n"
        "1. Node.js 설치: https://nodejs.org\n"
        "2. 명령 프롬프트에서: npm install -g @ohah/hwpjs\n"
        "설치 후 프로그램을 다시 실행하면 HWP 변환이 가능합니다.\n"
        "(PDF, Word, Excel 등 다른 형식은 설치 없이 바로 사용 가능합니다)"
    )


def strip_existing_frontmatter(text: str) -> str:
    if not text.startswith("---\n"):
        return text
    parts = text.split("---\n", 2)
    if len(parts) < 3:
        return text
    return parts[2].lstrip()


def _normalize_path(path: Path) -> str:
    """Return a normalized string path that works with Korean/CJK characters on Windows."""
    import os, sys
    s = str(path.resolve())
    if sys.platform == "win32" and len(s) > 240 and not s.startswith("\\\\?\\"):
        s = "\\\\?\\" + s
    return s


def convert_with_markitdown(source_path: Path, output_path: Path) -> ConversionResult:
    ensure_markitdown()
    md = MarkItDown(enable_plugins=False)
    source_str = _normalize_path(source_path)
    try:
        result = md.convert(source_str)
    except (FileNotFoundError, OSError):
        # Fallback: copy to temp file with ASCII-safe path for Korean/CJK filenames
        with tempfile.NamedTemporaryFile(suffix=source_path.suffix, delete=False, prefix="docbridge_") as f:
            tmp_path = Path(f.name)
        try:
            shutil.copy2(source_str, str(tmp_path))
            result = md.convert(str(tmp_path))
        finally:
            tmp_path.unlink(missing_ok=True)
    content = build_frontmatter(source_path, "MarkItDown") + result.text_content
    output_path.write_text(content.strip() + "\n", encoding="utf-8")
    return ConversionResult(source=source_path, output=output_path, engine="MarkItDown")


def convert_hwp(source_path: Path, output_path: Path, image_dir: Path | None) -> ConversionResult:
    hwpjs_cmd = ensure_hwpjs()
    command = [hwpjs_cmd, "to-markdown", _normalize_path(source_path), "-o", _normalize_path(output_path)]
    note = "Converted via hwpjs"
    if image_dir is not None:
        image_dir.mkdir(parents=True, exist_ok=True)
        command.extend(["--images-dir", str(image_dir)])
        note = "Converted via hwpjs with extracted images"
    subprocess.run(command, check=True)
    body = output_path.read_text(encoding="utf-8")
    content = build_frontmatter(source_path, "hwpjs", note) + strip_existing_frontmatter(body)
    output_path.write_text(content.strip() + "\n", encoding="utf-8")
    return ConversionResult(source=source_path, output=output_path, engine="hwpjs", note=note)


def convert_legacy_office_to_modern(source_path: Path, temp_dir: Path) -> Path:
    suffix = source_path.suffix.lower()
    if win32com is None or pythoncom is None:
        raise RuntimeError(
            "pywin32 is not installed. Run `py -m pip install -r requirements.txt` first."
        )

    pythoncom.CoInitialize()
    app = None
    try:
        if suffix == ".doc":
            app = win32com.client.DispatchEx("Word.Application")
            app.Visible = False
            document = app.Documents.Open(_normalize_path(source_path))
            target = temp_dir / f"{source_path.stem}.docx"
            document.SaveAs(str(target), FileFormat=WORD_FORMAT_DOCX)
            document.Close(False)
            return target
        if suffix == ".xls":
            app = win32com.client.DispatchEx("Excel.Application")
            app.Visible = False
            workbook = app.Workbooks.Open(_normalize_path(source_path))
            target = temp_dir / f"{source_path.stem}.xlsx"
            workbook.SaveAs(str(target), FileFormat=EXCEL_FORMAT_XLSX)
            workbook.Close(False)
            return target
        if suffix == ".ppt":
            app = win32com.client.DispatchEx("PowerPoint.Application")
            presentation = app.Presentations.Open(_normalize_path(source_path), WithWindow=False)
            target = temp_dir / f"{source_path.stem}.pptx"
            presentation.SaveAs(str(target), FileFormat=POWERPOINT_FORMAT_PPTX)
            presentation.Close()
            return target
        raise RuntimeError(f"Unsupported legacy Office extension: {suffix}")
    finally:
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def convert_pdf(source_path: Path, output_path: Path, extract_images: bool) -> ConversionResult:
    convert_pdf_custom(
        Path(_normalize_path(source_path)),
        Path(_normalize_path(output_path)),
        add_page_markers=True,
        extract_images=extract_images,
    )
    return ConversionResult(source=source_path, output=output_path, engine="PyMuPDF custom")


def convert_one(source_path: Path, output_dir: Path, extract_images: bool) -> ConversionResult:
    source_path = Path(_normalize_path(source_path))
    output_dir = Path(_normalize_path(output_dir))
    output_path = output_dir / f"{source_path.stem}.md"
    suffix = source_path.suffix.lower()

    if suffix in PDF_EXTENSIONS:
        return convert_pdf(source_path, output_path, extract_images=extract_images)

    if suffix in HWP_EXTENSIONS:
        image_dir = output_dir / f"{source_path.stem}_assets" if extract_images else None
        return convert_hwp(source_path, output_path, image_dir=image_dir)

    if suffix in LEGACY_OFFICE_EXTENSIONS:
        with tempfile.TemporaryDirectory(prefix="docbridge_") as temp_name:
            temp_dir = Path(temp_name)
            modern_path = convert_legacy_office_to_modern(source_path, temp_dir)
            result = convert_with_markitdown(modern_path, output_path)
            result.source = source_path
            result.engine = "Microsoft Office COM + MarkItDown"
            result.note = f"Legacy {suffix} converted through installed Microsoft Office"
            body = strip_existing_frontmatter(output_path.read_text(encoding="utf-8"))
            output_path.write_text(
                build_frontmatter(source_path, result.engine, result.note) + body,
                encoding="utf-8",
            )
            return result

    if suffix in MARKITDOWN_EXTENSIONS:
        return convert_with_markitdown(source_path, output_path)

    raise RuntimeError(f"Unsupported extension: {suffix}")


def resolve_inputs(inputs: Iterable[str]) -> list[Path]:
    files: list[Path] = []
    for item in inputs:
        path = Path(item)
        if path.is_dir():
            for child in path.rglob("*"):
                if child.is_file():
                    files.append(child)
            continue
        if any(ch in item for ch in "*?[]"):
            files.extend(Path().glob(item))
            continue
        files.append(path)
    deduped: list[Path] = []
    seen: set[Path] = set()
    for file_path in files:
        if not file_path.exists():
            continue
        resolved = file_path.resolve()
        if resolved in seen:
            continue
        seen.add(resolved)
        deduped.append(file_path)
    return deduped


def is_supported(path: Path) -> bool:
    suffix = path.suffix.lower()
    return suffix in (PDF_EXTENSIONS | HWP_EXTENSIONS | LEGACY_OFFICE_EXTENSIONS | MARKITDOWN_EXTENSIONS)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Convert local documents into AI-friendly Markdown on Windows."
    )
    parser.add_argument("inputs", nargs="+", help="File(s), directories, or glob patterns")
    parser.add_argument("-o", "--output-dir", default="output_md", help="Markdown output directory")
    parser.add_argument("--no-images", action="store_true", help="Disable image extraction where supported")
    parser.add_argument("--skip-unsupported", action="store_true", help="Skip unsupported files instead of failing")
    args = parser.parse_args()

    paths = resolve_inputs(args.inputs)
    if not paths:
        parser.error("No input files were found.")

    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    failures = 0
    for path in paths:
        if not is_supported(path):
            message = f"[SKIP] {path.name} unsupported extension: {path.suffix.lower()}"
            if args.skip_unsupported:
                print(message)
                continue
            raise RuntimeError(message)
        try:
            result = convert_one(path, output_dir, extract_images=not args.no_images)
            note = f", note={result.note}" if result.note else ""
            print(f"[OK] {path.name} -> {result.output} (engine={result.engine}{note})")
        except Exception as exc:
            failures += 1
            print(f"[FAIL] {path.name}: {exc}", file=sys.stderr)

    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())
