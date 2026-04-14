from __future__ import annotations

import argparse
import re
import statistics
from collections import Counter
from dataclasses import dataclass
from datetime import date
from pathlib import Path

import fitz


KOREAN_RANGE = r"\u3131-\u318E\uAC00-\uD7A3"
HEADING_PREFIX_RE = re.compile(
    rf"^(PART\s*\d+|[IVX]+\.\s|[0-9]+\.\s|\[[^\]]+\]\s|[A-Za-z0-9{KOREAN_RANGE}]+\s*[:\uFF1A])"
)
TOC_RE = re.compile(r"^(?P<title>.+?)[\u00B7.]{4,}\s*(?P<page>[0-9ivxlcdmIVXLCDM-]+)$")
BULLET_RE = re.compile(r"^([\u2022\u25E6\u25AA\u25C6\u25C7\u25A0\u25A1\u25CB\u25CF\u203B\-]\s+)")
SPACE_RE = re.compile(r"[ \t]+")
PAGE_NUM_RE = re.compile(r"^\s*[\-\u2013]?\s*[0-9ivxlcdmIVXLCDM]+\s*[\-\u2013]?\s*$")


@dataclass
class Line:
    text: str
    size: float
    is_bold: bool
    y: float


def clean_text(text: str) -> str:
    text = text.replace("\u00ad", "")
    text = text.replace("\uf000", "\u2022")
    text = text.replace("\uff65", "\u00B7")
    text = text.replace("\u3000", " ")
    text = text.replace("  ", " ")
    return SPACE_RE.sub(" ", text).strip()


def extract_lines(page: fitz.Page) -> list[Line]:
    data = page.get_text("dict")
    lines: list[Line] = []
    for block in data.get("blocks", []):
        if "lines" not in block:
            continue
        for raw_line in block["lines"]:
            spans = raw_line.get("spans", [])
            parts: list[str] = []
            sizes: list[float] = []
            bold_flags: list[bool] = []
            y = None
            for span in spans:
                text = clean_text(span.get("text", ""))
                if not text:
                    continue
                parts.append(text)
                sizes.append(float(span.get("size", 0.0)))
                font_name = span.get("font", "").lower()
                flags = int(span.get("flags", 0))
                bold_flags.append("bold" in font_name or bool(flags & 16))
                y = raw_line["bbox"][1]
            if not parts:
                continue
            line_text = " ".join(parts)
            line_text = re.sub(r"\s+([,.;:%])", r"\1", line_text)
            lines.append(
                Line(
                    text=line_text,
                    size=max(sizes) if sizes else 0.0,
                    is_bold=any(bold_flags),
                    y=float(y or 0.0),
                )
            )
    return lines


def body_font_size(lines: list[Line]) -> float:
    sizes = [round(line.size, 1) for line in lines if len(line.text) >= 8]
    if not sizes:
        return 11.0
    return Counter(sizes).most_common(1)[0][0]


def garbage_ratio(text: str) -> float:
    if not text:
        return 1.0
    weird = sum(
        1
        for ch in text
        if ch == "\ufffd" or ("\ue000" <= ch <= "\uf8ff") or ord(ch) < 9
    )
    return weird / max(len(text), 1)


def is_page_number(text: str) -> bool:
    return bool(PAGE_NUM_RE.fullmatch(text))


def is_heading(line: Line, body_size: float) -> bool:
    text = line.text
    if len(text) > 120:
        return False
    if HEADING_PREFIX_RE.match(text):
        return True
    if line.size >= body_size + 2.0:
        return True
    if line.is_bold and line.size >= body_size + 0.8 and len(text) <= 60:
        return True
    return False


def merge_paragraphs(lines: list[Line], body_size: float) -> list[str]:
    chunks: list[str] = []
    paragraph: list[str] = []
    previous_y = None

    def flush_paragraph() -> None:
        nonlocal paragraph
        if paragraph:
            joined = " ".join(paragraph)
            joined = re.sub(r"\s+([,.;:%])", r"\1", joined)
            chunks.append(joined.strip())
            paragraph = []

    for line in lines:
        text = clean_text(line.text)
        if not text or is_page_number(text):
            continue

        toc_match = TOC_RE.match(text)
        if toc_match:
            flush_paragraph()
            chunks.append(f"- {toc_match.group('title').strip()} (p. {toc_match.group('page')})")
            previous_y = line.y
            continue

        if is_heading(line, body_size):
            flush_paragraph()
            level = "##" if line.size >= body_size + 4 else "###"
            chunks.append(f"{level} {text}")
            previous_y = line.y
            continue

        if BULLET_RE.match(text):
            flush_paragraph()
            bullet_text = BULLET_RE.sub("", text, count=1).strip()
            chunks.append(f"- {bullet_text}")
            previous_y = line.y
            continue

        new_paragraph = False
        if previous_y is not None and line.y - previous_y > body_size * 1.5:
            new_paragraph = True
        if line.size >= body_size + 1.5 and len(text) < 50:
            new_paragraph = True

        if new_paragraph:
            flush_paragraph()

        if paragraph and paragraph[-1].endswith((".", "?", "!", ":", "\u2026")):
            flush_paragraph()

        paragraph.append(text)
        previous_y = line.y

    flush_paragraph()
    return chunks


def save_page_images(page: fitz.Page, image_dir: Path, doc_slug: str) -> list[str]:
    image_refs: list[str] = []
    seen_xrefs: set[int] = set()

    for image_index, image_info in enumerate(page.get_images(full=True), start=1):
        xref = int(image_info[0])
        if xref in seen_xrefs:
            continue
        seen_xrefs.add(xref)

        pix = fitz.Pixmap(page.parent, xref)
        try:
            if pix.width < 80 or pix.height < 80:
                continue
            if pix.colorspace is None:
                continue
            if pix.colorspace.name not in ("DeviceGray", "DeviceRGB") or pix.alpha or pix.colorspace.n > 3:
                pix = fitz.Pixmap(fitz.csRGB, pix)

            image_name = f"{doc_slug}_page{page.number + 1:04d}_img{image_index:02d}.png"
            image_path = image_dir / image_name
            try:
                image_path.write_bytes(pix.tobytes("png"))
            except Exception:
                pix = fitz.Pixmap(fitz.csRGB, pix)
                image_path.write_bytes(pix.tobytes("png"))
            image_refs.append(image_name)
        finally:
            pix = None

    return image_refs


def build_image_blocks(image_names: list[str], rel_image_dir: str, page_number: int) -> list[str]:
    if not image_names:
        return []

    blocks = ["## Extracted Images", ""]
    for idx, image_name in enumerate(image_names, start=1):
        image_path = f"{rel_image_dir}/{image_name}"
        blocks.extend(
            [
                f"![page-{page_number}-image-{idx}]({image_path})",
                "",
                "> [!note] Image Notes",
                f"> - Page: {page_number}",
                f"> - Image: {idx}",
                "> - Type guess: chart / table / diagram / figure",
                "> - Key values:",
                "> - Summary:",
                "> - AI interpretation notes:",
                "",
            ]
        )
    return blocks


def convert_pdf(
    pdf_path: Path,
    output_path: Path,
    add_page_markers: bool,
    extract_images: bool,
) -> dict[str, float | int]:
    # Read as bytes to avoid C library (MuPDF) encoding issues with Korean/CJK file paths
    doc = fitz.open(stream=pdf_path.read_bytes(), filetype="pdf")
    page_texts: list[str] = []
    body_sizes: list[float] = []
    all_text: list[str] = []
    total_images = 0

    doc_slug = output_path.stem
    image_dir = output_path.parent / f"{doc_slug}_assets"
    rel_image_dir = image_dir.name
    if extract_images:
        image_dir.mkdir(parents=True, exist_ok=True)

    for page in doc:
        lines = extract_lines(page)
        body_size = body_font_size(lines)
        body_sizes.append(body_size)
        page_blocks = merge_paragraphs(lines, body_size)
        image_names = save_page_images(page, image_dir, doc_slug) if extract_images else []
        total_images += len(image_names)

        if add_page_markers:
            page_texts.append(f"\n<!-- page: {page.number + 1} -->\n")
        page_texts.extend(page_blocks)
        if image_names:
            page_texts.append("")
            page_texts.extend(build_image_blocks(image_names, rel_image_dir, page.number + 1))
        page_texts.append("")
        all_text.append(page.get_text("text"))

    title = pdf_path.stem
    content = "\n".join(page_texts).strip() + "\n"
    full_text = "\n".join(all_text)
    ratio = garbage_ratio(full_text)
    today = date.today().isoformat()

    frontmatter = [
        "---",
        f'title: "{title}"',
        f'aliases: ["{pdf_path.stem}"]',
        "tags:",
        "  - pdf",
        "  - markdown",
        "  - ai-ready",
        "  - obsidian",
        f'source_pdf: "{pdf_path.name}"',
        f"total_pages: {len(doc)}",
        'extraction_engine: "PyMuPDF"',
        f"estimated_body_font_size: {statistics.median(body_sizes):.1f}",
        f"garbled_char_ratio: {ratio:.4f}",
        f"converted_at: {today}",
        f"extracted_images: {total_images}",
    ]
    if ratio > 0.01:
        frontmatter.append('warning: "Extracted text may include broken font mapping; OCR fallback may be needed."')
    frontmatter.extend(["---", ""])

    header = [
        f"# {title}",
        "",
        "> [!abstract] Document Overview",
        f"> - Source PDF: `{pdf_path.name}`",
        f"> - Total pages: {len(doc)}",
        f"> - Extraction engine: `PyMuPDF`",
        f"> - Estimated body font size: {statistics.median(body_sizes):.1f}",
        f"> - Garbled text ratio: {ratio:.4f}",
        f"> - Extracted images: {total_images}",
    ]
    if ratio > 0.01:
        header.append("> - Warning: text extraction quality looks degraded; OCR fallback may be needed.")
    header.extend(
        [
            "",
            "> [!tip] Obsidian Notes",
            "> - Add `[[links]]`, `#tags`, and checklists directly in this note.",
            "> - Use `<!-- page: N -->` markers when you need page-level references.",
            "> - Each extracted image includes a placeholder note block for chart/table interpretation.",
            "",
        ]
    )

    output_path.write_text("\n".join(frontmatter + header) + content, encoding="utf-8")
    return {"pages": len(doc), "garbled_ratio": ratio, "images": total_images}


def resolve_inputs(inputs: list[str]) -> list[Path]:
    results: list[Path] = []
    for item in inputs:
        path = Path(item)
        if path.is_dir():
            results.extend(sorted(path.glob("*.pdf")))
            continue
        if any(ch in item for ch in "*?[]"):
            results.extend(sorted(Path().glob(item)))
            continue
        results.append(path)
    return [path for path in results if path.suffix.lower() == ".pdf" and path.exists()]


def main() -> int:
    parser = argparse.ArgumentParser(description="Convert PDF files into AI-friendly Markdown.")
    parser.add_argument("inputs", nargs="+", help="PDF file(s), directory, or glob pattern")
    parser.add_argument("-o", "--output-dir", default="output_md", help="Directory where Markdown files will be written")
    parser.add_argument("--no-page-markers", action="store_true", help="Do not include HTML page markers in the Markdown output")
    parser.add_argument("--no-images", action="store_true", help="Do not extract embedded PDF images into Markdown")
    args = parser.parse_args()

    pdf_paths = resolve_inputs(args.inputs)
    if not pdf_paths:
        parser.error("No PDF files were found.")

    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    for pdf_path in pdf_paths:
        output_path = output_dir / f"{pdf_path.stem}.md"
        stats = convert_pdf(
            pdf_path,
            output_path,
            add_page_markers=not args.no_page_markers,
            extract_images=not args.no_images,
        )
        print(
            f"[OK] {pdf_path.name} -> {output_path} "
            f"(pages={stats['pages']}, images={stats['images']}, garbled_ratio={stats['garbled_ratio']:.4f})"
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
