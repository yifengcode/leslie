#!/usr/bin/env python3
"""Convert a Microsoft Word .docx document to Markdown."""

import argparse
import re
import sys
from pathlib import Path

from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph


def _escape_markdown(text):
    return re.sub(r"([\\`*_{}\[\]<>])", r"\\\1", text)


def _render_paragraph(paragraph):
    parts = []
    for run in paragraph.runs:
        text = _escape_markdown(run.text).replace("\n", "<br>")
        if run.bold:
            text = f"**{text}**"
        if run.italic:
            text = f"*{text}*"
        if run.font.strike:
            text = f"~~{text}~~"
        parts.append(text)

    text = "".join(parts).strip()
    if not text:
        return ""

    style_name = paragraph.style.name if paragraph.style else ""
    heading = re.fullmatch(r"Heading ([1-6])", style_name)
    if heading:
        return f'{"#" * int(heading.group(1))} {text}'
    if style_name == "Title":
        return f"# {text}"
    if "List Bullet" in style_name or "List Number" in style_name:
        level_match = re.search(r"(\d+)$", style_name)
        indent = "  " * (int(level_match.group(1)) - 1) if level_match else ""
        marker = "1." if "List Number" in style_name else "-"
        return f"{indent}{marker} {text}"
    if style_name == "Quote":
        return f"> {text}"
    return text


def _escape_table_cell(text):
    return text.replace("\\", "\\\\").replace("|", "\\|").replace("\n", "<br>")


def _render_table(table):
    rows = [
        [_escape_table_cell(cell.text.strip()) for cell in row.cells]
        for row in table.rows
    ]
    if not rows:
        return ""

    width = max(len(row) for row in rows)
    rows = [row + [""] * (width - len(row)) for row in rows]
    lines = [
        "| " + " | ".join(rows[0]) + " |",
        "| " + " | ".join("---" for _ in range(width)) + " |",
    ]
    lines.extend("| " + " | ".join(row) + " |" for row in rows[1:])
    return "\n".join(lines)


def convert_document(input_path):
    """Return the Markdown representation of a .docx file."""
    document = Document(input_path)
    blocks = []
    for item in document.iter_inner_content():
        if isinstance(item, Paragraph):
            block = _render_paragraph(item)
        elif isinstance(item, Table):
            block = _render_table(item)
        else:
            continue
        if block:
            blocks.append(block)
    return "\n\n".join(blocks) + ("\n" if blocks else "")


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="Convert a .docx Word document to Markdown."
    )
    parser.add_argument("input", type=Path, help="path to the .docx document")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        help="Markdown output path (default: input path with a .md extension)",
    )
    args = parser.parse_args(argv)
    output_path = args.output or args.input.with_suffix(".md")
    if args.input.resolve() == output_path.resolve():
        parser.error("output path must not overwrite the input document")

    try:
        markdown = convert_document(args.input)
        output_path.write_text(markdown, encoding="utf-8")
    except (OSError, ValueError, KeyError) as error:
        parser.exit(1, f"Error converting document: {error}\n")

    return 0


if __name__ == "__main__":
    sys.exit(main())
