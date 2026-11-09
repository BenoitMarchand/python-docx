"""Apply a few visible edits to Word files with python-docx, for a manual check in Word.

Usage (from the repository root)::

    uv run python integration/edit_samples.py <file> [<file> ...]

Each ``<name>.<ext>`` is saved as ``<name>-edited.<ext>`` (same kind of file). Edits:

1. a yellow-highlighted banner paragraph inserted at the very top,
2. the text of the first table's top-left cell replaced (bold, red),
3. a comment added on the banner,
4. a closing paragraph with bold/italic runs appended at the end,
5. the core properties title and "last modified by" changed.
"""

from __future__ import annotations

import sys
from pathlib import Path

from docx import Document
from docx.enum.text import WD_COLOR_INDEX
from docx.shared import Pt, RGBColor

BANNER = "EDITED BY PYTHON-DOCX – this banner was inserted at the top"


def edit(src: Path) -> Path:
    dst = src.with_name(f"{src.stem}-edited{src.suffix}")
    document = Document(str(src))

    # 1. banner at the very top
    first = document.paragraphs[0]
    banner = first.insert_paragraph_before()
    run = banner.add_run(BANNER)
    run.bold = True
    run.font.size = Pt(14)
    run.font.highlight_color = WD_COLOR_INDEX.YELLOW

    # 2. first table, top-left cell
    if document.tables:
        cell = document.tables[0].cell(0, 0)
        cell.text = ""
        r = cell.paragraphs[0].add_run("cell edited by python-docx")
        r.bold = True
        r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)

    # 3. comment on the banner
    document.add_comment(run, text="Comment added by python-docx", author="python-docx",
                         initials="PD")  # fmt: skip

    # 4. closing paragraph
    p = document.add_paragraph("Closing paragraph added by python-docx: ")
    p.add_run("bold").bold = True
    p.add_run(" and ")
    p.add_run("italic").italic = True
    p.add_run(f" (source file: {src.name}).")

    # 5. core properties
    document.core_properties.title = f"{src.stem} (edited by python-docx)"
    document.core_properties.last_modified_by = "python-docx"

    document.save(str(dst))

    reloaded = Document(str(dst))
    assert reloaded.part.content_type == document.part.content_type
    assert reloaded.paragraphs[0].text == BANNER
    print(f"{dst}  [{reloaded.part.content_type.split('/')[-1]}]")
    return dst


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(2)
    for arg in sys.argv[1:]:
        edit(Path(arg))
