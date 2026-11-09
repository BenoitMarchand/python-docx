"""Round-trip a Word file made by Word through python-docx and check nothing is lost.

Usage (from the repository root)::

    uv run python integration/check_roundtrip.py <file.docx|.docm|.dotx|.dotm> [...]

For each file: open it with python-docx, append the paragraph "Round-tripped by
python-docx", save it next to the source as ``<name>-pydocx<ext>``, then compare the
two packages:

* package kind (main document content-type) is unchanged,
* no part is lost or added,
* every non-XML part (images, ``vbaProject.bin``, ...) is byte-identical,
* the VBA project is still related to the document part,
* the macro source still contains the probe marker ``PYDOCX-MACRO-PROBE`` (searched in the
  decompressed VBA source when ``oletools`` is installed, otherwise reported as unchecked).

Then open ``<name>-pydocx<ext>`` in Word and run the ``PythonDocxCheck`` macro (Alt+F8):
it should report "Edited by python-docx: YES".

Exit code is 0 when every check passes, 1 otherwise.
"""

from __future__ import annotations

import re
import sys
import zipfile
from pathlib import Path

from docx import Document

ADDED_TEXT = "Round-tripped by python-docx"
PROBE_MARKER = "PYDOCX-MACRO-PROBE"
VBA_PART = "word/vbaProject.bin"
VBA_RELTYPE = "http://schemas.microsoft.com/office/2006/relationships/vbaProject"


def vba_source(blob: bytes) -> str | None:
    """Return the decompressed VBA source code in `blob`, or None if oletools is missing."""
    try:
        from oletools.olevba import VBA_Parser  # pyright: ignore
    except ImportError:
        return None
    parser = VBA_Parser("vbaProject.bin", data=blob)
    try:
        return "\n".join(code for _, _, _, code in parser.extract_macros())
    finally:
        parser.close()


def check(src: Path) -> list[str]:
    failures: list[str] = []

    def expect(name: str, ok: bool, detail: str = "") -> None:
        print(f"  [{'ok' if ok else 'FAIL'}] {name}{f' - {detail}' if detail and not ok else ''}")
        if not ok:
            failures.append(f"{src.name}: {name} {detail}")

    dst = src.with_name(f"{src.stem}-pydocx{src.suffix}")
    print(f"== {src}")

    with zipfile.ZipFile(src) as z:
        has_comments = "word/comments.xml" in z.namelist()

    document = Document(str(src))
    content_type = document.part.content_type
    # -- reading `.comments` creates an empty comments part when there is none, so only
    # -- count comments when the file already has them, to keep the round-trip faithful --
    n_comments = len(list(document.comments)) if has_comments else 0
    print(f"  content type: {content_type}")
    print(
        f"  paragraphs: {len(document.paragraphs)}, tables: {len(document.tables)}, "
        f"sections: {len(document.sections)}, pictures: {len(document.inline_shapes)}, "
        f"comments: {n_comments}"
    )
    document.add_paragraph(ADDED_TEXT)
    document.save(str(dst))
    print(f"  saved: {dst}")

    reloaded = Document(str(dst))
    expect("package kind preserved", reloaded.part.content_type == content_type)
    expect("added paragraph present", reloaded.paragraphs[-1].text == ADDED_TEXT)

    with zipfile.ZipFile(src) as zs, zipfile.ZipFile(dst) as zd:
        src_names, dst_names = set(zs.namelist()), set(zd.namelist())
        expect("no part lost", not (src_names - dst_names), str(sorted(src_names - dst_names)))
        expect("no part added", not (dst_names - src_names), str(sorted(dst_names - src_names)))

        binaries = sorted(
            n for n in src_names & dst_names if not n.endswith((".xml", ".rels", "/"))
        )
        changed = [n for n in binaries if zs.read(n) != zd.read(n)]
        expect(f"{len(binaries)} binary parts byte-identical", not changed, str(changed))

        if VBA_PART not in src_names:
            print("  [--] no VBA project in source file (no macro to check)")
            return failures

        rels = zd.read("word/_rels/document.xml.rels").decode("utf-8")
        expect("VBA project still related to document", VBA_RELTYPE in rels)
        types = zd.read("[Content_Types].xml").decode("utf-8")
        expect(
            "VBA project content type declared",
            bool(re.search(r"application/vnd\.ms-office\.vbaProject", types)),
        )

        source = vba_source(zd.read(VBA_PART))
        if source is None:
            print("  [--] macro source not inspected (pip install oletools to enable)")
        else:
            expect("macro source contains probe marker", PROBE_MARKER in source)
            subs = re.findall(r"^\s*(?:Public |Private )?Sub (\w+)", source, re.MULTILINE)
            print(f"  macros found: {subs}")

    return failures


def main(argv: list[str]) -> int:
    if not argv:
        print(__doc__)
        return 2
    failures = [f for arg in argv for f in check(Path(arg))]
    print(f"\n{'ALL CHECKS PASSED' if not failures else f'{len(failures)} check(s) FAILED'}")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
