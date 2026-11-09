# CLAUDE.md

Guidance for working in this repository (a clone of
[python-openxml/python-docx](https://github.com/python-openxml/python-docx)).

## Project overview

`python-docx` is a pure-Python library (MIT) to create, read and update Microsoft Word
`.docx` files. Version lives in `src/docx/__init__.py` (`__version__`, currently 1.2.0).
Runtime deps are only `lxml` and `typing_extensions`. Requires Python >= 3.9; tox tests
3.9 – 3.13.

- Docs: https://python-docx.readthedocs.io/en/latest/ (built from `docs/` with Sphinx)
  - User guide: `docs/user/` (quickstart, text, tables, styles, sections, headers/footers, comments)
  - API reference: `docs/api/`
  - Design analyses: `docs/dev/analysis/` (rendered under "Contributor Guide" → "Analysis");
    read the matching page before changing a feature, it documents Word behaviour + XML.
- Changelog: `HISTORY.rst` (update it for every user-visible change)

## Architecture (layers, top to bottom)

```
docx.api.Document()            entry point: opens a .docx (or the default template)
  │
  ├─ Proxy / domain objects     src/docx/{document,section,table,shape,comments,settings}.py,
  │                             src/docx/text/, src/docx/styles/, src/docx/dml/, src/docx/drawing/
  │   Thin wrappers around an lxml element (e.g. Paragraph wraps CT_P as self._p / self._element).
  │   They hold NO state of their own – everything is read from/written to the XML.
  │
  ├─ oxml layer                 src/docx/oxml/
  │   Custom lxml element classes named after the schema type: CT_P, CT_PPr, CT_Tbl, ...
  │   Declared with the "xmlchemy" DSL (oxml/xmlchemy.py): ZeroOrOne, ZeroOrMore, OneAndOnlyOne,
  │   OptionalAttribute, RequiredAttribute, ... which auto-generate get_or_add_x(), _add_x(),
  │   _remove_x(), _insert_x() etc. Value conversion via oxml/simpletypes.py (ST_*).
  │   Every class must be registered in oxml/__init__.py with register_element_cls("w:tag", CT_X).
  │
  ├─ Parts                      src/docx/parts/ (DocumentPart, StylesPart, CommentsPart, HeaderPart...)
  │
  └─ OPC packaging              src/docx/opc/ (zip package, rels, content types, core properties)
```

Other notable modules:
- `docx/shared.py` – `Length` units (`Inches`, `Pt`, `Cm`, `Emu`, `Twips`), `RGBColor`,
  `lazyproperty`, `ElementProxy`, `StoryChild`, `Parented`.
- `docx/enum/` – `WD_*` enumerations (mirror Word VBA enumerations).
- `docx/templates/` – default `.docx` template and XML templates for new parts.
- `docx/types.py` – protocol types (e.g. `ProvidesStoryPart`).
- `typings/` – local type stubs (behave). `ref/` – reference XSD schemas.

### Adding a feature – the usual path
1. Analysis page in `docs/dev/analysis/features/...` (Word behaviour, sample XML, relevant
   schema excerpt from ISO/IEC 29500). Maintainer expects this for non-trivial features.
2. Acceptance test (`features/*.feature` + `features/steps/*.py`), initially failing
   (commit prefix `xfail:`).
3. oxml: add/extend `CT_*` class with xmlchemy declarations, register it.
4. Proxy: public property/method on the domain object delegating to the oxml element.
5. Unit tests for each layer, API docs in `docs/api/`, user docs in `docs/user/`, `HISTORY.rst`.

### Definition of done for every issue (mandatory)
An issue is solved only when the change includes all three of the following:
1. **Code** – the fix or feature in `src/docx/`, following the layering and coding rules below
   (ruff and strict pyright clean).
2. **Tests, written in the same style as the existing ones** – unit tests in the matching
   `tests/` module (`Describe*` classes, `it_*` functions, CXML fixtures, `unitutil.mock`
   helpers, parametrize for variants) that fail before the fix and pass after; plus a behave
   scenario in `features/` when public API behaviour is added or changed. Reproduce the
   reported bug in a test first. The full suite (`pytest` + `behave`) must pass.
3. **Documentation** – update the docstrings, `docs/api/` and/or `docs/user/` for anything the
   user sees, `docs/dev/analysis/` for new features, and add an entry to `HISTORY.rst`
   (e.g. `- Fix #1638 Zero page margin treated as 1 inch in table width`).

Don't consider an issue done, or open a PR, while any of the three is missing.

## Commands

The project uses [uv](https://docs.astral.sh/uv/) (`uv.lock`). On Windows `make` may be
unavailable – call the underlying commands directly.

```bash
uv sync                                   # create .venv with dev deps
uv run pytest -x                          # unit tests            (make test)
uv run pytest tests/text/test_paragraph.py::DescribeParagraph::it_knows_its_paragraph_style
uv run behave --stop                      # acceptance tests      (make accept)
uv run behave --tags=-wip                 # skip work-in-progress scenarios
uv run ruff check . && uv run ruff format .   # lint / format
uv run pyright                            # strict type check
uv run tox                                # full matrix (py39..py313: pytest + behave)
make -C docs html                         # build docs (docs/.build/html)
uv build                                  # sdist + wheel
```

## Coding rules

- **Style**: ruff, line length **100**, isort with `docx` first-party. Docstrings on public
  API use reST/Sphinx (`:ref:`, `|Paragraph|` substitutions); the first line is a summary.
  Comments often use the `# -- comment --` form.
- **Typing**: pyright **strict** on `src/docx` and `tests`. Use `from __future__ import
  annotations`; put import-only-for-typing in `if TYPE_CHECKING:`. xmlchemy-generated
  attributes are declared as class annotations (e.g. `get_or_add_ind: Callable[[], CT_Ind]`)
  and descriptor assignments carry `# pyright: ignore[reportAssignmentType]`. Do not add
  unnecessary casts or `type: ignore` (both are reported as errors).
- **Proxies never cache XML state**; read through to the element each time. `None` on a
  property typically means "not directly applied, inherited from style hierarchy" and
  assigning `None` removes the element/attribute.
- Keep the public API backward-compatible; when moving things, leave aliases (see #1256
  where moving `parse_xml`/`OxmlElement` broke users).
- Be tolerant of real-world/invalid XML when reading (several bugs: invalid floats in
  margins #1335/#1539, `w:highlight w:val="none"` #1559, empty core-property dates #1622,
  zero DPI images #1494). Prefer returning `None` or a sane default over raising.
- No new runtime dependencies without a very strong reason.

## Testing rules

- **pytest** with BDD-style naming (configured in `pyproject.toml`):
  classes `Describe<Thing>`, test functions `it_...`, `its_...`, `they_...`, `and_...`,
  `but_...`. Files `test_*.py`, mirroring `src/docx` layout under `tests/`.
- **Warnings are errors** (`filterwarnings = error`).
- Build XML fixtures with **CXML** (`tests/unitutil/cxml.py`): `element('w:p/w:r/w:t"foo"')`,
  `xml(...)` for expected output. Use `@pytest.mark.parametrize` for XML variants.
- Mocks via `tests/unitutil/mock.py` helpers (`instance_mock`, `class_mock`, `method_mock`,
  `property_mock`), each taking the pytest `request` fixture.
- Fixtures live at the bottom of the test class under `# fixtures ----` /
  `# fixture components ----` headings. Newer code prefers plain parametrized tests over
  the old `*_fixture` tuple style; class-scoped fixtures should not be instance methods (#1564).
- Acceptance tests: behave, `.feature` files named `<area>-<topic>.feature`
  (`par-`, `run-`, `tbl-`, `sty-`, `doc-`, `sct-`, `hdr-`, `cmt-`, ...); test documents in
  `features/steps/test_files/`; output written to `_scratch/`.

## Git / commit conventions

Commit subjects are prefixed by area/kind, lower-case:
`fix:`, `feat(table):`, `comments:`, `rfctr:`, `rfctr(lint):`, `docs:`, `build:`, `modn:`
(modernization), `xfail:` (failing acceptance test added before the implementation),
`release:`. Reference issue numbers as `Fix #1234 ...` in `HISTORY.rst`.

Branches: `master` (releases), `develop`, plus `feature/*` branches upstream.

### This clone's fork workflow
- Remotes: `origin` = fork https://github.com/BenoitMarchand/python-docx (push here),
  `upstream` = python-openxml/python-docx (fetch only; push URL is set to `DISABLED`).
- `master` mirrors `upstream/master` – never commit on it. Sync with
  `git fetch upstream && git switch master && git merge --ff-only upstream/master && git push`.
- One topic branch per change, created from `master` (e.g. `fix/1638-zero-margin-width`),
  so each can be proposed upstream as a standalone PR (`gh pr create -R python-openxml/python-docx`).
- `main-bm` = integration branch merging the topic branches; this is what own projects
  install: `pip install git+https://github.com/BenoitMarchand/python-docx@main-bm`.
- Fork-only changes (like this file) go on `main-bm`, never in branches meant for upstream.

## Upstream context (GitHub issues, Nov 2026)

~386 open issues / ~164 open PRs; the project is maintained by a single maintainer (scanny)
and merges are slow and selective. The `shortlist` label marks issues the maintainer intends
to address. Recurring themes worth knowing before touching an area:

- **Known open bugs**: `_block_width` treats a zero page margin as 1 inch (#1638);
  `add_table()` with a missing style doesn't return the table (#1631); `Paragraph.text`
  setter drops comment/footnote references (#1604); `add_comment()` uses an undefined
  `CommentReference` style (#1609); `lazyproperty` re-evaluates when cached value is `None`
  (#1600); `run.add_picture` `'Part' object has no attribute 'sha1'` (#1565);
  unhandled exceptions on corrupt files (#1561); 3-digit RGB strings (#1466);
  non-integer font sizes (#1475); `cell._tc.bottom` ValueError (#1458).
- **Packaging debt**: trove classifiers still list Python 3.7/3.8 (#1497, #1620); the
  `.travis.yml` is obsolete (CI via tox); lxml minimum is very low (#1499); default template
  is Word 2010 compatibility mode (#1634).
- **Frequently requested, not supported**: bookmarks (#425, #1640), numbering/list
  introspection (#1454, #1540, #1541), table of contents refresh (#1403), content controls /
  SDT (#1417, #1489), floating images & wrapping (#1547), RTL/bidi (#1473, #1502),
  `.dotx`/`.docm` loading (#1532), Strict OOXML (#1520), charts (#1504), `pathlib.Path`
  input (#1544), restrict-editing (#1641).

When fixing one of these, cite the issue number in the commit message and in `HISTORY.rst`.
