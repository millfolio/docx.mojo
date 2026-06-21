# word.mojo

A from-scratch **.docx → plain text** extractor in Mojo. Sibling to
`pdftotext.mojo`; mirrors its layout and conventions so it can slot into
`vault/core/src/readers.mojo` next to `pdf_text`.

## The format

A `.docx` is an **Office Open XML (OOXML)** package: a ZIP archive whose main
text lives in `word/document.xml`. Extraction is two stages:

1. **ZIP container** — locate and decompress `word/document.xml`. We read the
   End-Of-Central-Directory record (sig `0x06054b50`), walk the central
   directory (sig `0x02014b50`), then follow each entry to its local file header
   (sig `0x04034b50`). Each entry is **STORED** (method 0 — bytes verbatim) or
   **DEFLATE** (method 8). DEFLATE is inflated via the sibling
   [`zlib.mojo`](../zlib.mojo) shim — ZIP uses **raw** deflate (no zlib header),
   which `zlib`'s `inflate()` auto-detects (it tries the zlib wrapper, then raw).
2. **XML text scan** — a focused scanner over `word/document.xml` (no full XML
   parser, in the spirit of `pdf.mojo`):
   - concatenate the text inside every `<w:t>…</w:t>` run,
   - newline at each paragraph end `</w:p>`,
   - `<w:tab/>` → tab, `<w:br/>` / `<w:cr/>` → newline,
   - unescape XML entities (`&amp; &lt; &gt; &quot; &apos;`, plus numeric
     `&#NN;` / `&#xHH;`).

   It also appends the text of any `word/header*.xml` / `word/footer*.xml`
   parts. The `w:` namespace prefix is matched by local name, so unprefixed or
   alternately-prefixed documents still parse.

## Public API (`src/docx.mojo`)

Mirrors `pdftotext.mojo/src/pdf.mojo` exactly, so wiring into `readers.mojo` is
trivial:

```mojo
def read_file(path: String) raises -> List[UInt8]          # raw .docx bytes
def extract_text(data: List[UInt8]) raises -> String       # the document's text
```

Supporting (also exported): `parse_central_directory`, `extract_entry`,
`read_zip_member`, `extract_text_from_xml`, the `ZipEntry` struct, and the `Buf`
byte-buffer helper (amortized-O(1) output append; `String +=` in a loop is
O(n²)).

## Build / test

Built with the unified Mojo toolchain via `pixi`, like every Mojo project here:

```bash
pixi run build     # builds the zlib shim + the `word` CLI → build/word
pixi run test      # builds REAL .docx fixtures (python3 + stdlib zipfile),
                   # then asserts extraction over BOTH ZIP_DEFLATED and
                   # ZIP_STORED, including an &amp; → "Tom & Jerry" entity
pixi run extract -- file.docx     # extract + print (sets CONDA_PREFIX)
```

`pixi run test` runs `tests/make_fixtures.py`, which hand-writes a minimal valid
`.docx` (a `word/document.xml` + `[Content_Types].xml`) and zips it twice — once
`ZIP_DEFLATED`, once `ZIP_STORED` — proving both decompression paths.

## CLI

```bash
tools/word file.docx           # extract text (sets CONDA_PREFIX, builds on first use)
tools/word --info file.docx    # list ZIP members + their compression method
```

The compiled binary `dlopen`s `libzlibmojo.so` via `$CONDA_PREFIX/lib`. Run it
through `pixi run extract` or the `tools/word` wrapper so that resolves; a bare
`./build/word` without `CONDA_PREFIX` exits with an error rather than emitting
raw compressed bytes.

## Scope / not yet handled

- **Legacy binary `.doc`** (OLE2 / Compound File Binary, Word 97–2003) is **out
  of scope** — a separate format (a structured-storage container, not ZIP+XML).
  Future follow-up: detect the OLE2 magic (`D0 CF 11 E0 A1 B1 1A E1`) and parse
  the `WordDocument` stream's piece table, or shell out to a converter.
- ZIP64 (archives ≥ 4 GB or ≥ 65 535 entries) — the 32-bit EOCD fields are used;
  a real Word doc never hits this, but a ZIP64 EOCD locator is not parsed.
- Tables render as their cell text run together (each cell paragraph still
  yields its `</w:p>` newline); column structure is not reconstructed.
- Numbered/bulleted list markers are not synthesized (the list text is kept; the
  "1." / "•" prefix lives in numbering.xml and is not rendered).
- Footnotes/endnotes (`word/footnotes.xml`, `endnotes.xml`), comments, and
  textbox/drawing text are not appended.
- Fields (e.g. page numbers, TOC) are not evaluated; only their cached result
  text inside `<w:t>` is captured.
```
