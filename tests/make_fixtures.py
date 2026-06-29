#!/usr/bin/env python3
"""Build minimal real .docx fixtures for the docx extraction gate.

Crafts a valid Office Open XML package by hand (a hand-written word/document.xml
+ the required [Content_Types].xml) and zips it two ways:

  tests/fixtures/deflated.docx   — ZIP_DEFLATED  (exercises the inflate path)
  tests/fixtures/stored.docx     — ZIP_STORED    (exercises the verbatim path)

The document body has KNOWN text across multiple <w:p> paragraphs, an entity
(Tom &amp; Jerry -> "Tom & Jerry"), a <w:tab/>, and a <w:br/>. Both fixtures
share the same XML, so extraction must produce identical text regardless of the
ZIP method.
"""

import os
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
FIXTURES = os.path.join(HERE, "fixtures")

CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" '
    'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    "</Types>"
)

DOCUMENT_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    "<w:body>"
    '<w:p><w:r><w:t>Hello, Word!</w:t></w:r></w:p>'
    '<w:p><w:r><w:t>Tom &amp; Jerry</w:t></w:r></w:p>'
    '<w:p><w:r><w:t xml:space="preserve">First</w:t><w:tab/><w:t>Second</w:t>'
    '<w:br/><w:t>Third</w:t></w:r></w:p>'
    '<w:p><w:r><w:t>The &lt;end&gt;.</w:t></w:r></w:p>'
    "</w:body></w:document>"
)


# A FIXED zip timestamp so regeneration is byte-for-byte reproducible. `writestr`
# with a bare name stamps each entry with the CURRENT local time, which rewrote these
# committed fixtures on every test run (a perpetually-dirty submodule). 1980-01-01 is
# the ZIP epoch (the earliest a DOS date field can encode).
_EPOCH = (1980, 1, 1, 0, 0, 0)


def build(path, compression):
    with zipfile.ZipFile(path, "w") as z:
        for name, data in (
            ("[Content_Types].xml", CONTENT_TYPES),
            ("word/document.xml", DOCUMENT_XML),
        ):
            info = zipfile.ZipInfo(name, date_time=_EPOCH)
            info.compress_type = compression  # per-entry; ZipInfo ignores the ZipFile default
            z.writestr(info, data)


def main():
    os.makedirs(FIXTURES, exist_ok=True)
    build(os.path.join(FIXTURES, "deflated.docx"), zipfile.ZIP_DEFLATED)
    build(os.path.join(FIXTURES, "stored.docx"), zipfile.ZIP_STORED)
    print("wrote fixtures:", os.listdir(FIXTURES))


if __name__ == "__main__":
    main()
