"""Docx extraction gate: builds REAL .docx fixtures (via tests/make_fixtures.py,
run by the `pixi run test` task) and asserts extraction over BOTH ZIP methods —
ZIP_DEFLATED (the inflate path) and ZIP_STORED (the verbatim path).

Asserts: the known paragraph text appears IN ORDER with paragraph breaks, the
`&amp;` entity decodes to `Tom & Jerry`, `<w:tab/>` becomes a tab, `<w:br/>`
becomes a newline, and `&lt;`/`&gt;` decode to `<`/`>`."""

from docx import (
    extract_text,
    extract_text_from_xml,
    read_file,
    parse_central_directory,
)


def _b(s: String) -> List[UInt8]:
    var out = List[UInt8]()
    var p = s.unsafe_ptr()
    for i in range(s.byte_length()):
        out.append(p[i])
    return out^


def _check(label: String, text: String) raises:
    # Paragraphs present, in order.
    var i_hello = text.find("Hello, Word!")
    var i_tom = text.find("Tom & Jerry")
    var i_first = text.find("First")
    var i_end = text.find("The <end>.")
    if i_hello == -1:
        raise Error(label + ": missing 'Hello, Word!' in [" + text + "]")
    if i_tom == -1:
        raise Error(
            label + ": entity not decoded to 'Tom & Jerry' in [" + text + "]"
        )
    if i_first == -1:
        raise Error(label + ": missing 'First' in [" + text + "]")
    if i_end == -1:
        raise Error(
            label + ": &lt;/&gt; not decoded to 'The <end>.' in [" + text + "]"
        )
    # Order: Hello < Tom < First < end.
    if not (i_hello < i_tom and i_tom < i_first and i_first < i_end):
        raise Error(label + ": paragraphs out of order in [" + text + "]")
    # Paragraph break between Hello and Tom.
    if text.find("Hello, Word!\nTom & Jerry") == -1:
        raise Error(
            label + ": no paragraph break between p1 and p2 in [" + text + "]"
        )
    # <w:tab/> -> tab between First and Second.
    if text.find("First\tSecond") == -1:
        raise Error(label + ": <w:tab/> not a tab in [" + text + "]")
    # <w:br/> -> newline between Second and Third.
    if text.find("Second\nThird") == -1:
        raise Error(label + ": <w:br/> not a newline in [" + text + "]")


def main() raises:
    # 0. Pure-XML unit (no ZIP) — the scanner in isolation.
    var xml = _b(
        '<w:document xmlns:w="x"><w:body>'
        "<w:p><w:r><w:t>Alpha</w:t></w:r></w:p>"
        "<w:p><w:r><w:t>A &amp; B</w:t></w:r></w:p>"
        "</w:body></w:document>"
    )
    var tx = extract_text_from_xml(xml)
    if (
        tx.find("Alpha") == -1
        or tx.find("A & B") == -1
        or tx.find("Alpha\nA & B") == -1
    ):
        raise Error("xml-scan unit failed: [" + tx + "]")

    # 1. DEFLATE path — a ZIP_DEFLATED .docx.
    var d_data = read_file("tests/fixtures/deflated.docx")
    var d_entries = parse_central_directory(d_data)
    var saw_deflate = False
    for i in range(len(d_entries)):
        if (
            d_entries[i].name == "word/document.xml"
            and d_entries[i].method == 8
        ):
            saw_deflate = True
    if not saw_deflate:
        raise Error(
            "deflated.docx: word/document.xml is not DEFLATE (method 8)"
        )
    var d_text = extract_text(d_data)
    _check("deflated.docx", d_text)

    # 2. STORED path — a ZIP_STORED .docx.
    var s_data = read_file("tests/fixtures/stored.docx")
    var s_entries = parse_central_directory(s_data)
    var saw_stored = False
    for i in range(len(s_entries)):
        if (
            s_entries[i].name == "word/document.xml"
            and s_entries[i].method == 0
        ):
            saw_stored = True
    if not saw_stored:
        raise Error("stored.docx: word/document.xml is not STORED (method 0)")
    var s_text = extract_text(s_data)
    _check("stored.docx", s_text)

    # 3. Both methods must yield identical text.
    if d_text != s_text:
        raise Error(
            "DEFLATE and STORED produced different text:\n["
            + d_text
            + "]\nvs\n["
            + s_text
            + "]"
        )

    print("docx extraction OK")
    print("  xml-scan -> 'Alpha' / 'A & B' with paragraph break")
    print("  DEFLATE  -> Hello/Tom & Jerry/First\\tSecond\\nThird/The <end>.")
    print("  STORED   -> identical text via the verbatim path")
