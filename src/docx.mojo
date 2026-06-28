"""Docx — extract plain text from a Microsoft Word .docx document (Office Open
XML), from scratch in Mojo.

Pipeline:
  1. The .docx is a ZIP archive. Parse it (End-Of-Central-Directory → central
     directory → local file headers) and extract the `word/document.xml` entry.
     Each ZIP entry is STORED (method 0, bytes verbatim) or DEFLATE (method 8,
     inflated via zlib.mojo — ZIP uses RAW deflate, no zlib header).
  2. Scan that XML for the run text: concatenate every `<w:t>…</w:t>`, emit a
     newline at each `</w:p>` (paragraph end), a tab for `<w:tab/>`, a newline
     for `<w:br/>` / `<w:cr/>`, and unescape XML entities. We also append the
     text of any `word/header*.xml` / `word/footer*.xml` parts.

Scope (v1): the modern `.docx` (OOXML) format only. Legacy binary `.doc`
(OLE2 / Compound File Binary) is NOT handled — that is a separate follow-up
(see README). No styles/numbering rendering, no field evaluation, no images.

Mirrors pdftotext.mojo/src/pdf.mojo: same `read_file` / `extract_text` public
shape, the `Buf` byte-buffer helper, and the same byte-scan style — so wiring
this into vault/core/src/readers.mojo next to `pdf_text` is trivial.
"""

from zlib import inflate


# ── byte helpers ─────────────────────────────────────────────────────────────


def _ascii(s: String) -> List[UInt8]:
    var out = List[UInt8]()
    var p = s.unsafe_ptr()
    for i in range(s.byte_length()):
        out.append(p[i])
    return out^


def _find(data: List[UInt8], pat: List[UInt8], start: Int) -> Int:
    """Index of the first occurrence of `pat` in `data` at/after `start`, or -1.
    """
    var n = len(data)
    var m = len(pat)
    if m == 0 or m > n:
        return -1
    var i = start
    while i <= n - m:
        var ok = True
        for j in range(m):
            if data[i + j] != pat[j]:
                ok = False
                break
        if ok:
            return i
        i += 1
    return -1


def _eq_at(data: List[UInt8], pat: List[UInt8], at: Int) -> Bool:
    """True if `pat` matches `data` exactly at index `at`."""
    var m = len(pat)
    if at < 0 or at + m > len(data):
        return False
    for j in range(m):
        if data[at + j] != pat[j]:
            return False
    return True


def _u16le(data: List[UInt8], at: Int) -> Int:
    return Int(data[at]) | (Int(data[at + 1]) << 8)


def _u32le(data: List[UInt8], at: Int) -> Int:
    return (
        Int(data[at])
        | (Int(data[at + 1]) << 8)
        | (Int(data[at + 2]) << 16)
        | (Int(data[at + 3]) << 24)
    )


struct Buf(Movable):
    """An output buffer over `List[UInt8]` so `+=` is amortized O(1). Mojo's
    `String +=` reallocates per append — O(n^2) on the per-run text hot path.
    (Same helper as pdftotext.mojo/src/pdf.mojo.)"""

    var data: List[UInt8]

    def __init__(out self):
        self.data = List[UInt8]()

    def __iadd__(mut self, s: String):
        var b = s.as_bytes()
        for i in range(len(b)):
            self.data.append(b[i])

    def append_byte(mut self, b: UInt8):
        self.data.append(b)

    def append_bytes(mut self, b: List[UInt8]):
        for i in range(len(b)):
            self.data.append(b[i])

    def last_byte(self) -> Int:
        """Last byte, or -1 if empty (trailing newline dedupe)."""
        var n = len(self.data)
        if n == 0:
            return -1
        return Int(self.data[n - 1])

    def to_string(self) -> String:
        return String(unsafe_from_utf8=Span(self.data))


# ── ZIP container ────────────────────────────────────────────────────────────


struct ZipEntry(Copyable, Movable):
    """One central-directory record we care about: name + where/how its data is.
    """

    var name: String
    var method: Int  # 0 = STORED, 8 = DEFLATE
    var comp_size: Int
    var uncomp_size: Int
    var local_header_off: Int

    def __init__(out self):
        self.name = String("")
        self.method = 0
        self.comp_size = 0
        self.uncomp_size = 0
        self.local_header_off = 0


def _read_name(data: List[UInt8], at: Int, length: Int) -> String:
    var b = List[UInt8]()
    for i in range(length):
        b.append(data[at + i])
    return String(unsafe_from_utf8=Span(b))


def parse_central_directory(data: List[UInt8]) raises -> List[ZipEntry]:
    """Parse the ZIP central directory into entries. Locates the
    End-Of-Central-Directory record (sig 0x06054b50) by scanning backwards,
    then walks the central-directory file headers (sig 0x02014b50)."""
    var n = len(data)
    var entries = List[ZipEntry]()
    if n < 22:
        return entries^

    # EOCD: scan backwards for its signature (it's within the last 64KB + 22).
    var eocd_sig = List[UInt8]()
    eocd_sig.append(0x50)
    eocd_sig.append(0x4B)
    eocd_sig.append(0x05)
    eocd_sig.append(0x06)
    var eocd = -1
    var i = n - 22
    var floor = n - 22 - 65536
    if floor < 0:
        floor = 0
    while i >= floor:
        if _eq_at(data, eocd_sig, i):
            eocd = i
            break
        i -= 1
    if eocd < 0:
        raise Error("docx: not a ZIP (no End-Of-Central-Directory record)")

    var total = _u16le(data, eocd + 10)  # total CD entries
    var cd_off = _u32le(data, eocd + 16)  # CD start offset

    var cd_sig = List[UInt8]()
    cd_sig.append(0x50)
    cd_sig.append(0x4B)
    cd_sig.append(0x01)
    cd_sig.append(0x02)

    var p = cd_off
    var count = 0
    while count < total and p + 46 <= n:
        if not _eq_at(data, cd_sig, p):
            break
        var method = _u16le(data, p + 10)
        var comp_size = _u32le(data, p + 20)
        var uncomp_size = _u32le(data, p + 24)
        var name_len = _u16le(data, p + 28)
        var extra_len = _u16le(data, p + 30)
        var comment_len = _u16le(data, p + 32)
        var local_off = _u32le(data, p + 42)
        var e = ZipEntry()
        e.name = _read_name(data, p + 46, name_len)
        e.method = method
        e.comp_size = comp_size
        e.uncomp_size = uncomp_size
        e.local_header_off = local_off
        entries.append(e^)
        p += 46 + name_len + extra_len + comment_len
        count += 1
    return entries^


def extract_entry(data: List[UInt8], e: ZipEntry) raises -> List[UInt8]:
    """Decompress one ZIP entry. The central-directory record points at the
    local file header (sig 0x04034b50); the data follows the header + its own
    name/extra fields (these can differ in length from the CD copy). STORED
    (method 0) is copied verbatim; DEFLATE (method 8) is raw-inflated."""
    var n = len(data)
    var lh = e.local_header_off
    var lh_sig = List[UInt8]()
    lh_sig.append(0x50)
    lh_sig.append(0x4B)
    lh_sig.append(0x03)
    lh_sig.append(0x04)
    if not _eq_at(data, lh_sig, lh):
        raise Error("docx: bad local file header for " + e.name)
    var name_len = _u16le(data, lh + 26)
    var extra_len = _u16le(data, lh + 28)
    var ds = lh + 30 + name_len + extra_len
    var de = ds + e.comp_size
    if de > n:
        de = n
    var raw = List[UInt8]()
    for i in range(ds, de):
        raw.append(data[i])

    if e.method == 0:  # STORED
        return raw^
    elif e.method == 8:  # DEFLATE — ZIP uses RAW deflate (no zlib
        return inflate(raw)  # header); inflate() auto-detects raw.
    else:
        raise Error(
            "docx: unsupported ZIP compression method " + String(e.method)
        )


def read_zip_member(data: List[UInt8], name: String) raises -> List[UInt8]:
    """Decompressed bytes of the named ZIP member, or empty if absent."""
    var entries = parse_central_directory(data)
    for i in range(len(entries)):
        if entries[i].name == name:
            return extract_entry(data, entries[i])
    return List[UInt8]()


# ── XML text extraction ──────────────────────────────────────────────────────


def _unescape_into(mut out: Buf, ent: List[UInt8]):
    """Append the decoded form of an XML entity body (the text BETWEEN `&` and
    `;`, exclusive) to `out`. Handles the 5 predefined entities plus numeric
    `&#NN;` / `&#xHH;`. Unknown entities are passed through verbatim."""
    var m = len(ent)
    if m == 0:
        return
    if _list_eq_str(ent, "amp"):
        out.append_byte(38)  # &
    elif _list_eq_str(ent, "lt"):
        out.append_byte(60)  # <
    elif _list_eq_str(ent, "gt"):
        out.append_byte(62)  # >
    elif _list_eq_str(ent, "quot"):
        out.append_byte(34)  # "
    elif _list_eq_str(ent, "apos"):
        out.append_byte(39)  # '
    elif ent[0] == 35:  # '#'  numeric char reference
        var cp = 0
        var ok = True
        if m >= 2 and (ent[1] == 120 or ent[1] == 88):  # x / X  hex
            for k in range(2, m):
                var hv = _hexval(ent[k])
                if hv < 0:
                    ok = False
                    break
                cp = cp * 16 + hv
        else:
            for k in range(1, m):
                if ent[k] < 48 or ent[k] > 57:
                    ok = False
                    break
                cp = cp * 10 + (Int(ent[k]) - 48)
        if ok:
            _append_utf8(out, cp)
        else:
            _passthrough_entity(out, ent)
    else:
        _passthrough_entity(out, ent)


def _passthrough_entity(mut out: Buf, ent: List[UInt8]):
    out.append_byte(38)  # &
    out.append_bytes(ent)
    out.append_byte(59)  # ;


def _list_eq_str(b: List[UInt8], s: String) -> Bool:
    var sp = s.unsafe_ptr()
    var m = s.byte_length()
    if len(b) != m:
        return False
    for i in range(m):
        if b[i] != sp[i]:
            return False
    return True


def _hexval(c: UInt8) -> Int:
    if c >= 48 and c <= 57:
        return Int(c) - 48
    if c >= 65 and c <= 70:
        return Int(c) - 55
    if c >= 97 and c <= 102:
        return Int(c) - 87
    return -1


def _append_utf8(mut out: Buf, cp: Int):
    if cp < 0x80:
        out.append_byte(UInt8(cp))
    elif cp < 0x800:
        out.append_byte(UInt8(0xC0 | (cp >> 6)))
        out.append_byte(UInt8(0x80 | (cp & 0x3F)))
    elif cp < 0x10000:
        out.append_byte(UInt8(0xE0 | (cp >> 12)))
        out.append_byte(UInt8(0x80 | ((cp >> 6) & 0x3F)))
        out.append_byte(UInt8(0x80 | (cp & 0x3F)))
    else:
        out.append_byte(UInt8(0xF0 | (cp >> 18)))
        out.append_byte(UInt8(0x80 | ((cp >> 12) & 0x3F)))
        out.append_byte(UInt8(0x80 | ((cp >> 6) & 0x3F)))
        out.append_byte(UInt8(0x80 | (cp & 0x3F)))


def _local_name_is(data: List[UInt8], tag_start: Int, name: String) -> Bool:
    """True if the element whose `<` is at `tag_start` has local-name `name`,
    i.e. `<name`, `<name>`, `<name/`, `<name ` or `<prefix:name…` (the `w:`
    namespace prefix). Matches the local part after any single `prefix:`."""
    var n = len(data)
    var p = tag_start + 1  # past '<'
    # skip an optional namespace prefix "xxx:"
    var q = p
    var colon = -1
    while q < n:
        var c = data[q]
        if c == 58:  # ':'
            colon = q
            break
        if c == 62 or c == 47 or c == 32 or c == 9 or c == 10 or c == 13:
            break
        q += 1
    if colon >= 0:
        p = colon + 1
    var sp = name.unsafe_ptr()
    var m = name.byte_length()
    if p + m > n:
        return False
    for i in range(m):
        if data[p + i] != sp[i]:
            return False
    var after = data[p + m]  # delimiter after the local name
    return (
        after == 62
        or after == 47
        or after == 32
        or after == 9
        or after == 10
        or after == 13
    )


def extract_text_from_xml(xml: List[UInt8]) raises -> String:
    """Pull document text from one OOXML part (`word/document.xml`, header, or
    footer). Concatenate `<w:t>…</w:t>` run text; `</w:p>` → newline; `<w:tab/>`
    → tab; `<w:br/>` / `<w:cr/>` → newline. XML entities are unescaped. A focused
    tag scanner (no full XML parser), in the spirit of pdf.mojo.

    `<w:t>` may carry `xml:space="preserve"` and arbitrary attributes, so we
    copy the raw text from `>` to the next `<` verbatim (whitespace preserved).
    """
    var out = Buf()
    var n = len(xml)
    var i = 0
    var in_t = False  # inside a <w:t> … </w:t> text run
    while i < n:
        var c = xml[i]
        if c == 60:  # '<' — a tag starts
            in_t = False
            if i + 1 < n and xml[i + 1] == 47:  # '</' closing tag
                if _local_name_is(xml, i + 1, "p"):  # </w:p>
                    if out.last_byte() != 10:
                        out.append_byte(10)
                # </w:t> just ends the run (in_t already cleared)
            else:  # opening / empty tag
                if _local_name_is(xml, i, "t"):
                    # Enter text mode only for a non-self-closing <w:t …>.
                    var gt = _find(xml, _ascii(">"), i)
                    if gt != -1 and not (gt > 0 and xml[gt - 1] == 47):
                        in_t = True
                elif _local_name_is(xml, i, "tab"):
                    out.append_byte(9)  # tab
                elif _local_name_is(xml, i, "br") or _local_name_is(
                    xml, i, "cr"
                ):
                    out.append_byte(10)  # line break
            # Skip to the end of this tag.
            var gt = _find(xml, _ascii(">"), i)
            if gt == -1:
                break
            i = gt + 1
            continue
        elif c == 38 and in_t:  # '&' entity inside text
            var semi = _find(xml, _ascii(";"), i)
            if semi != -1 and semi - i <= 12:
                var ent = List[UInt8]()
                for k in range(i + 1, semi):
                    ent.append(xml[k])
                _unescape_into(out, ent)
                i = semi + 1
                continue
            else:
                out.append_byte(c)  # stray '&'
                i += 1
                continue
        else:
            if in_t:
                out.append_byte(c)
            i += 1
    return out.to_string()


# ── public API (mirrors pdftotext.mojo/src/pdf.mojo) ─────────────────────────


def extract_text(data: List[UInt8]) raises -> String:
    """Top level: parse the .docx ZIP, extract `word/document.xml` (required),
    then append any `word/header*.xml` / `word/footer*.xml` parts, and pull the
    text from each. Returns the document's plain text."""
    var entries = parse_central_directory(data)

    # The main story.
    var body = List[UInt8]()
    var have_body = False
    for i in range(len(entries)):
        if entries[i].name == "word/document.xml":
            body = extract_entry(data, entries[i])
            have_body = True
            break
    if not have_body:
        raise Error("docx: no word/document.xml (not a Word document?)")

    var out = Buf()
    out += extract_text_from_xml(body)

    # Stretch: headers/footers (word/header1.xml, word/footer2.xml, …). Append
    # in name order so output is deterministic.
    var hf_idx = _headers_footers_sorted(entries)
    for k in range(len(hf_idx)):
        var part = extract_entry(data, entries[hf_idx[k]])
        var txt = extract_text_from_xml(part)
        if txt.byte_length() > 0:
            if out.last_byte() != 10 and out.last_byte() != -1:
                out.append_byte(10)
            out += txt
    return out.to_string()


def _headers_footers_sorted(entries: List[ZipEntry]) -> List[Int]:
    """Indices of word/header*.xml + word/footer*.xml entries, sorted by name.
    """
    var idx = List[Int]()
    for i in range(len(entries)):
        var nm = entries[i].name
        if (
            nm.startswith("word/header") or nm.startswith("word/footer")
        ) and nm.endswith(".xml"):
            idx.append(i)
    # selection sort by name (small N)
    for a in range(len(idx)):
        var mi = a
        for b in range(a + 1, len(idx)):
            if entries[idx[b]].name < entries[idx[mi]].name:
                mi = b
        var t = idx[a]
        idx[a] = idx[mi]
        idx[mi] = t
    return idx^


def read_file(path: String) raises -> List[UInt8]:
    """Read a file as raw bytes."""
    var out = List[UInt8]()
    with open(path, "r") as f:
        var b = f.read_bytes()
        for i in range(len(b)):
            out.append(b[i])
    return out^
