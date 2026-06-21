"""CLI: extract plain text from .docx files (the `docx` binary).

    docx <file.docx> [more.docx ...]   extract text and print it
    docx --info <file.docx>            list the ZIP members (no text)
    docx --help

Run it via `pixi run extract -- …` (resolves the zlib shim through CONDA_PREFIX).
Control characters in extracted text are escaped (\\xNN) so output never garbles
the terminal; if the zlib decoder can't load, it exits with an error rather than
emitting raw compressed bytes.
"""

from std.sys import argv
from docx import read_file, extract_text, parse_central_directory
from zlib import inflate


def _usage():
    print("usage: docx [--info] <file.docx> [more.docx ...]")
    print("  -i, --info   list the ZIP members instead of the text")
    print("  -h, --help   show this help")


def _zlib_ok() -> Bool:
    """Self-check: inflate a tiny known buffer (zlib of 'ok' -> 2 bytes). Fails if
    libzlibmojo.so can't be loaded — in which case DEFLATE entries can't decode."""
    var probe = List[Int]()
    probe.append(0x78); probe.append(0x9C); probe.append(0xCB); probe.append(0xCF)
    probe.append(0x06); probe.append(0x00); probe.append(0x01); probe.append(0x4B)
    probe.append(0x00); probe.append(0xDB)
    var b = List[UInt8]()
    for i in range(len(probe)):
        b.append(UInt8(probe[i]))
    try:
        return len(inflate(b)) == 2
    except:
        return False


def _hexdigit(n: Int) -> String:
    return chr(48 + n) if n < 10 else chr(87 + n)  # 0-9, a-f


def _escape_controls(s: String) -> String:
    """Escape control characters as \\xNN (keep tab/newline; pass UTF-8 through)
    so extracted text can never garble the terminal."""
    var out = String("")
    for cp in s.codepoints():
        var v = Int(cp)
        if v == 9 or v == 10:
            out += chr(v)
        elif v < 32 or v == 127:
            out += "\\x" + _hexdigit((v >> 4) & 0xF) + _hexdigit(v & 0xF)
        else:
            out += chr(v)
    return out^


def main() raises:
    var a = argv()
    var info = False
    var files = List[String]()
    for i in range(1, len(a)):
        var arg = String(a[i])
        if arg == "-h" or arg == "--help":
            _usage()
            return
        elif arg == "-i" or arg == "--info":
            info = True
        else:
            files.append(arg)

    if len(files) == 0:
        _usage()
        return

    if not info and not _zlib_ok():
        print("error: cannot load the zlib decoder (libzlibmojo.so) —")
        print("  ZIP DEFLATE entries can't be decompressed.")
        print("  Run via 'pixi run extract -- <file>' so CONDA_PREFIX is set.")
        return

    for fi in range(len(files)):
        var path = files[fi]
        if len(files) > 1:
            print("===== ", path, " =====", sep="")
        var data = read_file(path)
        if info:
            var entries = parse_central_directory(data)
            print("file:    ", path, sep="")
            print("  bytes:   ", len(data))
            print("  members: ", len(entries))
            for e in range(len(entries)):
                var m = "STORED" if entries[e].method == 0 else (
                    "DEFLATE" if entries[e].method == 8 else "method?")
                print("    ", entries[e].name, "  (", m, ", ",
                      entries[e].uncomp_size, " bytes)", sep="")
        else:
            print(_escape_controls(extract_text(data)))
