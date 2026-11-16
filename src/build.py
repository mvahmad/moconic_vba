"""Build the paste-ready Mokonic.bas from the readable source.

Every VBA string literal that contains non-ASCII characters is rewritten as
T("...") where "~XX" stands for U+06XX and "^XXXX" for any other code point.
The output is pure ASCII with CRLF line endings, so it survives being pasted
into the VBA editor on any Windows locale.
"""
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
SRC = HERE / "Mokonic_readable.bas"
OUT = HERE.parent / "Mokonic.bas"
MAX_LINE = 1000  # VBA hard limit is 1023


def encode_literal(content: str) -> str:
    """content = the string value (quotes already un-doubled)."""
    out = []
    for ch in content:
        o = ord(ch)
        if ch == '"':
            out.append('""')
        elif ch in "~^":
            out.append("^%04X" % o)
        elif o < 128:
            out.append(ch)
        elif 0x600 <= o <= 0x6FF:
            out.append("~%02X" % (o - 0x600))
        elif o <= 0xFFFF:
            out.append("^%04X" % o)
        else:
            raise ValueError("character outside BMP: %r" % ch)
    return 'T("' + "".join(out) + '")'


def convert_line(line: str, lineno: int) -> str:
    res = []
    i, n = 0, len(line)
    while i < n:
        ch = line[i]
        if ch == '"':
            j = i + 1
            buf = []
            while True:
                if j >= n:
                    raise ValueError("line %d: unterminated string" % lineno)
                if line[j] == '"':
                    if j + 1 < n and line[j + 1] == '"':
                        buf.append('"')
                        j += 2
                        continue
                    break
                buf.append(line[j])
                j += 1
            content = "".join(buf)
            if any(ord(c) > 127 for c in content):
                res.append(encode_literal(content))
            else:
                res.append(line[i:j + 1])
            i = j + 1
        elif ch == "'":
            rest = line[i:]
            if any(ord(c) > 127 for c in rest):
                raise ValueError("line %d: non-ASCII text in a comment" % lineno)
            res.append(rest)
            break
        else:
            if ord(ch) > 127:
                raise ValueError("line %d: non-ASCII character outside a string" % lineno)
            res.append(ch)
            i += 1
    return "".join(res)


def main() -> int:
    text = SRC.read_text(encoding="utf-8").replace("\r\n", "\n")
    out_lines = []
    errors = []
    for k, line in enumerate(text.split("\n"), 1):
        try:
            conv = convert_line(line, k)
        except ValueError as e:
            errors.append(str(e))
            continue
        stripped = conv.lstrip().lower()
        if "t(\"" in conv and (stripped.startswith("const ") or stripped.startswith("private const ")
                               or stripped.startswith("public const ")):
            errors.append("line %d: Persian text in a Const (not allowed)" % k)
        if len(conv) > MAX_LINE:
            errors.append("line %d: %d chars after encoding (max %d)" % (k, len(conv), MAX_LINE))
        out_lines.append(conv)
    if errors:
        print("\n".join(errors))
        return 1
    data = "\r\n".join(out_lines).rstrip("\r\n") + "\r\n"
    data.encode("ascii")  # guarantees pure ASCII
    OUT.write_bytes(data.encode("ascii"))
    print("OK: %s (%d lines, %d bytes, longest line %d)" % (
        OUT, len(out_lines), len(data), max(len(l) for l in out_lines)))
    return 0


if __name__ == "__main__":
    sys.exit(main())
