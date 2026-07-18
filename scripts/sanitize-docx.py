#!/usr/bin/env python3
"""Anonymize a .docx for use as a parity fixture: every word in every text
run is replaced by a deterministic pseudoword with the same length and
capitalization shape, digits are remapped, document properties and comment
authors are scrubbed, and external hyperlink targets point at example.com.
Structure (styles, tables, SDTs, drawings, fields, breaks) is untouched, so
the layout exercises the same code paths; the Word reference PDF is
exported from the SANITIZED file, so line-break parity is preserved by
construction.

Usage: sanitize-docx.py <in.docx> <out.docx>
"""
import sys, re, zipfile, hashlib

VOWELS = "aeiou"
CONSONANTS = "bcdfghjklmnpqrstvwz"

# Small function words keep their rhythm; everything else scrambles.
KEEP = {
    "a", "an", "the", "of", "to", "in", "on", "at", "by", "or", "and", "for",
    "is", "are", "was", "be", "as", "with", "that", "this", "it", "its",
}


def pseudoword(word: str) -> str:
    if word.lower() in KEEP:
        return word
    seed = int.from_bytes(hashlib.sha1(("dxw:" + word).encode()).digest()[:8], "big")
    out = []
    for i, ch in enumerate(word):
        if ch.isdigit():
            out.append(str((int(ch) * 7 + 3 + i) % 10))
            continue
        if not ch.isalpha():
            out.append(ch)
            continue
        pool = VOWELS if i % 2 == 1 else CONSONANTS
        pick = pool[(seed >> (i * 3 % 48)) % len(pool)]
        out.append(pick.upper() if ch.isupper() else pick)
    return "".join(out)


def scramble_text(text: str) -> str:
    # The part text is raw XML: leave entities (&amp; &#8217; ...) intact.
    pieces = re.split(r"(&#?\w+;)", text)
    for i in range(0, len(pieces), 2):
        pieces[i] = re.sub(
            r"[A-Za-z0-9][A-Za-z0-9'’@.\-]*[A-Za-z0-9]|[A-Za-z0-9]",
            lambda m: pseudoword(m.group(0)),
            pieces[i],
        )
    return "".join(pieces)


def scramble_part(xml: str) -> str:
    def repl(m):
        return m.group(1) + scramble_text(m.group(2)) + m.group(3)

    for tag in ("w:t", "w:delText", "w:instrText"):
        if tag == "w:instrText":
            continue  # field instructions are structural
        xml = re.sub(rf"(<{tag}[^>]*>)([^<]+)(</{tag}>)", repl, xml)
    # comment/revision authors
    xml = re.sub(r'(w:author=")[^"]+(")', r"\1Reviewer\2", xml)
    # Field-instruction hyperlink targets (kept structural above, but the
    # URL itself is content and can carry real hosts).
    xml = re.sub(r'(HYPERLINK[^<]*?)https?://[^\s"&<]+', r"\1https://example.com/", xml)
    return xml


def scrub_core(xml: str) -> str:
    for tag in (
        "dc:creator", "cp:lastModifiedBy", "dc:title", "dc:subject",
        "dc:description", "cp:keywords", "cp:category", "cp:contentStatus",
    ):
        xml = re.sub(rf"(<{tag}[^>]*>)[^<]*(</{tag}>)", r"\g<1>Fixture\g<2>", xml)
    return xml


def scrub_rels(xml: str) -> str:
    # External hyperlinks -> example.com (keeps rel ids/structure).
    return re.sub(
        r'(Type="[^"]*hyperlink"[^>]*Target=")[^"]+(")',
        r"\1https://example.com/\2",
        xml,
    )


def scrub_app(xml: str) -> str:
    # docProps/app.xml: organization fields + the HLinks URL cache.
    for tag in ("Company", "Manager"):
        xml = re.sub(rf"(<{tag}>)[^<]*(</{tag}>)", r"\1Fixture\2", xml)
    xml = re.sub(r"(<vt:lpwstr>)https?://[^<]*(</vt:lpwstr>)", r"\1https://example.com/\2", xml)
    xml = re.sub(
        r"(<vt:lpw?str>)([^<]+)(</vt:lpw?str>)",
        lambda m: m.group(1) + scramble_text(m.group(2)) + m.group(3),
        xml,
    )
    return xml


def scrub_custom_xml(xml: str) -> str:
    # customXml data stores carry real values (e-signature labels, databound
    # SDT sources). Scramble text nodes same-length; redact content URLs in
    # attributes while preserving namespace/schema URIs.
    def text_repl(m):
        inner = m.group(1)
        return ">" + scramble_text(inner) + "<" if inner.strip() else m.group(0)

    xml = re.sub(r">([^<>]+)<", text_repl, xml)

    def attr_repl(m):
        name, val = m.group(1), m.group(2)
        if name.startswith("xmlns") or "schemas." in val or val.startswith("urn:"):
            return m.group(0)
        if val.startswith(("http://", "https://")):
            return f'{name}="https://example.com/"'
        return m.group(0)

    return re.sub(r'([A-Za-z_:][-A-Za-z0-9_:.]*)="([^"]*)"', attr_repl, xml)


def main() -> None:
    src, dst = sys.argv[1], sys.argv[2]
    zin = zipfile.ZipFile(src)
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            name = info.filename
            # NOTE: .rels files do NOT end with .xml - gating everything on
            # .xml silently skipped them (external hyperlink Targets shipped
            # unredacted until this was caught on the wild-* fixtures).
            if name.endswith(".xml") or name.endswith(".rels"):
                xml = data.decode("utf-8", "surrogateescape")
                if name.endswith(".rels"):
                    xml = scrub_rels(xml)
                elif name.startswith("word/") and re.search(r"/(document|header\d*|footer\d*|footnotes|endnotes|comments\w*)\.xml$", "/" + name):
                    xml = scramble_part(xml)
                elif name == "docProps/core.xml":
                    xml = scrub_core(xml)
                elif name == "docProps/app.xml":
                    xml = scrub_app(xml)
                elif name == "docProps/custom.xml":
                    xml = scrub_custom_xml(xml)
                elif name.startswith("customXml/"):
                    xml = scrub_custom_xml(xml)
                data = xml.encode("utf-8", "surrogateescape")
            zout.writestr(info, data)
    print("sanitized ->", dst)


if __name__ == "__main__":
    main()
