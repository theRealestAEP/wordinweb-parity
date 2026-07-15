#!/usr/bin/env python3
"""
audit-fixtures.py — scan WordInWeb test fixtures for real / identifying / sensitive
information before they are published to the public internet.

Covers, per .docx (an OOXML zip):
  1. Package metadata     docProps/core.xml, app.xml, custom.xml
  2. Author trails        w:author / w:initials on comments, tracked changes (ins/del),
                          word/people.xml, w:lastModifiedBy, moveFrom/moveTo authors
  3. Missed content       headers/footers, foot/endnotes, textboxes (w:txbxContent),
                          field instructions (instrText — HYPERLINK urls), bookmark
                          names, drawing alt-text (wp:docPr @name/@descr)
  4. Hyperlink targets    word/_rels/*.rels external targets (real URLs / mailto)
  5. Embedded media       word/media/* listed with sha256 + dimensions for manual review
  6. Fonts / OLE / macros word/fonts/*.odttf, embeddings/*.bin, vbaProject.bin
  7. Word owner-lock      ~$*.docx files (embed last opener's username in cleartext)
  8. Reference PDFs       pdfinfo Author/Creator/Producer metadata (needs pdfinfo)

Output: a machine-readable JSON (--json) and a Markdown table (--md), plus a console
summary. Metadata/author/hyperlink/alt-text checks are fully automated and severity-
tagged; media is listed with hashes for a human to eyeball.

Usage:
  python3 scripts/audit-fixtures.py \
      --root /Users/alexpickett/Desktop/Projects/wordinweb \
      --md docs/FIXTURE-AUDIT.md --json /tmp/audit.json

This script only READS. It never modifies a fixture or PDF.
"""
import argparse, hashlib, io, json, os, re, subprocess, sys, zipfile
import xml.etree.ElementTree as ET

# ---- OOXML namespaces -------------------------------------------------------
NS = {
    'w':   'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'cp':  'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
    'dc':  'http://purl.org/dc/elements/1.1/',
    'dcterms': 'http://purl.org/dc/terms/',
    'ep':  'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
    'vt':  'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes',
    'wp':  'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'rel': 'http://schemas.openxmlformats.org/package/2006/relationships',
}

# Values that are benign defaults — leave as OK (not even worth scrubbing).
BENIGN_AUTHORS = {
    '', 'Microsoft Office User', 'Unknown', 'user', 'User',
    'Windows User', 'Word', 'LibreOffice',
}
# Placeholder / tooling / synthetic names: NOT real PII, but worth scrubbing => WARN.
# (matched case-insensitively, substring on the role words)
PLACEHOLDER_NAMES = {
    'un-named', 'unnamed', 'fixture', 'python-docx', 'admin', 'administrator',
    'microsoft', 'author', 'test', 'testuser', 'pscript5.dll',
}
ROLE_WORDS = ('reviewer', 'editor', 'manager', 'admin', 'author', 'user',
              'fixture', 'test')
# Regexes for scanning free text channels.
RE_EMAIL = re.compile(r'[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}')
RE_URL   = re.compile(r'https?://[^\s"<>]+', re.I)
RE_REALNAME = re.compile(r'\b[A-Z][a-z]+ [A-Z][a-z]+\b')  # "Firstname Lastname"
# Absolute local paths that embed the machine username (attachedTemplate rels).
RE_FILE_USER_PATH = re.compile(r'file:///+Users/[^"<\s]+', re.IGNORECASE)


def classify_person(val):
    """Severity for a name/author-ish value.
    BLOCKER = looks like a real person / real email; WARN = placeholder/synthetic;
    OK = empty/well-known default."""
    v = (val or '').strip()
    if not v or v in BENIGN_AUTHORS:
        return 'OK'
    if RE_EMAIL.search(v):
        return 'BLOCKER'
    low = v.lower()
    if low in PLACEHOLDER_NAMES:
        return 'WARN'
    if any(w in low for w in ROLE_WORDS):   # "Reviewer A", "Bob Editor", "Manager"
        return 'WARN'
    if RE_REALNAME.search(v):               # "Alex Pickett", "Nelson, Jeffrey"
        return 'BLOCKER'
    return 'WARN'                           # single pseudoword like "Cobbery"


def sev_rank(s):
    return {'BLOCKER': 0, 'WARN': 1, 'REVIEW': 2, 'OK': 3}.get(s, 4)


class Finding:
    __slots__ = ('fixture', 'channel', 'detail', 'severity')
    def __init__(self, fixture, channel, detail, severity):
        self.fixture, self.channel, self.detail, self.severity = fixture, channel, detail, severity
    def row(self):
        return [self.fixture, self.channel, self.detail, self.severity]


def read_xml(z, name):
    try:
        return ET.fromstring(z.read(name))
    except KeyError:
        return None
    except ET.ParseError:
        return None


def text_of(el):
    return ''.join(el.itertext()) if el is not None else ''


def q(tag):
    """Expand a w:foo prefixed tag into Clark notation for findall."""
    pfx, local = tag.split(':')
    return f'{{{NS[pfx]}}}{local}'


# ---- individual channel scanners -------------------------------------------

def scan_core(z, fx, out):
    root = read_xml(z, 'docProps/core.xml')
    if root is None:
        return
    checks = [
        ('dc:creator',        'core.xml creator'),
        ('cp:lastModifiedBy', 'core.xml lastModifiedBy'),
        ('dc:title',          'core.xml title'),
        ('dc:subject',        'core.xml subject'),
        ('cp:keywords',       'core.xml keywords'),
        ('dc:description',    'core.xml description'),
        ('cp:category',       'core.xml category'),
        ('cp:contentStatus',  'core.xml contentStatus'),
    ]
    for path, label in checks:
        el = root.find(path, NS)
        val = (el.text or '').strip() if el is not None else ''
        if not val:
            continue
        if path in ('dc:creator', 'cp:lastModifiedBy'):
            sev = classify_person(val)
        else:
            # title/subject/keywords etc: BLOCKER if it carries an email or a real
            # personal name; otherwise WARN (metadata worth scrubbing).
            if RE_EMAIL.search(val) or (RE_REALNAME.search(val)
                                        and val.lower() not in PLACEHOLDER_NAMES):
                sev = 'BLOCKER'
            else:
                sev = 'WARN'
        out.append(Finding(fx, label, repr(val), sev))


def scan_app(z, fx, out):
    root = read_xml(z, 'docProps/app.xml')
    if root is None:
        return
    for path, label in [('ep:Company', 'app.xml Company'),
                        ('ep:Manager', 'app.xml Manager')]:
        el = root.find(path, NS)
        val = (el.text or '').strip() if el is not None else ''
        if val:
            out.append(Finding(fx, label, repr(val), classify_person(val)))
    # TitlesOfParts / HeadingPairs can leak the original doc title text
    tp = root.find('ep:TitlesOfParts', NS)
    if tp is not None:
        vals = [t.text for t in tp.iter() if t.text and t.text.strip()]
        vals = [v.strip() for v in vals if v.strip()]
        if vals:
            out.append(Finding(fx, 'app.xml TitlesOfParts',
                               repr(vals[:6]), 'WARN'))


def scan_custom(z, fx, out):
    try:
        raw = z.read('docProps/custom.xml')
    except KeyError:
        return
    root = read_xml(z, 'docProps/custom.xml')
    props = []
    if root is not None:
        for p in root.iter():
            if p.tag.endswith('}property'):
                name = p.get('name', '')
                val = text_of(p).strip()
                if name or val:
                    props.append(f'{name}={val!r}')
    if props:
        emaily = any(RE_EMAIL.search(s) for s in props)
        out.append(Finding(fx, 'custom.xml properties', '; '.join(props[:8]),
                           'BLOCKER' if emaily else 'WARN'))


def scan_authors(z, fx, out):
    """Author names on comments, tracked changes, people.xml across all parts."""
    authors = {}   # author -> set(channels)
    initials = set()
    # people.xml
    ppl = read_xml(z, 'word/people.xml')
    if ppl is not None:
        for person in ppl.iter(q('w:person')):
            a = person.get(q('w:author')) or person.get('author')
            if a:
                authors.setdefault(a, set()).add('people.xml')
        # presenceInfo / userId can hold AD / email identifiers
        for pi in ppl.iter():
            if pi.tag.endswith('}presenceInfo'):
                uid = pi.get(q('w:userId')) or pi.get('userId')
                if uid:
                    out.append(Finding(fx, 'people.xml presenceInfo',
                                       repr(uid), 'BLOCKER'))
    # every part: attributes w:author, w:initials on ins/del/comment/move*
    for name in z.namelist():
        if not name.endswith('.xml'):
            continue
        root = read_xml(z, name)
        if root is None:
            continue
        for el in root.iter():
            a = el.get(q('w:author'))
            if a:
                authors.setdefault(a, set()).add(name.split('/')[-1])
            ini = el.get(q('w:initials'))
            if ini:
                initials.add(ini)
    for a, chans in sorted(authors.items()):
        out.append(Finding(fx, 'author attribute',
                           f'{a!r} in {sorted(chans)}', classify_person(a)))
    real_ini = {i for i in initials if i and i.strip() and i not in ('', 'A', 'MOU')}
    if real_ini:
        out.append(Finding(fx, 'author initials', repr(sorted(real_ini)), 'WARN'))


def scan_owner_leaks(z, fx, out):
    """Two leak classes the 2026-07-14 sweep found the hard way:
    1. The repo owner's name/handle in BODY TEXT (pickett.docx carried
       'Alex Pickett' in w:t runs — the metadata scanners never look there).
    2. file:///Users/<name>/… paths anywhere (settings.xml.rels embeds the
       attachedTemplate as an absolute path containing the macOS username).
    """
    OWNER_TOKENS = ('pickett', 'ultimatetournament')
    for name in z.namelist():
        if not name.endswith(('.xml', '.rels')):
            continue
        try:
            s = z.read(name).decode('utf8', errors='ignore')
        except Exception:
            continue
        low = s.lower()
        for tok in OWNER_TOKENS:
            if tok in low:
                out.append(Finding(fx, 'owner token', f'{tok!r} in {name}', 'BLOCKER'))
        for m in RE_FILE_USER_PATH.finditer(s):
            out.append(Finding(fx, 'local user path', f'{m.group(0)[:80]!r} in {name}', 'BLOCKER'))


def scan_hyperlinks(z, fx, out):
    """External relationship targets (real URLs / mailto) in *.rels."""
    for name in z.namelist():
        if not name.endswith('.rels'):
            continue
        root = read_xml(z, name)
        if root is None:
            continue
        for r in root:
            if not r.tag.endswith('}Relationship'):
                continue
            mode = r.get('TargetMode', '')
            tgt = r.get('Target', '')
            typ = r.get('Type', '')
            if mode == 'External' and (tgt.startswith('http') or tgt.startswith('mailto')):
                # hyperlink/image/oleObject external targets
                low = tgt.lower()
                benign = ('schemas.openxmlformats.org', 'schemas.microsoft.com',
                          'w3.org', 'purl.org')
                if any(b in low for b in benign):
                    continue
                sev = 'BLOCKER' if ('mailto:' in low and 'example.com' not in low) else 'WARN'
                short = tgt if len(tgt) < 120 else tgt[:117] + '...'
                out.append(Finding(fx, 'external rel target',
                                   f'{short} ({name.split("/")[-1]})', sev))


def scan_field_instructions(z, fx, out):
    """instrText fields: HYPERLINK urls, AUTHOR, filename paths."""
    for name in z.namelist():
        if not (name.endswith('.xml') and ('document' in name or 'header' in name
                or 'footer' in name or 'footnote' in name or 'endnote' in name)):
            continue
        root = read_xml(z, name)
        if root is None:
            continue
        instrs = []
        for el in root.iter(q('w:instrText')):
            if el.text:
                instrs.append(el.text)
        blob = ''.join(instrs)
        for m in RE_URL.findall(blob):
            out.append(Finding(fx, 'field instrText URL',
                               f'{m[:110]} ({name.split("/")[-1]})', 'WARN'))
        for m in RE_EMAIL.findall(blob):
            out.append(Finding(fx, 'field instrText email',
                               f'{m} ({name.split("/")[-1]})', 'BLOCKER'))
        # FILENAME / path leakage e.g. INCLUDETEXT "C:\\Users\\name\\..."
        for m in re.findall(r'[A-Za-z]:\\\\?Users\\\\?[^\s"]+', blob):
            out.append(Finding(fx, 'field instrText path', repr(m), 'BLOCKER'))


def scan_alttext(z, fx, out):
    """wp:docPr @name / @descr on drawings (alt text can hold real captions)."""
    for name in z.namelist():
        if not name.endswith('.xml'):
            continue
        root = read_xml(z, name)
        if root is None:
            continue
        for dp in root.iter(q('wp:docPr')):
            nm = (dp.get('name') or '').strip()
            desc = (dp.get('descr') or '').strip()
            # generic auto names like "Picture 1", "Image1" are benign
            for val, kind in [(nm, 'name'), (desc, 'descr')]:
                if not val:
                    continue
                if re.fullmatch(r'(Picture|Image|Graphic|Chart|Diagram|Object|'
                                r'Shape|Group|Text Box|Content Placeholder|'
                                r'Rectangle|Table)\s*\d*', val):
                    continue
                sev = 'BLOCKER' if RE_EMAIL.search(val) else 'WARN'
                out.append(Finding(fx, f'drawing alt-{kind}',
                                   f'{val!r} ({name.split("/")[-1]})', sev))


def scan_bookmarks(z, fx, out):
    root = read_xml(z, 'word/document.xml')
    if root is None:
        return
    names = set()
    for bm in root.iter(q('w:bookmarkStart')):
        nm = bm.get(q('w:name'))
        if nm and not re.fullmatch(r'(_GoBack|_Toc\d+|_Ref\d+|_Hlk\d+|OLE_LINK\d+|'
                                   r'_Int_\w+)', nm):
            names.add(nm)
    if names:
        # bookmark names are usually structural; only flag if they look like a
        # person/company (contain space + capitalised words) or an email.
        interesting = {n for n in names
                       if RE_EMAIL.search(n) or re.search(r'[A-Z][a-z]+ [A-Z][a-z]+', n)}
        if interesting:
            out.append(Finding(fx, 'bookmark names',
                               repr(sorted(interesting)[:8]), 'WARN'))


def scan_media_fonts_ole(z, fx, out):
    for name in z.namelist():
        low = name.lower()
        if low.startswith('word/media/'):
            data = z.read(name)
            h = hashlib.sha256(data).hexdigest()[:16]
            dim = ''
            try:
                from PIL import Image
                im = Image.open(io.BytesIO(data))
                dim = f'{im.width}x{im.height} {im.format}'
            except Exception:
                dim = 'unreadable(vector/EMF/WMF?)'
            out.append(Finding(fx, 'media', f'{name.split("/")[-1]} sha256:{h} '
                               f'{len(data)}B {dim}', 'REVIEW'))
        elif low.endswith('.odttf') or '/fonts/' in low:
            data = z.read(name)
            out.append(Finding(fx, 'embedded font',
                               f'{name.split("/")[-1]} {len(data)}B (licensed font data)',
                               'WARN'))
        elif low.endswith('vbaproject.bin'):
            out.append(Finding(fx, 'macros', 'vbaProject.bin present (VBA macros)',
                               'BLOCKER'))
        elif '/embeddings/' in low or low.endswith('.bin') and 'ole' in low:
            data = z.read(name)
            out.append(Finding(fx, 'OLE object',
                               f'{name.split("/")[-1]} {len(data)}B', 'WARN'))


def audit_docx(path, fx, out):
    try:
        z = zipfile.ZipFile(path)
    except zipfile.BadZipFile:
        out.append(Finding(fx, 'container', 'not a valid zip (damaged fixture?)', 'REVIEW'))
        return
    with z:
        scan_core(z, fx, out)
        scan_app(z, fx, out)
        scan_custom(z, fx, out)
        scan_authors(z, fx, out)
        scan_owner_leaks(z, fx, out)
        scan_hyperlinks(z, fx, out)
        scan_field_instructions(z, fx, out)
        scan_alttext(z, fx, out)
        scan_bookmarks(z, fx, out)
        scan_media_fonts_ole(z, fx, out)


def audit_owner_lock(path, fx, out):
    """~$*.docx Word owner file: byte0 = name length, then cleartext name."""
    try:
        d = open(path, 'rb').read()
    except OSError:
        return
    if not d:
        return
    n = d[0]
    try:
        name = d[1:1 + n].decode('latin-1', 'replace')
    except Exception:
        name = repr(d[1:40])
    out.append(Finding(fx, 'owner-lock username',
                       f'{name!r} (Word temp file — should not ship at all)', 'BLOCKER'))


def audit_pdf(path, fx, out):
    try:
        r = subprocess.run(['pdfinfo', path], capture_output=True, text=True, timeout=30)
    except Exception:
        return
    for line in r.stdout.splitlines():
        if ':' not in line:
            continue
        k, v = line.split(':', 1)
        k, v = k.strip(), v.strip()
        if k in ('Author', 'Creator', 'Producer', 'Title', 'Subject', 'Keywords') and v:
            if k == 'Author':
                sev = classify_person(v)
            elif k in ('Subject', 'Keywords', 'Title'):
                sev = 'BLOCKER' if RE_EMAIL.search(v) else 'WARN'
            else:  # Creator/Producer are usually 'Microsoft Word' — informational
                sev = 'OK'
            out.append(Finding(fx, f'pdfinfo {k}', repr(v), sev))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--root', default=os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    ap.add_argument('--md', default=None, help='write Markdown table here')
    ap.add_argument('--json', default=None, help='write JSON findings here')
    ap.add_argument('--dirs', nargs='*', default=[
        'apps/demo/public/fixtures', 'fixtures-staging', 'parity'])
    args = ap.parse_args()

    out = []
    scanned_docx = scanned_pdf = 0
    for d in args.dirs:
        base = os.path.join(args.root, d)
        if not os.path.isdir(base):
            continue
        for dirpath, _, files in os.walk(base):
            for f in sorted(files):
                p = os.path.join(dirpath, f)
                fx = os.path.relpath(p, args.root)
                if f.startswith('~$') and f.endswith('.docx'):
                    audit_owner_lock(p, fx, out)
                elif f.endswith('.docx'):
                    audit_docx(p, fx, out); scanned_docx += 1
                elif f.endswith('.pdf'):
                    audit_pdf(p, fx, out); scanned_pdf += 1

    out.sort(key=lambda x: (sev_rank(x.severity), x.fixture, x.channel))

    # console summary
    counts = {}
    for f in out:
        counts[f.severity] = counts.get(f.severity, 0) + 1
    print(f'Scanned {scanned_docx} docx + {scanned_pdf} pdf; findings by severity: '
          + ', '.join(f'{k}={v}' for k, v in sorted(counts.items())), file=sys.stderr)

    if args.json:
        with open(args.json, 'w') as fh:
            json.dump([f.row() for f in out], fh, indent=1)
    if args.md:
        with open(args.md, 'w') as fh:
            fh.write('| Fixture | Channel | Detail | Severity |\n')
            fh.write('|---|---|---|---|\n')
            for f in out:
                det = f.detail.replace('|', '\\|')
                fh.write(f'| `{f.fixture}` | {f.channel} | {det} | **{f.severity}** |\n')
    # also dump table to stdout
    for f in out:
        print(f'{f.severity:8} {f.fixture:60} {f.channel:26} {f.detail}')


if __name__ == '__main__':
    main()
