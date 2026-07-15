#!/usr/bin/env python3
"""Pre-Word validation gate for generated .docx fixtures.

A fixture that Word has to REPAIR is not trustworthy ground truth (Word's
repaired in-memory doc may not match the .docx bytes our engine renders), and a
corrupt open can wedge the AppleEvents queue for a whole export batch. So every
fixture must pass this gate BEFORE it goes to Word.

Checks (in order; any failure = do NOT export):
  1. zip integrity (testzip).
  2. [Content_Types].xml present.
  3. every .xml / .rels part is well-formed XML.
  4. every relationship Target resolves to a part that exists in the zip.
  5. every part carries a content type (Default extension or Override).
  6. document.xml r:id / r:embed references resolve to a declared relationship.
  7. footnote / endnote ids referenced in document.xml exist in the notes part
     (and the -1/0 separator pseudo-notes are present); no orphan note ids.
  8. numId -> num -> abstractNumId all resolve.
  9. ORDERED-SEQUENCE + duplicate check for the OOXML complex types whose
     children are a schema <xsd:sequence> (w:pPr, w:rPr, w:tcPr, w:tblPr,
     w:sectPr, w:settings, w:trPr). Out-of-order or duplicated singleton
     children are THE classic raw-XML cause of Word's "corrupt / repaired"
     dialog, and LibreOffice/our engine silently tolerate them - so this,
     not an external converter, is the real Word-corruption detector.

Usage: python3 scripts/validate-docx.py <name-or-path> [...]
        python3 scripts/validate-docx.py            # all parity2-*.docx
Exit code is nonzero if any fixture fails.
"""
import sys, os, re, zipfile, glob
import xml.etree.ElementTree as ET

FIX = os.path.join(os.path.dirname(__file__), "..", "apps", "demo", "public", "fixtures")

def LN(tag):  # strip namespace
    return tag.split("}", 1)[1] if "}" in tag else tag

# --- canonical child orders (subset covering what the generators emit; each is
# an xsd:sequence in ECMA-376). Elements not listed are ignored for ordering. ---
ORDERS = {
    "rPr": ["rStyle","rFonts","b","bCs","i","iCs","caps","smallCaps","strike","dstrike",
            "outline","shadow","emboss","imprint","noProof","snapToGrid","vanish","webHidden",
            "color","spacing","w","kern","position","sz","szCs","highlight","u","effect","bdr",
            "shd","fitText","vertAlign","rtl","cs","em","lang","eastAsianLayout","specVanish",
            "oMath","rPrChange"],
    "pPr": ["pStyle","keepNext","keepLines","pageBreakBefore","framePr","widowControl","numPr",
            "suppressLineNumbers","pBdr","shd","tabs","suppressAutoHyphens","kinsoku","wordWrap",
            "overflowPunct","topLinePunct","autoSpaceDE","autoSpaceDN","bidi","adjustRightInd",
            "snapToGrid","spacing","ind","contextualSpacing","mirrorIndents","suppressOverlap",
            "jc","textDirection","textAlignment","textboxTightWrap","outlineLvl","divId","cnfStyle",
            "rPr","sectPr","pPrChange"],
    "tcPr": ["cnfStyle","tcW","gridSpan","hMerge","vMerge","tcBorders","shd","noWrap","tcMar",
             "textDirection","tcFit","vAlign","hideMark","cellIns","cellDel","cellMerge","tcPrChange"],
    "trPr": ["cnfStyle","divId","gridBefore","gridAfter","wBefore","wAfter","cantSplit","trHeight",
             "tblHeader","tblCellSpacing","jc","hidden","ins","del","trPrChange"],
    "tblPr": ["tblStyle","tblpPr","tblOverlap","bidiVisual","tblStyleRowBandSize","tblStyleColBandSize",
              "tblW","jc","tblCellSpacing","tblInd","tblBorders","shd","tblLayout","tblCellMar",
              "tblLook","tblCaption","tblDescription","tblPrChange"],
    "sectPr": ["headerReference","footerReference","footnotePr","endnotePr","type","pgSz","pgMar",
               "paperSrc","pgBorders","lnNumType","pgNumType","cols","formProt","vAlign","noEndnote",
               "titlePg","textDirection","bidi","rtlGutter","docGrid","printerSettings","sectPrChange"],
    # settings: only the tail we emit; earlier settings elements are omitted (ignored).
    "settings": ["footnotePr","endnotePr","compat"],
}
# headerReference/footerReference share order slot 0 in sectPr.
SECT_ALIAS = {"footerReference": "headerReference"}

R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

def check_order(root, errs):
    """Walk the tree; for every element whose localname is an ordered container,
    verify its recognized children are non-decreasing in canonical index and no
    singleton repeats."""
    for el in root.iter():
        cont = LN(el.tag)
        order = ORDERS.get(cont)
        if not order:
            continue
        idx = {name: i for i, name in enumerate(order)}
        last = -1
        last_name = None
        seen = {}
        for child in list(el):
            cln = LN(child.tag)
            key = SECT_ALIAS.get(cln, cln) if cont == "sectPr" else cln
            if key not in idx:
                continue
            i = idx[key]
            if i < last:
                errs.append(f"<w:{cont}> child <w:{cln}> is out of order "
                            f"(after <w:{last_name}>); schema sequence requires it earlier")
            last, last_name = max(last, i), cln
            # duplicate singleton (allow repeated headerReference/footerReference)
            if cont in ("rPr","pPr","tcPr","trPr","tblPr") or (cont == "sectPr" and key not in ("headerReference",)):
                seen[cln] = seen.get(cln, 0) + 1
                if seen[cln] == 2:
                    errs.append(f"<w:{cont}> has duplicate <w:{cln}> (schema allows one)")

def validate(path):
    errs = []
    name = os.path.basename(path)
    try:
        z = zipfile.ZipFile(path)
    except Exception as e:
        return [f"not a zip: {e}"]
    bad = z.testzip()
    if bad:
        errs.append(f"corrupt zip entry: {bad}")
    names = set(z.namelist())
    if "[Content_Types].xml" not in names:
        errs.append("missing [Content_Types].xml")

    # parse all xml/rels
    trees = {}
    for n in names:
        if n.endswith(".xml") or n.endswith(".rels"):
            try:
                trees[n] = ET.fromstring(z.read(n))
            except Exception as e:
                errs.append(f"XML not well-formed: {n}: {e}")

    # content types
    ct = trees.get("[Content_Types].xml")
    defaults, overrides = set(), set()
    if ct is not None:
        for c in ct:
            if LN(c.tag) == "Default":
                defaults.add(c.attrib.get("Extension", "").lower())
            elif LN(c.tag) == "Override":
                overrides.add(c.attrib.get("PartName", ""))
    # every part (except dirs and _rels and [Content_Types]) needs a CT
    for n in names:
        if n.endswith("/") or n == "[Content_Types].xml" or "/_rels/" in n or n.startswith("_rels/"):
            continue
        ext = n.rsplit(".", 1)[-1].lower() if "." in n else ""
        part = "/" + n
        if part not in overrides and ext not in defaults:
            errs.append(f"part has no content type (no Override, extension .{ext} not Default): {n}")

    # relationships resolve
    for n, root in trees.items():
        if not n.endswith(".rels"):
            continue
        base = os.path.dirname(os.path.dirname(n))  # _rels lives beside its owner dir
        for rel in root:
            tgt = rel.attrib.get("Target", "")
            mode = rel.attrib.get("TargetMode", "")
            if mode == "External" or tgt.startswith("http"):
                continue
            resolved = os.path.normpath(os.path.join(base, tgt)).replace("\\", "/")
            if resolved not in names:
                errs.append(f"{n}: relationship {rel.attrib.get('Id')} -> {tgt} (resolved {resolved}) missing")

    # document r:id references resolve
    drels = trees.get("word/_rels/document.xml.rels")
    rel_ids = set()
    if drels is not None:
        rel_ids = {r.attrib.get("Id") for r in drels}
    doc = trees.get("word/document.xml")
    if doc is not None:
        used = set()
        for el in doc.iter():
            for a, v in el.attrib.items():
                if a in (f"{R}id", f"{R}embed", f"{R}link"):
                    used.add(v)
        for rid in used:
            if rid not in rel_ids:
                errs.append(f"document.xml references {rid} with no relationship")

        # footnote / endnote id cross-check
        def note_ids(part, refln):
            refs = {el.attrib.get(f"{{http://schemas.openxmlformats.org/wordprocessingml/2006/main}}id")
                    for el in doc.iter() if LN(el.tag) == refln}
            root = trees.get(part)
            if refs and root is None:
                errs.append(f"document.xml has {refln} but {part} is missing")
                return
            if root is None:
                return
            defined = {}
            for el in root:
                if LN(el.tag) in ("footnote", "endnote"):
                    defined[el.attrib.get(f"{{http://schemas.openxmlformats.org/wordprocessingml/2006/main}}id")] = True
            for want in ("-1", "0"):
                if want not in defined:
                    errs.append(f"{part}: missing separator pseudo-note id={want}")
            for rid in refs:
                if rid not in defined:
                    errs.append(f"{refln} id={rid} in document.xml not defined in {part}")
        note_ids("word/footnotes.xml", "footnoteReference")
        note_ids("word/endnotes.xml", "endnoteReference")

        # numbering resolve
        numbering = trees.get("word/numbering.xml")
        if numbering is not None:
            NUM = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
            nums, abstracts = {}, set()
            for el in numbering:
                if LN(el.tag) == "num":
                    nid = el.attrib.get(f"{NUM}numId")
                    an = None
                    for c in el:
                        if LN(c.tag) == "abstractNumId":
                            an = c.attrib.get(f"{NUM}val")
                    nums[nid] = an
                elif LN(el.tag) == "abstractNum":
                    abstracts.add(el.attrib.get(f"{NUM}abstractNumId"))
            used_num = {el.attrib.get(f"{NUM}val") for el in doc.iter()
                        if LN(el.tag) == "numId"}
            for nid in used_num:
                if nid not in nums:
                    errs.append(f"numId={nid} used but not defined in numbering.xml")
                elif nums[nid] not in abstracts:
                    errs.append(f"numId={nid} -> abstractNumId={nums[nid]} not defined")

    # ordered-sequence + duplicate checks on the main content parts
    for n, root in trees.items():
        if n.endswith(".rels") or n == "[Content_Types].xml":
            continue
        check_order(root, errs)

    return errs


def main():
    args = sys.argv[1:]
    if args:
        paths = [a if os.path.exists(a) else os.path.join(FIX, a if a.endswith(".docx") else a + ".docx")
                 for a in args]
    else:
        paths = sorted(glob.glob(os.path.join(FIX, "parity2-*.docx")))
    fails = 0
    for p in paths:
        if not os.path.exists(p):
            print(f"MISSING  {p}"); fails += 1; continue
        errs = validate(p)
        nm = os.path.basename(p)
        if errs:
            fails += 1
            print(f"FAIL  {nm}")
            for e in errs:
                print(f"        - {e}")
        else:
            print(f"PASS  {nm}")
    print(f"\n{len(paths)-fails}/{len(paths)} passed")
    sys.exit(1 if fails else 0)


if __name__ == "__main__":
    main()
