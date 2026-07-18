# Fixture corpus sources and staging notes

This file records provenance for public-source fixtures, whether they are still in
`fixtures-staging` or have been promoted. A fixture is benchmark-live only when its
DOCX is in `apps/demo/public/fixtures`, its Word PDF is in `parity`, and it has a
`parity/word-reference-manifest.json` entry.

Each entry is `<name>.docx` (the fixture) plus `<name>-word.pdf` (Word reference PDF,
exported from the *sanitized* .docx so line-break parity holds by construction).

---

## wild3 — public agency authoring templates

Five current government templates promoted to the live benchmark. These target
template-authoring structures that the earlier corpus only touched incidentally:
repeating header/footer layers, court forms, content controls, policy numbering,
embedded template fonts, and multi-section author guidance. All five were sanitized,
passed `scripts/validate-docx.py`, and received Word-exported PDF references.

| doc | source | Word pages | notable features |
|---|---|---:|---|
| `wild3-template-caed-pleading.docx` | U.S. District Court, Eastern District of California | 1 | official pleading-paper template; default/first/even header and footer parts plus the numbered pleading grid |
| `wild3-template-fws-manual.docx` | U.S. Fish & Wildlife Service | 5 | policy-manual chapter template; multilevel numbering, 4 tables, repeated footer lines, multiple header/footer parts |
| `wild3-template-nps-science-report.docx` | National Park Service | 12 | 3 sections, 374 paragraphs, 3 tables, 2 figures, extensive custom styles, embedded template fonts and mixed report examples |
| `wild3-template-us-courts-answer.docx` | Administrative Office of the U.S. Courts | 7 | dense defendant-answer form with 189 field markers and repeating court-form headers/footers |
| `wild3-template-uspto-follow-on.docx` | U.S. Patent and Trademark Office | 7 | structured follow-on filing template with 7 content controls, filing metadata and repeating page furniture |

### Source URLs

- CAED pleading paper: https://www.caed.uscourts.gov/caednew/index.cfm/attorney-info/word-format/ — `ArialPleadingPaper.docx`
- FWS Service Manual chapter: https://www.fws.gov/policy-library/e1011fw2 — `e1011fw2-word-template-revised-v2.docx`
- NPS Science Report: https://www.nps.gov/im/report-templates.htm — IRMA file `Simplified_Author_Template_v1.4.docx`
- U.S. Courts defendant answer: https://www.uscourts.gov/forms-rules/forms/defendants-answer-complaint — `the_defendants_answer_to_the_complaint.docx`
- USPTO DOCX templates: https://www.uspto.gov/patents/docx — `Follow-On-Template-August-2025.docx`

The USPTO initial-filing template was evaluated but left out because the source uses
an undefined `numId=0` and an out-of-schema `w:rPr` child order; the follow-on template
covers the same producer and content-control family while passing the hard validator.

---

## wild2 — real-world documents across content domains

Hunted from public sources to widen domain coverage (math, science, medical,
literature, legal). Every doc was verified as a real OOXML zip (`unzip -t`), run
through `scripts/validate-docx.py`, and — except public-domain literature — passed
through `scripts/sanitize-docx.py` (text scrambled to same-length/shape pseudowords;
OMML math in `m:t`, images, structure all preserved, so layout code paths and line
breaks are unchanged). Sources recorded below.

| doc | domain | source | ref PDF | sanitized | notable features |
|-----|--------|--------|--------:|:---------:|------------------|
| wild2-math-omml-dense.docx | math | Zenodo rec/439037 (Mottin 2016, CC) | retrying | yes | 38 `m:oMath` blocks, equation-dense math prose, 3 figures |
| wild2-math-eq-as-images.docx | math | Zenodo rec/1472644 (Fourier series, CC) | retrying | yes | equations rasterized as **41 inline images** (contrast to OMML doc) + 48 fields |
| wild2-sci-chem-omml.docx | science / chem | Zenodo rec/5011138 (Y2Mo4O15 preprint, CC) | **13 pp** | yes | 9 `m:oMath`, chemistry notation, 5 figures, 2 tables, 2 MB |
| wild2-sci-ieee-2col.docx | science | IEEE PES sample Word template | **4 pp** | yes | **two-column** layout, 4 figures, 1 footnote, 86 numPr; numId=0 quirk |
| wild2-sci-elsevier-template.docx | science | Elsevier "Research Article template and guidance" (legacyfileshare) | **11 pp** | yes | multi-column author template, numbered heads, guidance callouts |
| wild2-med-nccih-protocol.docx | medical | NIH NCCIH clinical-trial protocol template | **23 pp** | yes | 490 paras, TOC + 144 fields (cross-refs), 38 numPr, header/footer set — **blank template, no data** |
| wild2-med-phase23-protocol.docx | medical | NIH/FDA Phase 2-3 IND/IDE protocol template (via Feinstein) | **70 pp** | yes | 2049 paras, 7 tables, 656 numPr, 8 footnotes, 198 fields, SDT; numId=0 + rPr-order quirks — **blank template, no data** |
| wild2-legal-nih-contract.docx | legal | NIH OAMP DGS contract workform (508) | retrying | yes | **728 tables**, 376 fields, 3730 numPr, 10.5k paras — extreme table/field density |
| wild2-legal-ca-agreement.docx | legal | CA Courts JBCM Standard Agreement | **23 pp** | yes | numbered clauses (302 numPr), 4 tables, defined-terms layout |
| wild2-lit-yiddish-rtl.docx | literature | archive.org `FunemJarid` — Sholem Aleichem, *Funem Yarid* (YIVO orthography), public domain | retrying | **no** (public-domain text kept) | **RTL/bidi stress**: 62,497 `w:rtl` runs, 2214 justified paras, mixed Hebrew/Cyrillic/Latin, 3134 paras |

### Source URLs
- math-omml-dense: https://zenodo.org/records/439037 — `2016_mottin_37.docx`
- math-eq-as-images: https://zenodo.org/records/1472644 — `Fourier series expression-english.docx`
- sci-chem-omml: https://zenodo.org/records/5011138 — `article-preprint Y2Mo4O15.docx`
- sci-ieee-2col: https://ieee-pes.org/wp-content/uploads/2023/01/pg4-sample-word-template.docx
- sci-elsevier-template: https://legacyfileshare.elsevier.com/promis_misc/Research%20Article%20template%20and%20guidance.docx
- med-nccih-protocol: https://files.nccih.nih.gov/s3fs-public/CR-Toolbox/ProtocolTemplate_NCCIH_07-17-2015.docx
- med-phase23-protocol: https://feinstein.northwell.edu/sites/northwell.edu/files/2019-06/Clinical-Trial-Template-NIH-Phase-2-3-Protocol-Template.docx (NIH/FDA Phase 2-3 IND/IDE template)
- legal-nih-contract: https://oamp.od.nih.gov/sites/default/files/DGS/contracting-forms/dgs_contract_workform_11-10-16-508.docx
- legal-ca-agreement: https://courts.ca.gov/documents/JBCM-Standard-Agreement.docx
- lit-yiddish-rtl: https://archive.org/details/FunemJarid — `FunemJarid_YIVO.docx`

### Validation notes (producer quirks, NOT skipped)
`scripts/validate-docx.py` flags two docs; both are benign quirks Word tolerates on
read without a repair dialog, and are exactly the kind of wild-producer coverage worth
keeping for a future compat pass:
- **wild2-sci-ieee-2col** & **wild2-med-phase23-protocol**: `numId=0 used but not defined`
  — numId 0 is the OOXML "cancel inherited list numbering" reference; it legitimately
  has no `num` definition.
- **wild2-med-phase23-protocol**: `<w:rPr>` has `<w:shadow>` after `<w:vertAlign>`
  (out of the ECMA-376 rPr sequence). Word re-serializes in order on save; no repair.

### Not staged (findings)
- **Yiddish classic is large**: `wild2-lit-yiddish-rtl` has a ~16 MB `document.xml`
  (per-character RTL run formatting), though only 3134 paragraphs. Word export is
  feasible but slower; flag for the maintainer as an RTL stress fixture.
- Zenodo blocks binary downloads without full browser headers (403 "unusual traffic");
  `osp.od.nih.gov` `.docx` links return an HTML interstitial (not the file) to curl.

---

## staging-* — generated feature-coverage fixtures

A separate, complementary batch to the `wild2-*` real-world docs above: 12
**hand-generated** fixtures that each hammer one cluster of OOXML behavior the
current benchmark bed does not exercise yet. Source of truth is
`scripts/make-staging-fixtures.py` (raw-XML zip pattern, same lineage as
`scripts/make-parity2-more.py`). All are US Letter, deterministic, multi-page
where meaningful.

Regenerate / validate / export:

```
python3 scripts/make-staging-fixtures.py                      # all 12
python3 scripts/make-staging-fixtures.py grid4 bidi           # subset (stem after "staging-")
python3 scripts/validate-docx.py fixtures-staging/staging-*.docx   # hard gate -> 12/12 PASS
scripts/export-staging.sh staging-grid4 staging-bidi ...      # Word refs -> fixtures-staging/<name>-word.pdf
```

**Status: 12/12 validate, 12/12 exported from Word, 12/12 render with no engine crash.**

| Fixture | Coverage | Pages (Word) | Word ref | Engine sanity |
|---|---|---:|:---:|---|
| staging-grid4 | 5-level table nesting (L1>...>L5), gridSpan + vMerge at every level, per-level borders (single/double/dashed/dotted) + shading + cellMar; plus a ~39-row inner table in a wrapper cell that paginates | 3 | yes | renders, 3 pages, no errors |
| staging-tblextreme | table-in-textbox, textbox-in-cell, equation in a cell (OMML sSup/frac), footnote ref inside a nested table, a tab-heavy cell (L/C/R + decimal + dot leader), and list > table > list | 2 | yes | renders, 2 pages, no errors |
| staging-frames | `w:framePr` positioned paragraphs beyond drop caps: page-/margin-/text-anchored floating frames, `wrap="around"`, bordered + shaded, overlapping, body text wrapping past them | 2 | yes | renders, no errors |
| staging-styles | `basedOn` chain 5 deep with toggle b/i flips; char styles stacked (basedOn) + direct-format override; a table style with conditional firstRow/lastRow/firstCol/band1Horz driven by `tblLook` | 1 | yes | renders, no errors |
| staging-fields2 | PAGEREF + REF `\h` cross-refs, SEQ chain with `\r 1` restart, IF field, MERGEFIELD (cached), legacy FORMTEXT + FORMCHECKBOX (`ffData`), XE entries + cached INDEX | 2 | yes | renders, 2 pages, no errors |
| staging-hf2 | different-first + even/odd headers across 3 sections (`titlePg` + `evenAndOddHeaders`), per-section watermarks (rotated anchored shapes: DRAFT/CONFIDENTIAL/FINAL), roman->arabic `pgNumType` restart | 5 | yes | renders, no errors |
| staging-bidi | RTL paragraphs (Arabic + Hebrew) `w:bidi` + `w:rtl`, mixed LTR runs inside RTL paragraphs, an RTL (`bidiVisual`) table | 1 | yes | renders, no errors |
| staging-eastasian | CJK (Japanese + Chinese) with `w:kinsoku` + `overflowPunct`, `eastAsia` font, a combined-char run (`eastAsianLayout w:combine`), a `docGrid type=lines linePitch=360` section | 1 | yes | renders, no errors |
| staging-breaks | `textWrapping` break beside a float, two consecutive page breaks, page break before a table, a 3-column continuous section with explicit column breaks | 7 | yes | renders, no errors |
| staging-anchors2 | three overlapping floats, ascending z-order + `allowOverlap` toggled + one `locked`; positioning relative-from page/margin/column/character/line; a `wp14` percent-sized box | 2 | yes | renders, 2 pages, no errors |
| staging-longtable | 200-row table, `tblHeader` repeat, `cantSplit` tall rows (every 25th), `atLeast` rows (every 7th), `exact` short rows (every 5th), banded shading | 9 | yes | renders, 9 pages (matches Word), no errors |
| staging-typography | `w:w` scaling 50/100/150/200%, letter-spacing -1pt..+6pt, `w:kern` threshold at 12pt vs 20pt, ligature-rich text, nbsp / soft-hyphen / em+en dashes, `autoHyphenation` (justified) | 1 | yes | renders, no errors |

### Validation & tooling

- `scripts/validate-docx.py` was extended to accept a **directory argument**
  (globs `*.docx` inside it) in addition to explicit paths; its no-argument
  default (`apps/demo/public/fixtures/parity2-*.docx`) is unchanged.
- Second opinion: LibreOffice converts all 12 without error (lenient about the
  two Word-only issues below).
- Word references exported via `scripts/export-staging.sh` (open-by-name /
  save-as-PDF / verify page count + first line — mirrors `export-parity2.sh`).

### Engine sanity (optional, no baselines recorded)

Each fixture was loaded once at `http://localhost:5299/?doc=/fixtures/<name>.docx`
(temporarily copied into the live fixtures dir, then **removed** — no
`parity/*-word.pdf` means it never joined any report). **No engine crashes**: all
12 render, no engine exceptions in console (only a benign React `createRoot`
dev-shell warning). Page counts spot-checked matched Word (grid4 3/3,
tblextreme 2/2, longtable 9/9). Pixel/vertical parity was not measured — deferred
until promotion.

### Two Word-only rejections found and fixed (validator blind spots)

Both passed the structural validator and LibreOffice, but Word declined to open
them (silently — no repair dialog, empty AppleScript result / -2753):

1. **Bare `<w:r>` under `<w:body>` (staging-fields2):** the `XE` helper returns a
   run; three were placed *between* paragraphs, leaving a run as a direct child
   of the body. Fixed by embedding each XE run inside its paragraph. The
   validator does not check for stray inline content in block context.
2. **`endnotePr` separators declared with no `endnotes.xml` part
   (staging-tblextreme):** the shared settings helper declared both footnote and
   endnote separator pseudo-notes while the fixture ships only `footnotes.xml`
   (the defect class already noted in `docs/FIXTURES.md`). Fixed by splitting
   into independent `footnotes=`/`endnotes=` flags. The validator checks that
   *referenced* note ids exist, but not that *declared settings separators* have
   a matching part.

Both are good candidates to teach `scripts/validate-docx.py` in a follow-up.

### Not expressed here (next gaps)

- **Percent position offsets** (`wp14:pctPosHOffset` in a
  `mc:AlternateContent`-wrapped `positionH/V`) were omitted from staging-anchors2
  to keep the anchor Word-valid without hand-verifying that extension's child
  order; it covers percent *sizes* (`wp14:sizeRelH/sizeRelV`) instead.
- **Watermarks** in staging-hf2 are rotated anchored `wps` shapes, not VML
  `v:textpath` WordArt (already a known engine gap — parity2-watermark); this
  fixture targets header-per-section + pgNumType behavior, not the WordArt path.
- Charts-as-fallback-image, OLE placeholders and ruby were considered but not
  built (they need embedded binary parts + bespoke content types); flagged as
  the obvious next additions.
