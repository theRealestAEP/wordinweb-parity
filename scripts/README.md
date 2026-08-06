# Scripts

The public scripts support three workflows.

## Selecting the engine under measurement

Every gate here measures whatever wordinweb the demo loads, so that choice has
to be explicit:

```bash
node scripts/use-engine.mjs ../wordinweb-likeoffice/packages/react   # a local build
node scripts/use-engine.mjs npm:0.1.22                              # a published version
node scripts/use-engine.mjs --help                                  # print what is selected now
```

The demo imports wordinweb from `apps/demo/src/main.tsx`, so Node and vite
resolve `apps/demo/node_modules/wordinweb` **before** the root one. Setting only
the root link therefore leaves the demo loading something else entirely, and
nothing says so: vite serves a pre-bundled copy and only re-reads the package
when it re-optimizes, so a dev-server restart can swap the engine mid-campaign.
That happened here — a stale published 0.1.22 sat in `apps/demo/node_modules`
while the root link named a local build, and a restart silently moved the suite
from 10/10 to 5/10.

`use-engine.mjs` sets **both** locations to the same engine, clears vite's dep
cache (`apps/demo/node_modules/.vite`, otherwise the next server keeps serving
the old pre-bundle), and prints the resolved version, realpath and git SHA. A
local path must already be built; linking a package whose `dist/` predates its
source measures the old code. Restart the dev server afterwards.

`apps/demo/package.json` declares a published `wordinweb`, so a plain
`npm install` resets the selection back to that version. Re-run `use-engine.mjs`
after any install, and check `--help` if you are unsure what is selected.

`engine-provenance.mjs` holds the resolution logic, shared so the selector and
the gates cannot disagree about what "the engine" is.
`edit-roundtrip-parity.mjs` refuses to run at all when the two locations
disagree, rather than producing results that describe an engine nobody chose.

## Cross-editor compatibility

`interop-smoke.mjs` is the structural cross-editor compatibility gate. It saves
a small representative corpus of table and tab-stop documents through
WordInWeb, sends each candidate through LibreOffice, and checks the re-exported
DOCX and PDF for retained tables, rows, cells, tab stops, text, bounded page
counts, and nonblank rendered content:

```bash
npm run test:interop
```

Pass `--google` (or use `npm run test:interop:google`) to run the same candidates
through native Google Docs import/export. Authentication uses
`GOOGLE_INTEROP_ACCESS_TOKEN` or `GOOGLE_INTEROP_SERVICE_ACCOUNT_JSON`, with an
optional `GOOGLE_INTEROP_FOLDER_ID` for a shared Drive destination. The script
deletes each imported Google Doc after its exports are checked.
Each successful run also writes the compact compatibility manifest and page
previews consumed by the Google Docs and LibreOffice tabs on `/report/`.

## Visual parity

- `parity-parallel.mjs` is the canonical runner for every complete or large
  Word comparison. It shards fixtures and large page ranges across workers.
- `parity-compare.mjs` runs selected fixtures and serves as the worker process
  launched by `parity-parallel.mjs`.
- `parity-render-report.mjs` rebuilds the HTML report from saved results.
- `parity-report.mjs` contains the shared report generator.
- `parity-metric.mjs` holds the canonical per-page metric (`severityPct` and the
  appearance channels) that `parity-compare.mjs` and `edit-roundtrip-parity.mjs`
  both evaluate in the browser.
- `use-engine.mjs` selects the wordinweb build every gate measures, and
  `engine-provenance.mjs` resolves and describes it.
- `word-download-parity.mjs` is the saved-DOCX release gate. It clicks the
  demo's built-in Download button, exports only that candidate with desktop
  Microsoft Word, rasterizes both Word PDFs at 192 DPI, and compares against
  the cached `parity/<fixture>-word.pdf` reference. A complete run refreshes the
  dashboard served at `/report/` from these exact results.
- `edit-roundtrip-parity.mjs` is the edit round-trip gate. It applies a scripted
  edit sequence in the demo, downloads the edited DOCX, and holds desktop Word's
  rendering of that file against the web renderer's.
- `edit-roundtrip-scenarios.mjs` holds the scenario library that gate runs.
- `word-export.mjs` holds the Word automation and raster helpers both DOCX gates
  share.
- `word-parity.sh` exports one source DOCX to a Word reference PDF. Use it only
  when intentionally updating source-of-truth references.
- `word-parity-all.sh` intentionally updates the standard reference set.

Run a complete parity check with the parallel runner, then publish its finished
report to the demo only when the `/report/` artifacts should be refreshed:

```bash
node scripts/parity-parallel.mjs
npm run report:snapshot
```

Set `DXW_PARITY_JOBS` to override the default worker count when needed. Direct
`parity-compare.mjs` runs are reserved for focused fixture work.

Run the candidate-only gate while the demo is available at port 5299:

```bash
node scripts/word-download-parity.mjs parity-text benchmark
```

Omit fixture names to test every fixture with a cached Word reference. The
runner stages Word automation files under
`~/Library/Containers/com.microsoft.Word/Data/Documents/WordInWebParity`, so
Microsoft Word does not request access to each temporary file after Full Disk
Access has been granted. Candidate PDFs are cached there by the SHA-256 of the
sorted DOCX package entries and their exact uncompressed bytes, so ZIP metadata
does not cause an unchanged document to be exported by Word again while XML
byte changes still require a new Word export. Candidate rasters and reference
rasters are cached by the exact Word-PDF SHA-256. Browser screenshots, browser
PDFs, report PNGs, and off-page comment UI are excluded from this gate.

### What makes a reference ground truth

A reference PDF is ground truth only when Word had to compute the layout in it,
and only when someone else can compute it again. Two things break that, and both
have caught real fixtures:

- **Stored pagination.** Word records where it last broke the pages as
  `w:lastRenderedPageBreak` and replays them until an edit disturbs the file, so
  a reference exported from a hint-carrying package can show a layout Word would
  no longer produce. Strip the hints with
  `scripts/strip-pagination-hints.py <source> <destination>` and export from the
  stripped copy. The strip is byte-level, so the copy stays a checkable function
  of its source.
- **A package Word refuses.** An untyped `docProps/custom.xml` that `_rels/.rels`
  references makes Word stop with a modal AppleScript cannot answer. Repair it
  with `scripts/fix-custom-properties-type.py` before exporting anything.

When a fixture needs either treatment, the reference is exported from a copy
under `parity/word-reference-docx/`, and the manifest's `referenceDocx` names
that copy. Keep the copy derivable from the corpus fixture by the two scripts
above and nothing else, so its lineage can be re-checked at any time by
re-deriving it and comparing.

### The build-drift screen (#58)

111 of the cached references were exported between 2 and 22 July 2026 by an older
Word build, and #55 proved that build differs behaviourally (VML extent rounding).
The screen re-exports each reference from the exact DOCX its manifest names, on
the Word build installed now, and compares page count first and then every page
rasterized at 192 DPI byte for byte. Only the build varies — same package, hints
included, same machine, same fonts. Hints are deliberately NOT stripped here:
#48 settled the hint variable already, and stripping would confound it with the
build variable.

| wave | references | pages | reproduce byte for byte | differ |
| --- | --- | --- | --- | --- |
| 1 (≤5 pages) | 76 | 142 | 60 | 16 |
| 2 (6–39 pages) | 16 | 214 | 10 | 6 |
| 3 (≥40 pages) | 3 | 680 | 3 | 0 |
| **total** | **95** | **1036** | **73** | **22** |

Wave 3 is the whole result in miniature: wild-multicolumn (46pp),
wild2-lit-yiddish-rtl (215pp, 206 hints, right-to-left) and
wild2-legal-nih-contract (419pp, 88 hints) reproduce byte for byte on every one
of their 680 pages. Two thirds of the corpus by page count, the two hardest
documents in it, and not one pixel moved. Only 32 of 1036 pages differ at all.

The drift itself is bounded. Every difference except the ones named below is one
of three harmless shapes: a few words nudged horizontally within a line, a whole
block nudged vertically (maximum 0.260pt), or an embedded JPEG re-encoded in the
same box at the same resolution. None reflows and none changes a page count.

**A difference is not drift until the same build reproduces itself.** Export one
package twice on the build installed now and compare the two exports with each
other; only a fixture that is stable under that control can have its difference
against the cached reference attributed to the older build. Eight fixtures were
controlled this way and seven are perfectly stable, which is what licences
reading their differences as build drift. The confirmed ceiling is **0.378pt**,
on wild-gatech.

wild-doerfp is the exception and the reason the control exists. Four exports of
the same package on the current build produce TWO different page 35s, and one of
them is byte-identical to the July reference. Current Word can produce the cached
layout and simply does not do so every time, so wild-doerfp's 0.554pt is
bistability, not drift — and a single re-export is never enough to call a
reference stale.

Beware the raw pixel percentage on dense pages: wild2-sci-ieee-2col reads 1.39%
from a single 0.26pt block shift, because two-column text has enormous edge
length per unit area.

The differences that were NOT harmless:

- **probe3-lo-provenance** — the only proven behavioural build change. Its
  SourceText style names "Liberation Mono", which neither build resolves; the
  July export fell back to Courier New and current Word falls back to Calibri.
  Reference replaced; severity 0.050% → 1.380%.
- **wild3-template-caed-pleading** — its reference never came from the committed
  fixture (recorded source hash bc31b7bf, fixture 4c689596, unchanged since
  01dbbb7, and the PDF carries no Creator). Re-exported, the body sits 120.04pt
  lower. Reference replaced; severity 0.000% → 36.810%.
- **wild2-legal-ca-agreement** — the corpus's only page-count change, 23 → 22,
  already known from #38 and #48 as replayed pagination rather than a computed
  layout. Now it will not reproduce even WITH its hints intact. Reference
  replaced; see "ca-agreement's 23rd page was never Word's" below.
- **parity2-fields** and **probe3-field-switches** — these embed their own export
  date, so they differ on every re-export by construction and are not drift.

Two references were replaced. Everything else was left alone deliberately:
re-exporting a tenth of a point is churn that discards the July baseline for no
measurable gain.

### wild-athabasca: what an unprovenanced reference costs

`parity/wild-athabasca-word.pdf` and `parity/wild-wirfp-word.pdf` arrived with
the workspace (1d0dc82) already exported, on a machine and from packages this
repository never held. Their corpus fixtures could not open in Word at all, so
the manifest named a source that demonstrably did not produce them. Both PDFs
also carry `Title`/`Subject`/`Keywords` of "Fixture" where the corpus files leave
those elements empty, which is a second, independent tell that the export source
was a sibling package rather than the fixture.

Re-exported from the repaired fixtures (f974a48) both references reproduce:
wirfp 20 of 20 pages byte-identical at 192 DPI, athabasca 30 of 31, and
athabasca's page 2 differs only in two lines whose words, baselines and widths
all match while one inter-word space is 0.062pt wider and one word 0.076pt
narrower. The keepNext chain the engine cites on this fixture is present in the
fresh export: document paragraphs 211-217 are seven consecutive Heading2/Heading3
paragraphs, they land together at the top of page 20 with the Normal paragraph
that terminates the chain, and page 19 ends 317.4pt early — 11.5 slots at that
run's 27.5pt line pitch.

So an unprovenanced reference is not automatically a wrong one. What it costs is
the ability to say so without re-exporting, and every measurement standing on it
stays provisional until someone does.

## Edit round-trip

The saved-DOCX gate proves the website re-serializes a document Word already
agreed with. The edit round-trip gate proves the harder half: that a document
the website EDITED still means the same thing to Word. Run it while the demo is
available at port 5299:

```bash
npm run parity:edit-roundtrip                                # every scenario
node scripts/edit-roundtrip-parity.mjs --scenario typing     # one, repeatable
```

Each scenario loads a fixture in the demo, applies a scripted edit sequence
through the editor api (published on `window.__dxwApi` only under `?apihook=1`),
and clicks the built-in Download. Four structural checks are unconditional hard
failures:

- the downloaded DOCX is a readable package;
- desktop Word opens and exports it — Word's repair prompt is modal and never
  answers AppleScript, so a damaged package surfaces as a failed open;
- re-opening and re-saving the edited DOCX produces byte-identical output;
- Word and the web renderer agree on the page count.

Pages are then scored, and both the mean and the worst page must stay within
the thresholds.

### Which metric, and why

Pages are graded on `severityPct` from `parity-metric.mjs` — the same
`ink-dilate-line-v5` measurement the corpus gate uses — and **not** on raw
mismatched-pixel percentage.

Raw mismatch cannot grade a web-vs-Word page. Word-PDF and Chrome disagree on
sub-pixel glyph placement across every line of text, so a perfectly correct page
still lands around 1%, and the tracked corpus averages 5.29% raw over 1188
pages. No raw threshold separates a correct round trip from a broken one. The
`typing` scenario makes the point: 1.34% raw, 0.00% severity. So does
`header-footer`: 6.72% raw, 0.40% severity — rasterization noise, not a defect.

`severityPct` registers one global page offset, then counts only ink with no
counterpart within a small spatial tolerance, plus line reflow corroborated by
independent misalignment. Different-looking glyphs at the same place score
zero; missing, extra, reflowed or displaced content scores in full.

Calibration, from the last full corpus run in `parity/history.jsonl` (1188
pages, same metric version): severity mean 0.358%, median 0.00%, p95 0.55%.
Thresholds are **mean ≤ 1%** and **worst page ≤ 5%**. The mean sits near 3x the
corpus mean and above its p95. Only 1.26% of corpus pages exceed 5%, and 5%
stays below the metric's own structural-classification floor (`STRUCT_LO` = 10),
so a page the corpus metric would call structurally broken fails this gate with
margin. Raw `mismatchPct` is still recorded per page as context.

Each compared page also writes the Word | web | diff triptych under
`diff-png/<scenario>/`, so a failure is diagnosable without re-running.

### Baseline attribution

Every scenario is also scored against its fixture **as authored**, before any
edit, and each page records four numbers rather than one: its edited severity,
its baseline severity, how far OUR render moved from our own baseline, and how
far WORD's moved from its. Failures are then classified automatically:

- `present-in-baseline` — the page already differed before the edit, so this is
  not the scenario's doing;
- `edit-introduced` — our render changed and now disagrees;
- `word-reacted` — our render held still and Word's moved.

Every one of these names an ASYMMETRY. None of them names a culprit, and each
has now been read as one at least once. `present-in-baseline` used to say "so
this is a renderer issue"; the section fix made that wrong too, and the wording
above is what is left after removing the claim.

`word-reacted` says WHICH side moved. It does NOT say which side is right, and
reading it as "our render is fine" is a mistake this gate has already made. On
field-update the label is accurate and the conclusion drawn from it was wrong:
Word raises page 4 by 34 CSS px and we do not, and Word reaches that same
position on its own — open the unedited document, press F9, re-save, render —
with nothing of ours involved. So 34 CSS px up is the CORRECT post-update
position, Word reflows to it, and our layout fails to. The label pointed at the
right asymmetry and the wrong culprit.

Settling that needs one more export than the gate performs, and it is worth
performing by hand on any `word-reacted` failure: have Word do the equivalent
edit itself, end to end, and see where it lands. If Word agrees with itself,
the movement is correct and ours is the side to fix.

The 34 CSS px is now isolated to two layout rules, measured by
`scripts/generate-sectcontinuous-probe.mjs` and its two Word references. Both
apply to the first paragraph of a page that a SECTION break created, and the
field update is incidental to both:

- a following section marked `<w:type w:val="continuous"/>` makes us restart the
  new page at the flow offset the previous section left, re-adding the
  terminating paragraph's line and its space-after — 19.80 pt. Word treats
  `continuous` and `nextPage` alike once a page break has already moved on;
- we then apply the paragraph's whole `w:spacing w:before`, where Word applies
  `max(0, before - the previous paragraph's space-after)` — a further 6.00 pt
  here.

A page created by a plain `<w:br w:type="page"/>` with no section break is
already right on both counts: Word and we both put the paragraph at the body
top.

That supersedes the split in the commit before it, which read Word as
suppressing space-before outright at the page top and left ~19 CSS px
unaccounted. Word suppresses it outright only after a plain page break; after a
SECTION break it keeps `before - previous space-after`, so our excess there is
8 CSS px rather than 16, and the rest — 26.40 CSS px — is the continuous-section
carry-over. Both readings sum to the same 34, and nothing is left over.

The other candidate for that residual — that the section's first-page body
origin tracks `w:header` distance rather than the top margin — is excluded.
Varying `w:header` over 0, 360, 1440 and 2880 twips, on the section that ends at
the break, on the one that begins after it, and on both, leaves the page-4 gap
at exactly 42.40 CSS px in all twelve. The same run's control, turning that
section's `continuous` into `nextPage`, moves it to 16.00. Header distance is
inert for this position.

Both rules are fixed as of engine `918da3a`, and the probe now reads what Word
reads: 6.00 pt above the body top for a continuous section start, 6.00 for a
nextPage one, 0.00 after a plain page break. field-update goes to severity mean
0.000%, worst 0.000% over 23 pages.

**That fix is what makes `wild2-med-nccih-protocol`'s BASELINE diverge**, and the
divergence is the reference's, not ours. `parity/wild2-med-nccih-protocol-word.pdf`
was exported from a file still carrying its stored `lastRenderedPageBreak`
hints, so Word replayed the stored pagination instead of computing one, and
page 4 sat 34 CSS px below where Word's own layout puts it. We used to match
that stale position and correctly stopped, so the page-4 baseline read 39.100%
and the gate labelled it `present-in-baseline` — the third distinct thing these
three labels have been misread as saying, and no renderer defect at all.

That reference is now re-exported from a hint-stripped copy of the same
document, and the fixture reads **mean 0.000%, worst 0.000% over all 23 pages**,
down from mean 1.700% / worst 39.100%. `parity/word-reference-docx/` holds the
DOCX it was exported from; that file now carries two deliberate transforms
against the corpus fixture — the content-type repair that lets Word open it at
all, and the `w:lastRenderedPageBreak` strip that makes Word compute a layout
instead of replaying a stored one. Neither changes a glyph.

**A reference is only ground truth if Word had to compute it.** Exporting a
hint-carrying fixture untouched gets you the pagination the file remembers,
which for a SANITIZED fixture is the pagination of the text that was there
before sanitizing. 17 of the 101 corpus fixtures with a cached reference carry
hints and are exposed to this; see the `#43` survey.

### ca-agreement's 23rd page was never Word's

`wild2-legal-ca-agreement` is the same story, and it is now settled with the
control the drift screen demands. Its reference is re-exported and the corpus
baseline moves from **23 pages to 22**.

`parity/word-reference-docx/wild2-legal-ca-agreement.docx` is the corpus fixture
with its 17 `w:lastRenderedPageBreak` hints stripped — 442 bytes, nothing else —
and `parity/wild2-legal-ca-agreement-word.pdf` is what current Word exports from
it. Four measurements, all on the build installed now:

| package | exports | pages | agreement |
| --- | --- | --- | --- |
| hint-stripped | 3 | 22 | byte identical on all 22 pages |
| hints intact | 2 | 22 | byte identical on all 22 pages |
| stripped vs hinted | — | 22 | byte identical on all 22 pages |

So this document is **stable**, not bistable like wild-doerfp, and its stored
pagination is now completely **inert**: current Word computes the same 22 pages
whether the hints are there or not. Both halves matter. The first licences
calling the cached 23-page reference stale on a re-export; the second says the
staleness is not a hint-replay Word still performs, it is a layout the July
build produced and this one does not.

The extra page in the old reference is a blank verso at page 2 — footer band
only, no body ink. Under current Word `KACUCUJI A` follows the title page
directly. Word's computed layout of the untouched fixture is page-for-page
identical to Word's layout of the TOC-inserted save, so the toc-insert edit
changes Word's pagination not at all; it only disturbs the file enough to make
Word recompute. Our 23rd page is ours, and it is there before any edit.

**What that costs the rule tuned against it.** The `sectPr`-gated break-only
rule was calibrated so that the unedited document paginated to 23 and the
TOC-inserted one to 22. Under current-build truth Word says 22 for BOTH, so the
"unedited" row of the two-way table below is a dead target and the room-
conditional reading it forced is unsupported. The rule is re-derived under
"One break, one advance" further down; the table is kept because the retraction
is the point.

Its cause is worth recording because it is not a page-break rule. Pages 3 and 4
of the same render fill to within 17.1 and 12.8 px of the body bottom, so there
is no early ceiling; page 1 simply arrives at its foot 3.7 px low, leaving
32.9 px where the next block needs 34.2 (19.4 space-before plus a 14.8 line). We
miss by 1.3 px, spill two paragraphs onto a page of their own, and the explicit
page break then starts the heading a page later than Word does. The 3.7 px is
two discrete spacing steps on that page, +2.4 and +1.7, not accumulation.

Measure that kind of question in the BROWSER. Through `ApproxMeasurer` the same
page looks catastrophic — different line breaking and 15.6 px of drift — and all
of it is the approximate measurer rather than the renderer.

The +2.4 px step is **adjacent paragraphs carrying identical borders**. Of the
415 paragraphs in that document exactly five carry a `w:pBdr`, and exactly one
ADJACENT pair does — "Diluqofa H" followed by "Diluqofa F", both with
`<w:bottom w:val="single" w:sz="6" w:space="1"/>`. Their three unbordered
siblings above sit 15.3 px apart in both renders; the one bordered pair sits
15.3 px apart in Word's and 17.7 px in ours. The arithmetic is exact: `w:sz="6"`
is 0.75 pt, `w:space="1"` is 1 pt, and 1.75 pt is 2.4 CSS px. **Word treats a run
of identically bordered paragraphs as one bordered block, with no rule and no
space between them; we charge each paragraph its own border and space.**

The second step is **2.31 px, not 1.7**, and the quarter-line paragraphs have
nothing to do with it. That reading is retracted here rather than quietly
dropped, because it was reached by elimination — "the only unusual construct in
the stretch" — and elimination named the wrong thing.

`generate-quarterline-probe.mjs` measures the construct three ways. A paragraph
whose `w:lineRule="auto"` multiple is a quarter line costs 5.000 CSS px at 13 pt
in our render and 5.000 in Word's; two cost us 10.000 and Word 9.667; over a
12-line stack the accumulated difference is 0.07 px. Word's baseline-to-baseline
advance depends on where on the page the line sits — at 13 pt the first advance
is 15.000 pt and the rest 14.750, and at 26 pt that order reverses — so a
per-paragraph cost read off a two-line control is not a quantity Word has, and
only the accumulation over a stack means anything. The same probe excludes
`w:jc`, `w:tabs` and text length: Word gives all six shapes identical geometry.

The structural argument is shorter and would have saved the measurement. Both
quarter-line paragraphs sit in the first row of the signature table, and that row
carries `<w:trHeight w:hRule="exact" w:val="495"/>`. An exact row is exactly that
tall whatever it holds, so nothing inside it can move anything.

The real rule is **a cell's own borders come out of an exact row's height, and a
table's borders do not.** `generate-exactrow-probe.mjs` rebuilds the fixture's
three rows — exact 495, 110 and 1089 twips — and varies one thing at a time,
measuring `top(row 2 mark) - top(row 0 mark)` in CSS px:

    variant                       ours     Word     diff
    exact rows, tblBorders       46.34    46.03    +0.31
    atLeast rows, tblBorders     58.34    58.03    +0.31
    exact rows, no borders       46.34    46.03    +0.31
    atLeast, no borders          54.34    54.03    +0.31
    exact, middle row removed    39.00    38.67    +0.33
    fixture tcPr, borders+shd    46.34    44.03    +2.31
    fixture tcPr, borders only   46.34    44.03    +2.31
    fixture tcPr, shading only   46.34    46.03    +0.31

Every variant sits within the same 0.31 px of Word except the two carrying the
fixture's own `<w:tcBorders>`. Word draws a cell border INSIDE the exact row and
takes its width out of the content box: `w:sz="12"` is 1.5 pt, and Word's number
drops by exactly 2.00 CSS px. A `<w:tblBorders>` rule of the same weight in the
same visual position costs Word nothing, and `<w:shd>` is inert. We charge
neither, so an exact row whose cells carry their own borders is 1.5 pt too tall.

That reproduces the fixture to 0.02 px: from `TIV'W VIQIMUSIM` to `[TIV gece]`
Word measures 44.00 and we measure 46.33, against the probe's 44.03 and 46.34.

So `hRule="exact"` is handled correctly, table borders are handled correctly, and
neither rule that closes #38 is a line-spacing rule. Both are borders charged
where Word absorbs them — one between adjacent bordered paragraphs, one between
an exact row's content and its own cell border.

`wild2-med-phase23-protocol`, whose re-exported reference shows the same
one-extra-page shape, has **zero** bordered paragraphs and **zero** quarter-line
paragraphs. Its extra page is a different cause, and closing #38 will not close
it. That cause is now isolated, and it is a page-fit rule.

### An empty paragraph carrying only a page break always fits

Word lays phase23 in 69 pages and we lay 70, with the same content on pages 1 to
68. The whole difference is one fit decision at the foot of page 68. Our render
and Word's agree there line for line — the section 10.4 table ends at the same
place and the empty paragraph after it sits at the same place — and then Word
puts ONE MORE empty paragraph on the page and we do not. That paragraph's only
run is `<w:br w:type="page"/>`, so spilling it costs a whole page: our page 69
holds nothing but that paragraph, and its break starts the heading on page 70
where Word starts it on 69.

The paragraph authors no `w:spacing`, so it inherits this document's
`w:pPrDefault` — `before="200" after="200" line="276"` — which at 11 pt is a
19.4 px line with 13.3 px above and 13.3 px below. There was 22.1 px of room.

`scripts/generate-pagefit-probe.mjs` fills a page with exact-height paragraphs,
tunes a shim so an exact amount of room is left, and puts one paragraph there.
Four sweeps over rooms of 18 to 45 CSS px, each differing from the next by one
authored thing, and the room at which the paragraph stops spilling:

    target paragraph                    ours     Word
    text, no page break                  33       33
    text plus a page break               33       33
    EMPTY, only a page break             33    fits at every room tested
    empty, no page break                 33       33

Three of the four agree exactly, and they pin the ordinary rule: a paragraph
needs its space-before AND its line to fit, and does NOT need its space-after —
13.3 + 19.4 is 32.7, which is why the threshold sits between 30 and 33 on both
sides. Emptiness alone does not change it, and a page break alone does not
change it.

**An empty paragraph whose only run is a page break is the exception. Word puts
it on the current page at 18 px of room, where the ordinary test wants 32.7, and
starts the new page after it. We apply the ordinary test and spill it.** In the
fixture that is 22.1 px of room against our 32.7 px demand, so we spill, and the
spill costs the page.

Both this and the two #38 rules are the same shape of defect: a quantity Word
declines to charge at a boundary, charged in full by us.

Sweep S's floor is 18 px, and the threshold turns out to sit 0.1 px below it, so
the sweep read "fits at every room tested" as "always fits". It is not always;
see "One break, one advance" below for the demand Word actually charges.

**Sweep S under-determines the rule, and toc-insert is what separates it.** The
swept target carries no `w:sectPr`, so only ONE page advance is ever available
and "the paragraph always fits" and "the break is absorbed" put the marker on
the same page at every room. `wild2-legal-ca-agreement` carries the same
break-only paragraph WITH the section's first `w:sectPr` on it (document.xml
offset 58394), which makes two advances available and a blank page possible.
Measured in both directions against Word, on the same document and the same
paragraph:

    case                          room at P   Word   ordinary test   always fits
    unedited   (P carries sectPr)   plenty      23     23 correct      22 WRONG
    toc-insert (P carries sectPr)   none        22     23 WRONG        22 correct
    probe sweep S (no sectPr)       18..45    no blank  spills <33     correct

Word's unedited page 2 really is blank (header and footer only) and its
toc-insert page 2 really is `KACUCUJI A`; inserting the TOC fills the title page
and takes the room away. So neither rule alone is Word's. The rule consistent
with all three rows: a break-only paragraph never spills, AND its page break is
honoured only when its line fitted the room that was left — when the page was
already full the break coincides with the page end that is happening anyway and
is absorbed. With no `sectPr` both readings collapse to one advance, which is
exactly why the probe could not see the difference.

Read a sweep that pins a rule only as far as the advances the probe makes
available. Two settings of the room is one variable; a second construct that
adds an advance is another, and this rule needed both.

**Every row of that table is now retracted except the last.** The two
ca-agreement rows read a 23-page reference the July Word build computed and this
one does not, and the re-export above settles it at 22 for the unedited document
and 22 for the TOC-inserted one. With both rows at 22 the table discriminates
nothing, and the room-conditional rule it forced was tuned on a dead target. The
replacement is measured below.

### One break, one advance; and the break-only paragraph costs one bare line

`scripts/generate-sectadvance-probe.mjs` writes two documents that between them
separate every variable this rule was ever confused by. Each case fills a page,
tunes a shim so an exact room is left, puts a target paragraph there and then a
MARKER, and the marker's page minus the last filler's page is the number of page
advances taken. `probe-sectadvance.docx` puts a `w:sectPr` on every target and
sweeps the following section's start type; `probe-sectadvance-nosect.docx` is the
same sweep with no sections at all. Both Word exports reproduced themselves byte
for byte on every page, so the control the drift screen demands passes here too.

**The advance count is not the defect.** At room 200, where the target certainly
fits and nothing is confounded, Word and we agree on every shape:

    target                     following section   Word   ours
    empty, break only              nextPage           1      1
    empty, break only              continuous         1      1
    text + page break              nextPage           1      1
    text + page break              continuous         1      1
    text, no break                 nextPage           1      1
    text, no break                 continuous         0      0
    empty, no break                nextPage           1      1
    empty, no break                continuous         0      0

A page break followed by a `nextPage` section start is ONE advance, not two —
the break and the section start are the same page end — and our layout already
collapses them. `evenPage` and `oddPage` are excluded from that comparison on
purpose: their advance count also depends on the parity of the page a case
happens to land on, and once the two renders' page counts diverge the same case
sits on different-parity pages, so those rows do not compare across engines.

**The defect is the fit test, and it is one quantity.** The room at which each
target stops spilling, Word against us:

    target                          sectPr   Word fits   Word spills   ours fits   ours spills
    empty, break only, 10 pt          yes        17           16          33            27
    empty, break only, 20 pt          yes        33           32        (200)           40
    empty, break only, 10 pt          no         17           16        every           never
    empty, break only, 20 pt          no         33           32        every           never
    text + page break, 10 pt          yes        33           27          33            27
    text, no break,    10 pt          yes        33           27          33            27

Three readings fall straight out.

**Word's demand doubles with the font size, so it is a line height, not a
constant.** 10 pt brackets it in (16, 17] px and 20 pt in (32, 33]; the two
brackets intersect at 1.60 to 1.65 px per point. That is the SINGLE-SPACED line
— about 1.221 em for this theme's Calibri — and it excludes both alternatives.
The `w:line="276"` multiple would make it 18.72 px at 10 pt, and Word fits at 17
and 18. Adding `w:before="200"` would make it 32.05 px, which is exactly the
ordinary demand the text-carrying rows show. **Word charges an empty break-only
paragraph its bare line, and neither its space-before nor its line multiple.**

**Word does not care about the `w:sectPr`.** The no-section control gives the
identical 17/16 and 33/32 thresholds. A gate on `sectPr` has nothing to gate.

**Our two branches are both wrong, in opposite directions.** With a `sectPr` we
apply the ordinary test, so we spill from room 27 down where Word carries on to
16. With no `sectPr` we never spill at all, at any room or size. The gate does
not make one branch right; it picks which way to be wrong.

That also reconciles sweep S rather than contradicting it. Sweep S ran at 11 pt,
where the demand is 17.6 to 18.2 px, and its floor was 18. Word fitted at 18
because 18 is at or above the demand, and the sweep called that "always".

The rule to implement is one line: **to fit at the foot of a page, an empty
paragraph whose only run is `<w:br w:type="page"/>` demands its single-spaced
line height. Every other paragraph demands space-before plus its full line, as
now.** Same shape as the #38 rules and the exact-row rule — a quantity Word
declines to charge at a boundary, charged in full by us.

Pins for that change, all measured at engine `f88a63f`:

| pin | expected | now |
| --- | --- | --- |
| probe-sectadvance BK 10 pt | fit 17, spill 16 | fit 33, spill 27 |
| probe-sectadvance B2 20 pt | fit 33, spill 32 | fit 200, spill 40 |
| probe-sectadvance-nosect NS 10 pt | fit 17, spill 16 | never spills |
| probe-sectadvance-nosect N2 20 pt | fit 33, spill 32 | never spills |
| probe-sectadvance TB / TO 10 pt | fit 33, spill 27 | same — must not move |
| probe-pagefit sweep S | fits at 18..45 | same — must not move |
| wild2-legal-ca-agreement | 22 pages | 23 |
| wild2-med-nccih-protocol | 23 pages | same — must not move |
| wild2-med-phase23-protocol | 69 pages | same — must not move |

The two fixture pins that motivated the whole question resolve like this.
ca-agreement's break-only paragraph sits at **28.06 px** of room and demands
16.3, so Word fits it and takes one advance; we demand 32.05, spill it, and the
spill buys the blank page 2 that makes us 23. nccih's sits at **543 px** of room
and demands 17.9, so it fits under any rule anyone has proposed. The two
instances were never distinguished by the section type, the header, or the
compat mode. **One of them is at the knife edge and the other is nowhere near
it**, and a rule was tuned on the pair as though both were evidence.

### toc-insert renders no TOC, so its PASS means nothing

Measured at engine tip dd7a26e. **Our render of the TOC-inserted document is
identical to the unedited fixture** — not similar, identical. The gate's own
edited `web-1.png` is byte-for-byte the same file as its baseline `web-1.png`
in a run that PASSED, and page 1 diffs clean line for line.

Neither the harness nor the package is at fault:

  - `scripts/check-edited-render.mjs` reproduces `renderWebPages` exactly —
    fixture by URL in viewing mode, then `setInputFiles` the edited DOCX — and
    page count, page-1 character count and dot-leader presence are all
    unchanged. The SAME path with a different document moves the render from 23
    pages to 8, so viewing-mode uploads work.
  - The saved `document.xml` carries a proper TOC field (begin / instrText /
    separate, 12 `PAGEREF _Toc` entries, no `w:sdt`). Word lays it out: Word's
    edited page 1 has 90 text lines against the unedited 46, with dot leaders.

We write a correct TOC and lay it out as nothing. Until that is fixed:

1. **toc-insert's PASS is spurious.** The metric scored our TOC-less page 1
   against Word's TOC-bearing page 1 at 0.65% severity.
2. **The scenario is not evidence about pagination rules.** Our web side is
   effectively the unedited document throughout. Word lays the unedited file in
   23 pages and the edited one in 22; we drew 23 before the break-only rule
   (correct for the document we are actually drawing) and 22 after (wrong for
   it). The scenario "passes" only because our unedited-content render fell to
   22 and Word's edited render happens to be 22 as well.
3. **The baseline is the only uncontaminated number in it**, and it reports the
   break-only rule as a straight regression here: 23 vs 23 matching at 0.65%
   became 22 vs Word's 23 at 99.07% worst.

A scenario whose web side silently renders the UNEDITED document still produces
a severity score, a page count and a PASS. Compare the edited and baseline PNGs
before believing any scenario that exercises a field.

### The room under the break-only paragraph

`scripts/browser-page-room.mjs <fixture> <pages>` reports the room left at the
foot of a page's body. The break-only paragraph carrying the first `sectPr` is
empty, so it paints nothing and cannot be found by text; what can be measured is
the room it was offered. bodyBottom is calibrated from the deepest body item
anywhere in the document rather than assumed from the margins.

Unedited `wild2-legal-ca-agreement` page 1: last body item bottom 945.82,
deepest body item anywhere 958.55, margin bottom 960. **Room = 12.73 px** (14.18
against the margin). TIGHT, not plenty — which kills the reading that the
unedited case had room to spare and that the two states of this document
therefore discriminate the rule. They do not.

That paragraph's rPr is `sz=20` (10pt), NOT phase23's 11pt, so phase23's 32.7px
demand does not transfer: a 10pt line alone is ~13.3px, within a rounding of the
room. The fit here is marginal, and no rule should be tuned on it.

Word's unedited blank page 2 carries text at baseline 58.04 (header band, above
the 96px margin) and 1005.00 (footer band, below 960), and nothing between — no
body ink, no images. This does NOT establish that Word left the page empty
rather than spilling the paragraph onto it: an empty paragraph's mark is
non-printing, so both readings predict exactly this content stream.

### wild2-math-eq-as-images: the line box, not the image

`scripts/pdf-page-geometry.py <pdf> <page>` reads a Word PDF's content stream
and reports image boxes from the CTM that scales the unit image square, and
text positions from `Tm`/`Td`. The image boxes are PLACED boxes and compare
directly with the DOM; the text positions are BASELINES and do not.
`scripts/browser-page-images.mjs <fixture> <page>` reports the same page from
the browser.

Page 2 carries the first equations. Our image boxes match Word's within a
rounding pixel — heights 41.33/53.33/50.66/20.00/41.33/41.33/41.33 against
42.00/53.00/50.33/19.67/42.00/42.00/42.00, widths within 0.67, and every x
identical. So the images are neither mis-sized nor mis-indented. Their TOPS
run -10.3, -21.0, -21.0, -20.0, -31.0, -52.0, -72.7 px against Word.

Removing the constant 10.41 px ascent offset (measured on the two text-only
opening lines, which agree to 0.00) aligns our line tops with Word's baselines
and says exactly where the space goes:

    our top     +asc   Word base    delta     step  line
     102.63   113.04      113.04     0.00     0.00  text
     126.63   137.04      137.04     0.00     0.00  text
     169.63   180.04      192.73   -12.69   -12.69  (7)  IMAGE
     212.97   223.38      244.07   -20.69    -8.00  text
     236.97   247.38      268.07   -20.69     0.00  text
     ...        ...          ...      ...    ~0.3   six text lines, flat
     677.63   688.04      721.53   -33.49   -12.00  (2)  IMAGE
     721.63   732.04      775.56   -43.52   -10.03  text
     763.30   773.71      828.24   -54.53   -11.01  (9)  IMAGE
     815.30   825.71      900.93   -75.22   -20.69  (6)  IMAGE
     858.97   869.38      954.60   -85.22   -10.00  text

Every text-only step is zero to a third of a pixel. Every loss lands on a line
carrying an equation image, about 12 px each. Our text layout is exact and our
images are the right size: what is short is the LINE BOX that holds an inline
image. The gap above the first image is 13.70 px where Word leaves 24.04, and
the 10.34 px difference is the 10.41 px text ascent to within rounding — we
appear not to reserve the text ascent on a line an inline image dominates.

It compounds: by the foot of page 2 we are 85 px ahead, so we pull Word's first
two page-3 equations onto page 2 (nine images against Word's seven) and every
later page is shifted. That is the 51% mean over 7 of 8 pages — one per-line
defect, not a per-page one. Fix the image line box and re-measure before
reading anything else in this fixture.

Baselines are close to free. The corpus already holds a Word export of every
unedited fixture as `parity/<fixture>-word.pdf`, so no scenario needs a second
Word round trip, and the web baseline is rendered once per FIXTURE rather than
once per scenario. A fixture with no cached reference is skipped and says so
rather than triggering an export.

Word PDFs and rasters cache under the same Word container directory the
saved-DOCX gate uses, keyed by the DOCX package hash, so re-running a scenario
whose edit produced identical content costs no Word round trip. Every run
appends one JSON line to the tracked `parity/edit-roundtrip-history.jsonl`
recording the thresholds, the metric version, this repo's git SHA, and which
wordinweb build was measured. Changing the metric maths means bumping
`METRIC_VERSION` — older lines stop being comparable.

Add a scenario by appending one entry to `edit-roundtrip-scenarios.mjs`. Address
text by content rather than by pixel coordinate, assert that the edit landed,
and leave no pending tracked changes — the web half renders the saved bytes in
viewing mode, which shows the final document.

`parity/word-reference-manifest.json` pins the source package-content hash,
cached reference DOCX, Word-PDF hash, and page count for every fixture. If a
source fixture changes, refresh that fixture's Word reference intentionally and
update the manifest before running the candidate gate.

## Fixture safety

- `validate-docx.py` rejects malformed generated fixtures before Word opens them.
- `sanitize-docx.py` anonymizes a document while retaining its layout structure.
- `audit-fixtures.py` scans fixtures for identifying or sensitive information.

## Local fonts

- `extract-dfonts.py` extracts licensed Office fonts for local, git-ignored use.
- `extract-font-metrics.py OUTPUT` generates the engine's font metrics table at an explicit output path.

One-off probe generators, forensic readers, and local export experiments are
kept outside the public repository under `internal/scripts/`.

### A negative top margin is an absolute distance, and we read it as a signed one

`wild3-template-caed-pleading` is the corpus's only fixture with negative page
margins (`w:top="-1325"`, `w:bottom="-1267"`, `w:header="432"`). Against the
re-exported reference our body sits 160.25 CSS px too high on its single page —
a uniform translation, identical horizontal positions, no reflow. The header is
placed correctly: its line-number column starts at 86.64 px in our render and at
baseline 101.04 in Word's, which is the same place. Only the body origin moves.

`scripts/generate-negmargin-probe.mjs` gives each case its own section and its
own `w:pgMar`, puts a marker with an exact line height first in the body so the
marker's top IS the body top, and moves three things one at a time: `w:top`
through the negative range and past zero, the header's height (the fixture's own
`hRule="exact"` 14880 tw row against a 2880 tw one), and `w:header`. Both Word
exports reproduced themselves byte for byte on all 22 pages.

Body top in CSS px, at `w:header="432"`, for the tall header and the short one:

    w:top (tw)   w:top (px)   Word tall   Word short   ours (both)
      -2880        -192.00      192.20       192.20      -187.39
      -2160        -144.00      144.20       144.20      -139.39
      -1440         -96.00       96.18        96.18       -91.39
      -1325         -88.33       88.51        88.51       -83.73
       -720         -48.00       48.18        48.18       -43.39
       -360         -24.00       24.14        24.14       -19.39

Both engines are exactly linear and neither looks at the header at all — the
tall and short columns are identical to the digit, and repeating the fixture's
own row at `w:header="1440"` gives 88.51 again on both. Header height and header
distance are inert under a negative top margin, which disposes of the standing
guess that this position tracks the first-page header.

What is left is the sign. **Word puts the body top at `|w:top|` below the top of
the page; we put it at `w:top` itself, which is above the page.** Word's residual
against `|w:top|` is a constant +0.18 px, ours against `w:top` a constant
+4.61 px, and both slopes are exactly 1. The engine change is to take the
absolute value; the +4.61 px is a separate constant worth explaining but not
worth guessing at from here.

Above zero the header does govern, and there we disagree for a different reason:
with the short header Word puts the body top at 236.87 px (`w:header` 28.8 plus
the header's 208 px of content) and we put it at 267.59, and at `w:header="1440"`
the same 30.4 px gap appears (304.24 against 334.59). **We overcharge a header's
height by about 30.5 px.** That is not #67's defect and it does not affect this
fixture, whose margin is negative, but the probe measured it and it should not be
lost.

Correcting the sign moves our body 172.24 px down where the fixture needs
160.25, so **about 12.6 px will remain**. That residual is the leading run of
empty `BodyText` paragraphs measuring differently in the two renders, and it may
well be an artefact of our body currently starting 84 px above the page top.
Re-measure the fixture after the sign fix before reading anything into it.

### Which "wild2-legal"? All four, and the pages moved by one

Four engine comments cite a bare `wild2-legal`, which names two fixtures with
opposite provenance — `wild2-legal-ca-agreement`, whose reference the section
above replaced, and `wild2-legal-nih-contract`, which the drift screen passed on
all 419 pages. Every one of the four is **ca-agreement**, and the evidence is
structural rather than circumstantial.

The construct those comments are about is an empty paragraph immediately
followed by a table. ca-agreement's body **opens** with exactly that — an empty
paragraph, then the caption table — and nih-contract's opens with a text
paragraph, so `engine.ts:3095`'s "wild2-legal p1" can only be ca-agreement.
Mid-flow, ca-agreement carries the construct three more times, at body blocks
21, 249 and 345. Block 21 is the signature table on page 1 that #38 measured;
blocks 249 and 345 are the notices table and the signature-block table, and they
land on **old pages 15 and 23** — which is `engine.ts:3106`'s "p15/p23"
exactly. `engine.ts:3239`'s bullet routing cites "p3", and ca-agreement's old
page 3 is the `KACUCUJI A` bullet list while nih-contract's page 3 is contents
with dot leaders.

**None of those measurements has to be re-taken.** Rasterized at 192 DPI, the
old 23-page reference and the new 22-page one are byte identical page for page
either side of the dropped blank: old page 1 equals new page 1, and old pages 3
to 23 equal new pages 2 to 22. The stale part of that reference was one inserted
blank verso and nothing else, so the pixels at every cited page never moved.
What changes is only the numbering — **every cited page from 3 up drops by one**:

| site | as written | correct |
| --- | --- | --- |
| `engine.ts:3095` | wild2-legal p1 | wild2-legal-ca-agreement p1 |
| `engine.ts:3106` | wild2-legal's p15/p23 | wild2-legal-ca-agreement's p14/p22 |
| `engine.ts:3146` | wild2-legal's 2 x 13.8 | wild2-legal-ca-agreement's 2 x 13.8 |
| `engine.ts:3239` | phase23 + wild2-legal p3 | phase23 + wild2-legal-ca-agreement p2 |

A citation that names a page number is only as durable as the reference's page
count. Write the fixture's full name and expect to re-check the number whenever
a reference is re-exported.

### The two date-volatile references, frozen (#69)

`parity2-fields` and `probe3-field-switches` carry `DATE` and `TIME` fields, and
Word recomputes those when it opens a document. Their reference PDFs therefore
recorded the day they were exported — `7/8/2026` and `Saturday, July 11, 2026` —
so nobody could ever reproduce them, and the drift screen reported both as
differences every time it ran. A screen that always cries wolf twice teaches
people to stop reading it, which is the real cost; the mis-scoring was tiny
(0.0081% and 0.1769% of a page).

Both are now frozen with `w:fldLock`, which is the mechanism OOXML provides for
exactly this (ECMA-376 17.16.18, "Field Shall Not Be Recalculated").
`scripts/lock-volatile-fields.py <source> <destination>` sets it on every begin
`fldChar` whose own field code is `DATE` or `TIME`, byte for byte and nothing
else, so the copy stays a checkable function of its source. `CREATEDATE`,
`PAGE`, `NUMPAGES`, `SEQ`, `STYLEREF`, `REF`, `QUOTE` and `AUTHOR` are left
alone: they are either stable across exports or are meant to recompute.

Word honours it. The locked `parity2-fields` exports `1/15/2026` — the result
already cached in the package, which is what our renderer draws and what the old
reference never showed — and `probe3-field-switches` exports its cached
`Friday, July 10, 2026`, `2026-07-10`, `3:07 pm` and `15:07:42`. Two exports of
each reproduced byte for byte on both pages.

`parity/word-reference-docx/` now holds a third kind of derived package, and it
is worth being explicit that it is not like the other two. The content-type
repair and the pagination-hint strip are **glyph-neutral** — they change what
Word is willing to do, never what it paints. This one changes what Word paints,
deliberately. Its justification is different: our renderer never recomputes these
fields, so an unlocked reference was comparing Word's export-day layout against
our cached-result layout and calling the difference parity. Locking the fields
is what makes the two sides describe the same document.

The corpus fixtures themselves are untouched, so the web side still exercises an
unlocked `DATE` field reading its cached result.

### Word's bistable page, and what the gate does about it (#71)

`wild-doerfp` page 35 is the one place in this corpus where desktop Word does
not compute the same layout every time. Four exports of the one package on one
build produced two different page 35s, pages 1-34 and 36-38 byte-identical in
all four, and one of the two is byte-identical to the July reference.

**It is not per-export random.** Four fresh exports taken in one sitting are
byte-identical to each other on all 38 pages, and all four land on the side that
DIFFERS from the cached reference. So whatever decides it persists across
consecutive exports; it is some state Word carries, not a coin flipped per
document. That is worth knowing before anyone tries to reproduce it: a run of
identical exports does not clear this fixture.

The amplitude, by the saved-DOCX gate's own metric
(`abs(Rdiff)+abs(Gdiff)+abs(Bdiff) > 90` at 192 DPI), is **0.274111%** of page
35 — 9,448 pixels of 3,446,784, all of it 22 words moving horizontally by up to
0.554 pt with no vertical movement and no reflow.

`word-download-parity.mjs` now reads two optional keys per fixture from
`parity/word-reference-manifest.json`:

    "bistablePages": [35],
    "bistableCeilingPct": 1

A declared page is still rasterized, still measured, still written to
`results.json` with `"bistable": true`, and still shown on the report. What
changes is that it is kept out of the mean and the worst-page statistic, and
held instead to its own ceiling. The ceiling is 1%: 3.6x the measured flip, and
half the gate's own 2% worst-page threshold, so a genuine regression on page 35
fails — and fails sooner than the ordinary threshold would have. Tolerating a
known wobble is not the same as not looking at it.

**The edit round-trip gate is deliberately not changed.** `wild-doerfp` is not
in its scenario library, and that gate does not read the reference manifest at
all, so adding the same handling there would be machinery for a case that does
not exist yet. If a doerfp-class fixture is ever added to
`edit-roundtrip-scenarios.mjs`, its Word baseline comes from the same
`parity/<fixture>-word.pdf` and will inherit the same flip; the fix is to load
the manifest and reuse these two keys.

The first full gate run after this change should confirm that doerfp page 35
reads under 1%. The 0.274111% above is reference-against-a-fresh-Word-export;
the gate's number for that page also carries our own serialization's parity
error on top of it.

### An exact row clips nothing; it just stops at the paper edge (#56)

Two observations looked incompatible. `generate-exactrow-probe.mjs` put 8 plain
paragraphs in an `hRule="exact"` row and was read as showing BOTH engines
clipping them identically. Then a 14-paragraph TOC in a 260 tw row vanished
entirely from our render while Word painted about 90 lines out of it. #56 was
filed asking what separates the cases — field content, `cantSplit`, the row's
position on the page.

**Nothing separates them.** `scripts/generate-exactoverflow-probe.mjs` moves all
four candidates one at a time over 16 cases, and the two observations turn out
to be two different channels of one behaviour:

    case  authored   Word paints   ours paints   Word MARK-TOP   ours MARK-TOP
     A1          1             1             1          34.33           34.34
     A8          8             8             1          34.33           34.34
    A32         32            32             1          34.33           34.34
    A64         64            59             1          34.33           34.34
    A90         90            59             1          34.33           34.34
   A160        160            59             1          34.33           34.34
    F90    90 (TOC)           58             0          34.33           34.34
    C90  90 cantSplit         59             1          34.33           34.34
    P90   90 at foot          12             1         786.33          786.34

**The layout channel is already correct in both engines.** MARK sits 34.33 px
below TOP in Word and 34.34 in ours, in every case, whether the row holds one
paragraph or 160. The row contributes exactly its authored 260 tw to the flow
and nothing else, and that is what the earlier exact-row probe measured — it
read `top(row 2 mark) - top(row 0 mark)`, so it never looked at what was
painted, and "both engines clip identically" was never something it established.

**The paint channel is where we differ, and Word does not clip at all.** Word
lays the cell's content out from the row's top and paints it straight through
the bottom of the row box, over whatever follows, stopping only at the bottom
edge of the PAPER. The counts prove it arithmetically: at the top of a page the
row's content starts at 112 px and the sheet ends at 1056, which is 59 lines of
16 px, and Word paints exactly 59 for every authored amount from 64 up. Pushed
to the foot of the page the same 90 paragraphs start at 856 px and Word paints
exactly 12. The row never continues onto a second page — MARK is always on the
same page as TOP — so the rest is simply lost.

We paint one line: the number that fits inside the 17.33 px row box.

The other three candidates are inert. `cantSplit` changes nothing (C90 equals
A90 in both engines). Position changes only how much paper is left below.
And the TOC field is not special either — `F90` gets 58 where `A90` gets 59
purely because the field's begin paragraph consumes a line, and OUR F cases
paint zero for the same reason: the one line we allow is spent on the invisible
field-begin paragraph. That is the whole mechanism by which the inserted TOC
"rendered as nothing".

So the fix is one-sided and small: **keep the row's layout height at the exact
value, which is already right, and stop clipping the cell's painted content to
it.** The content should overflow the row box and be cut off by the page edge,
not by the row.

### math-eq's page-8 residual is a text-line deficit, not a pagination one (#62)

Partial diagnosis. The cause is narrowed to one page-filling difference and a
named suspect, but the suspect is NOT confirmed and should be probed before
anything is changed.

**The equations are already right.** Images per page are identical in both
renders — 2, 7, 9, 6, 6, 9, 2, 0 — so the VML extent fix did its job and no
equation crosses a page boundary differently. Whatever is left is text.

**We do not fill the page.** The section is A4 with a 1440 tw bottom margin, so
the body bottom is 1122.53 - 96 = **1026.53 px**. Word's page 7 runs its last
baseline to 1018.30, filling that almost exactly. Our page 7 stops at
**938.36**, and the deepest body item we place anywhere in the document is
**983.63** — so our usable page is short by at least 43 px and page 7 is 88 px
short of Word's fill.

The consequence is the whole residual. Word's page-7 tail (the paragraph
containing `guseqotu`) is on OUR page 8, and our page 8 therefore carries Word's
page-7 tail followed by Word's entire page-8 bibliography. Both renders still
end at 8 pages, so nothing about the page COUNT reveals it; the tail is simply
shifted, which is exactly the reported shape — page 8 structural at 27.34%,
page 7 at 6.23%, every other page at or below 1.69%.

Pages 1 to 6 stay clean because they are image-dominated with little running
text. Word's page 7 carries 94 text items. A per-text-line deficit is invisible
until a page is mostly text, and then it lands all at once.

**The suspect, unconfirmed:** this section carries
`<w:docGrid w:type="lines" w:linePitch="312"/>`. A `lines` grid makes Word snap
each line to a 312 tw (20.8 px) pitch, and our body text on this document
measures 20.67 px per line. Over the ~44 lines a full body holds, 0.13 px per
line is only 6 px, so the pitch alone does not obviously account for 43, and the
honest position is that the grid is a candidate and not a finding.

**Next measurement, and do this before changing anything.** Sweep `w:docGrid`
`type` (`default`, `lines`, `linesAndChars`) against `linePitch` over at least
two settings, on a page of plain text with a known line height, and read the
line count and the last baseline per page in Word. That separates three things
this measurement cannot: a line pitch we ignore, a body bottom we compute short
by a constant, and a per-line advance that is simply wrong. The 43 px figure is
a floor rather than a constant — it is our deepest item anywhere, not a measured
body bottom — so do not tune against it.
