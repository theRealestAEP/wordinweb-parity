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

> **CORRECTION (engine commit 5285045):** the "Word paints" column below
> (59/12/58) was a PDF-extraction artifact — the reader ignored Word's clip
> path (`re W* n` at the row box), reporting emitted text operators, not
> rendered ink. PyMuPDF (clip-honoring) and 192 DPI rasters agree: Word
> paints 1 line in the A/C/P cases and 0 in F, identical to our engine.
> Both engines clip exact rows; no fix was needed and none landed. The
> layout-channel rows (mark-to-mark 34.33px) remain valid.

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

### math-eq's page-8 residual is a page-BOTTOM deficit, not a per-line one (#62)

> **CORRECTION.** The section below reads the residual as a per-text-line
> page-fill deficit with a floor of "at least 43 px", derived from a deepest
> body item of 983.63. **Both numbers are wrong and the reading with them.**
> 983.63 came from a browser scan that filtered body items to `bottom <= 985`,
> so it could not report anything deeper than 985 whatever the render did; the
> same document scanned to 1035 reaches **1006.97** on page 5 and 1000.63 on
> page 4. The "43 px" was the gap between a nominal body bottom and that
> artefact, so it never measured anything.
>
> **Our line pitch on the page in question is IDENTICAL to Word's.** Measured
> against `parity/wild2-math-eq-as-images-word.pdf` with PyMuPDF, page 7 runs
> at 26.00 px through the body text in both engines, and the paragraph that
> spills runs at 41.67 px in both. There is no per-line deficit to find.
>
> What actually happens is one fit decision. Word's page 7 ends
> `Hunefigipawetagi` at 939.30 and then places BOTH lines of the following
> paragraph, at 959.99..980.66 and 1001.67..1022.33, against a nominal body
> bottom of 1026.53. We end `Hunefigipawetagi` at 938.36 — the same place to
> within a pixel — and move that whole two-line paragraph to page 8. At our own
> pitch its second line would land at 1000.72..1021.39, which is INSIDE the
> nominal bottom, so the ordinary test should have kept it. The paragraph is
> two lines with `w:widowControl` on, which makes it unsplittable, so rejecting
> the second line moves both.
>
> That is the same shape engine commit `355be56` recorded on
> `wild2-legal-ca-agreement` and deliberately left alone: "the ordinary test's
> effective bottom for a line can sit ~14px above the nominal 960". Here the
> deficit is bracketed rather than pinned — our render reaches 1006.97
> somewhere, and it refuses 1021.39 here, so the effective bottom lies in
> **[1006.97, 1021.39)** and the deficit is **at most 14.42 px**. Two
> independent documents, two independent measurements, one ~14 px quantity.
>
> **So #62 and ca-agreement's remaining 23rd page are one defect, and it is in
> `planBreaks`' effective bottom — `updateBottom`, `paragraphOverhang` or the
> note reserves — not in the line pitch, not in the docGrid, and not in the
> break-only rule.** Pin it with a room sweep on a document whose lines are a
> known height before changing anything.
>
> The docGrid sweep the section below asks for was run anyway, and it found two
> real and separate defects. Neither explains this page. See "What a docGrid
> actually does" below.

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

### What a docGrid actually does, and the two things we get wrong

`scripts/generate-docgrid-probe.mjs` writes `probe-docgrid.docx`: six sections,
each starting `nextPage`, each holding the same 60 single-line paragraphs on the
fixture's own A4 geometry with no header or footer, differing in one authored
thing. Read per case as the number of the last line on the section's first page,
that line's bottom, and the line pitch.

    case   grid                              Word pitch  ours   Word 1st top  ours    Word p1  ours
    N      (none)                                 16.00  16.32         96.57   96.52       57    57
    D312   linePitch=312, w:type omitted           16.00  16.32         96.57   96.52       57    57
    L312   type="lines"        pitch 312           20.67  20.67         98.90  184.52       44    40
    L240   type="lines"        pitch 240           32.00  16.32        104.57  160.52       29    53
    L480   type="lines"        pitch 480           32.00  32.00        104.57  240.52       29    25
    C312   type="linesAndChars" pitch 312          20.67  20.67         98.90  184.52       44    40

Word's rule is one line: **a `lines` or `linesAndChars` grid snaps each line's
advance UP to a whole number of grid rows; a grid with `w:type` omitted is
inert.** `linesAndChars` behaves exactly as `lines` for vertical geometry. Our
engine already treats the `default` grid as inert (N and D312 agree with Word to
0.32 px, which is just our line height), so only the two `lines` cases matter.

**1. We reserve four grid rows at a section opening and Word reserves none.**
`engine.ts:2216` adds `4 * docGridLinePitch` to `bodyTop` on the first page of a
section. Word's first line under a lines grid starts at 98.90 px against a body
top of 96 — it reserves 2.33 px, which is just the snap of the first line
itself. Ours starts at 184.52, 88.00 px lower, and the reserve tracks the pitch
exactly: 88.00 at pitch 312, 64.00 at 240, 144.00 at 480. The comment at that
site claims "Word reserves four grid rows at a section opening"; on a plain text
section Word reserves nothing, and this is the whole reason our L312 page holds
40 lines where Word's holds 44.

**2. Our snap is `max(natural, pitch)` where Word's is
`ceil(natural / pitch) * pitch`.** L240 is the case that separates them: the
grid row is 16.00 px and the natural line is 16.32, so Word takes TWO rows and
advances 32.00 while we fall back to the bare natural 16.32 and fit 53 lines
where Word fits 29. L312 and L480 agree only because the natural line is smaller
than one row there, where both formulas return one row. A sweep with a single
pitch cannot see this; it needs a pitch on each side of the natural line.

**The page-bottom test is NOT implicated.** In every case both engines stop at
the last line that fits under the nominal body bottom of 1026.53 — ours ends
L312 at 1011.34 with the next line needing 1032.01, Word ends at 1008.99 with
the next needing 1029.66. Both are correct; they differ only because defect 1
started ours 88 px lower. So this probe does not explain #62's page 7, which is
not a section opening and whose pitch already matches Word's.

### Verification at engine tip 2ca1765

`npm run parity:edit-roundtrip` is **17/17**, measured against
`wordinweb 0.2.5 — likeoffice @ 2ca1765f8219` with both engine locations agreeing.
Every scenario reads severity mean 0.000% and worst 0.000% except `header-footer`
(0.400% / 0.800%, its long-standing rasterization noise) and `toc-insert`
(0.011% / 0.260%). **`field-update` is 0.000% / 0.000% over 23 pages on BOTH the
edited and the baseline halves**, so the continuous-section and space-before
fixes hold at tip.

**Browser page counts, which are what decide a pagefit landing.** The break-only
change (`355be56`) moved two fixtures headlessly and flagged the browser as the
decider. Measured in the real renderer at this tip:

| fixture | Word | headless before | headless after | BROWSER at tip |
| --- | --- | --- | --- | --- |
| wild-gatech | 22 | 23 | 24 | **22** |
| wild-wirfp | 20 | 19 | 20 | **20** |
| wild2-legal-ca-agreement | 22 | 23 | 23 | **23** |
| wild2-math-eq-as-images | 8 | — | — | **8** |

Both movers land on Word's own count in the browser, including `wild-gatech`,
which the ApproxMeasurer put at 24 — one page WORSE than before the change. A
headless delta is not evidence about a page count; this is the second time that
has been shown on this rule and it should not need showing a third.

**ca-agreement is still 23, and the break-only rule is no longer why.** The room
at the foot of its page 1 is now **28.06 px**, measured by
`scripts/browser-page-room.mjs` at this tip, against a demand of 16.3 — so the
paragraph fits and the rule fires correctly. The earlier reading of **12.73 px**
in "The room under the break-only paragraph" above was taken before the two #38
border fixes and the exact-row fix landed, and those reclaimed the difference;
the number is stale, not wrong-headed, and 28.06 is what current code offers.

Our page 2 is nevertheless a blank verso: its only body item is one empty
14.77 px line at the body top, which is the break-only paragraph itself, and the
"Qeliwecap Ca. ____, with ____." at y=48 and "A-2 / jow 8-54-09" at y=990 are the
running header and footer, not body ink. So the paragraph still spills — at 28.06
px of room against a 16.3 px demand — which means the fit test is rejecting it on
the EFFECTIVE bottom, not on the demand. That is the ~14 px deficit `355be56`
deferred, and it is the same one #62 runs into. **The `wild2-legal-ca-agreement =
22 pages` pin is unmet at this tip, and closing it means closing that deficit.**

The gate's own `toc-insert` numbers say the same thing from the other side: the
edited comparison passes at 23 pages against Word's 23, while the BASELINE reads
**99.080% worst** — our unedited 23 against the re-exported 22-page reference,
one page out of step from page 2 on. The baseline is the uncontaminated half
here, and it is the one that fails.

### The line chart's residual is a legend reservation, and 163% is not a bug (#64)

Two separate questions, and the second one has to be settled first or it
poisons the first.

**The >100% semantic-text figure is the metric's own scale, not an anomaly.**
`textWeightErrorPct` is `200 * Σ|webMass - wordMass| / (wordTotal + webTotal)` —
a symmetric difference normalized by the MEAN of the two masses, so its range is
**0 to 200%**, and `parity-metric.mjs` self-checks exactly that
(`symmetricWeightError(2, 1) === 200/3`). Every semantic weight channel uses the
same scale, including the `weight 2.95%` on the same page. So `163.12%` means
the two text layers are very nearly DISJOINT: solving `200|a-b|/(a+b) = 163.12`
puts the smaller side at about a tenth of the larger, i.e. roughly 90% of the
text mass has no counterpart in its paired region. It corroborates the
displacement rather than contradicting it, and nothing needs fixing in the
metric. Read any semantic weight number against 200, not 100.

**The geometry is a legend width reservation.** The chart SPACE agrees exactly —
both engines put it at 134.99, 96.00 to 614.99, 384.00, so the graphicFrame
extent is not in question. The plot rectangle inside it:

    edge          Word      ours     ours - Word
    left        170.20    175.80          +5.60
    right       510.61    531.70         +21.09
    top         150.40    151.84          +1.44
    bottom      350.71    352.16          +1.45
    width       340.41    355.90         +15.49
    height      200.31    200.32          +0.01

**The value axis is CLOSED**, which retires the "~7% larger along the value axis"
finding this file records from the column and bar pages: the height now agrees to
0.01 px and the whole vertical position to 1.45. What is left is horizontal, and
it is not symmetric — we are 5.60 px too far right on the left edge and 21.09 px
too far right on the right edge.

The right edge is the real one, and the legend is what sits there. Word's legend
text runs `Alpha` at x=568.47..601.88 and `Beta` at 568.47..594.88, with its keys
just left of that, so Word leaves **34.4 px between the plot's right edge and the
legend block**. Ours leaves 18.6 px: our keys start at 550.32 where Word's plot
has already stopped at 510.61. **We under-reserve the legend's horizontal band
and spend the difference on plot width.** The 4.24 px alignment figure in the
gate is that displacement seen through the metric, and it is a real one.

Measure this from the Word PDF's gridlines rather than from a raster: the seven
horizontal gridlines and the axis line give the plot rectangle directly, and
`fitz`'s `get_drawings()` reports them with their 1.0 pt width.

### A nil cell border costs nothing, but only when both sides say so (#51)

`scripts/generate-nilborder-probe.mjs` varies one thing at a time on the row
0 / row 1 boundary of a three-row `atLeast` table whose rows are governed by
their content, and reports `top(row 1 mark) - top(row 0 mark)` — a pure layout
distance that never looks at paint. The table rule is `w:sz="12"`, which is
1.5 pt, or **2.00 CSS px**.

    case        authored at the boundary          Word    ours    ours - Word
    D-norule    no tblBorders at all             15.33   15.35          +0.02
    A-none      tblBorders only                  17.33   17.34          +0.01
    B-nilboth   BOTH cells declare nil           15.33   17.34          +2.01
    C-nilone    only row 0's cells declare nil   17.37   17.35          -0.02
    E-own12     row 0 cells bottom single sz=12  17.33   17.34          +0.01
    F-own24     row 0 cells bottom single sz=24  19.33   19.34          +0.01

Word's model reads straight off the controls. `D` fixes the bare content height
at 15.33, `A` adds exactly 2.00 for one rule, and `F` adds exactly 4.00 for a
`w:sz="24"` (3 pt) border — so **what a boundary costs is the width of the border
actually drawn there, and it scales with `w:sz` rather than being a constant**.
`E` shows a cell border that merely restates the rule at the same width adds
nothing.

**`B` is the finding: a nil on BOTH sides of the shared boundary returns the row
to its no-rule height exactly, 15.33 against D's 15.33. Word charges zero.** And
`C` is the other half: a nil on only ONE side does NOT suppress, and Word charges
the rule in full. A shared boundary needs both cells to decline it.

**We agree with Word on five cases out of six and get `B` wrong by exactly one
rule width.** That is `rowBorderWidths` taking `max(tblBorders.insideH,
cellBottom, cellTop)` and reading `w:val="nil"` as "no opinion" rather than as a
zero that overrides — while `paintCellEdges` reads it correctly and draws no
rule. So wherever a nil cell border overrides a table rule we are 2.00 px too
tall on that boundary and paint nothing there. `C` passing in both engines says
the one-sided case is already right and must stay right: the fix is not "nil
means zero" but "nil on both sides means zero".

**Still unprobed:** the sided-ness half of #51 — a one-sided UNSHARED cell border
in an exact row, where half-share, own-top and own-bottom all agree on current
evidence. This probe does not separate them, and one variant would.

### We do not fall back at all on lo-provenance, and that is the asymmetry (#68)

#58 recorded `probe3-lo-provenance` as the corpus's only proven behavioural Word
build change: its `SourceText` style names "Liberation Mono", "which neither
build resolves", the July export fell back to Courier New and current Word falls
back to Calibri. The question left open was what OUR fallback is, so it could be
aligned. **There is nothing to align: we never fall back.**

`~/Library/Fonts` holds LiberationMono, LiberationSans and LiberationSerif, so
the browser resolves all three and paints the real faces. Our computed styles on
that fixture are `"Liberation Serif", sans-serif` (108 runs), `"Liberation
Mono", sans-serif` (7 runs) and `"Liberation Sans", sans-serif` (3), and a canvas
measurement confirms the Mono request lands on a monospace face — 288.05 px for
`MMMMiiiill` at 48 px, against Calibri's 230.25.

Word substitutes **all three**, and the reference PDF names exactly what it used:

| authored | Word used | metric-compatible? |
| --- | --- | --- |
| Liberation Serif | Times New Roman | yes, by design |
| Liberation Sans | Arial | yes, by design |
| Liberation Mono | **Calibri** | **no** — its twin is Courier New |

The Liberation family exists to be metrically compatible with those three
Microsoft faces, so the Serif and Sans substitutions move almost nothing, and
that is why only the Mono run showed up in the drift screen. The July build's
Courier New was the CORRECT metric twin; current Word swapping it for a
proportional Calibri is a substitution regression on Microsoft's side.

**So the 1.380% is a measurement-environment asymmetry, not an engine defect: we
paint the font the document asks for and Word paints a substitute, because Word
cannot see fonts installed in the user's `~/Library/Fonts`.** Three independent
substitutions in one file are the evidence that it sees none of them.

**Do not "align" this by teaching the engine that Liberation Mono means
Calibri.** That would reproduce a Word bug rather than Word's layout, it would
pin us to one Word build (July's answer was a different font), and it would make
the document render worse for any user who actually has the font. If the 1.380%
needs to stop being reported, the honest fixes are to annotate the manifest entry
as an environment artefact, or to stop the fixture depending on a font only one
side can see — not to copy the substitution.

Worth generalizing: a reference is only comparable if BOTH renderers resolve the
same faces. Installing a font locally silently changes what our side paints and
nothing about what Word's cached references did.

### Word does not quantize a rule's width; we do, on the three lightest weights (#19)

`scripts/generate-rulewidth-probe.mjs` puts one paragraph per border weight —
`w:sz` 2, 4, 6, 8, 12, 18, 24, 36, 48, i.e. 0.25 to 6.00 pt — each carrying only
a `w:space="0"` bottom border, so the width question is not confounded with a
placement one.

**Word paints `sz/8` points faithfully at every one of the nine weights**, as a
filled rectangle rather than a stroke, ratio 1.000x throughout (the 6 pt row
reads 6.025 from the rect extraction, which is the extraction's rounding). So
there is no Word-side quantization to match, and the whole error is ours.

`renderEdge` snaps the painted width to a whole DEVICE pixel —
`Math.max(1/dpr, Math.round(declaredWidth * dpr) / dpr)`. Measured in the
browser at both scales, six weights are exact and three are not:

    w:sz   authored   in px    ours @1x   ours @2x    error @2x
       2     0.25 pt  0.3333      0.5000     0.5000       +50.0%
       4     0.50 pt  0.6667      0.5000     0.5000       -25.0%
       6     0.75 pt  1.0000      1.0000     1.0000            0
       8     1.00 pt  1.3333      1.0000     1.5000       +12.5%
      12     1.50 pt  2.0000      2.0000     2.0000            0
      18+          —       —       exact      exact            0

Two things follow that the backlog note did not have.

**The error is concentrated on the two commonest weights.** `sz=4` is Word's
default table rule and `sz=8` the next most used, and they are exactly the two
that miss. Total ink mass is preserved under antialiasing, so the painted-width
ratio IS the ink ratio, and the metric's symmetric weight error follows directly:
28.58% for an all-`sz4` grid. The 27.5% measured on `staging-tblextreme` with
positions already matching is that number, which closes the loop on where the
residual comes from.

**`sz=8` changes SIGN with the display.** At 1x we paint 1.000 px against 1.333
(-25.0%, a 28.57% weight error); at 2x we paint 1.500 (+12.5%, 11.77%). So a
non-retina user sees a different rule weight than the parity gate measures, and
any calibration read at one scale does not transfer to the other.

**The snap buys crispness the reference does not have.** The comment at that site
justifies it by Word's 0.5 pt rule being "one physical pixel" at 2x. It is not:
0.5 pt is 1.333 device px at 2x, and the reference raster at 192 DPI is
0.5/72*192 = 1.333 device px too — antialiased across two rows, exactly the
unsnapped case. **Matching Word means not snapping the width.**

Recommended change, NOT made here: drop the width snap and keep the position
handling, which is separate and already correct (`placeExact`/`applyFrac` carry
the fractional offset in a transform). Predicted effect: every weight goes to
1.000x, `tableRuleWeightErrorPct` on rule-heavy fixtures falls from ~27.5% to
about nothing, and the 1x/2x sign flip disappears. Keep a floor only where a
declared width would otherwise round to zero — Word paints 0.25 pt faithfully,
so faint is the correct appearance and the current `1/dpr` floor is what makes
`sz=2` 50% too heavy.

Landing it needs a full gate run, because it moves ink on every bordered fixture
in the corpus at once.

### The header overcharge is the 22.5 pt table clearance, not the trailing space (#72)

`scripts/generate-headerheight-probe.mjs` settles which component of
`measureHeaderFooter` puts the body ~30.5 px too low above a zero top margin.
Six sections, each `w:top="0"` so the HEADER governs the body top, each with its
own header part, varying one authored thing. Header lines are
`w:line="240" w:lineRule="exact"` — 12 pt, 16.00 CSS px — and TB's row is an
exact 240 tw, so TB and P1 carry the SAME content height and differ only in
being a table.

    case  header                              Word     ours   ours - Word
    P1    one paragraph                      64.64    64.59         -0.05
    P2    two paragraphs                     80.64    80.59         -0.05
    P3    three paragraphs                   96.64    96.59         -0.05
    S2    one paragraph, w:after=200 (10pt)  77.97    77.92         -0.05
    S4    one paragraph, w:after=400 (20pt)  91.31    91.25         -0.06
    TB    one table row, exact 240 tw        64.64    94.59        +29.95

**Word's rule is simply `headerDistance + header height`, with the trailing
space-after included and no clearance of any kind.** Line count scales it exactly
16.00 px per line (P1→P2→P3), the trailing space-after enters at exactly its
authored value (+13.33 for 10 pt, +26.67 for 20 pt), and **a table header costs
Word exactly what a paragraph header of the same content height costs** — TB and
P1 are the same number to the digit.

**Five of the six agree with us to 0.05 px, and only TB diverges.** The named
suspect for #72 — the header's trailing space-after — is **refuted**: S2 and S4
match on both sides, so we already charge that quantity correctly. Line count is
right too.

The whole overcharge is the last term of `measureHeaderFooter`:

    Math.max(height, contentBottom) +
      (!hasOnlyUnwrappedAnchors && complexHeader ? ptToPx(22.5) : 0)

`ptToPx(22.5)` is 30.00 px, `complexHeader` is true when the header holds a table
or a positioned shape, and the measured divergence is 29.95 — which is 30.00 once
the -0.05 every row shares is taken out. **For a TABLE header Word reserves
nothing, so that clearance should not apply to the table half of
`complexHeader`.**

That also explains the negmargin probe's number: its header is a single-row table
with an exact `trHeight`, so the clearance fires there by construction, which is
why the overcharge showed up as ~30.4-30.7 px in a measurement that never varied
the header's composition.

**Scope the fix to what was measured.** This probe moves the TABLE case only. The
positioned-shape half of `complexHeader`, and the `hasOnlyUnwrappedAnchors`
carve-out that already exists beside it, are NOT tested here — the 22.5 pt may
well be right for an anchored shape, and #67's pleading-rail finding is the
reason that carve-out exists at all. Vary an anchored shape the same way before
removing the term outright.

### An exact row's cell border is charged to its OWN side, and we already do (#51b)

The sided-ness half of #51, left open because every case the exact-row probe
measured carried borders on both edges, where all three candidate models predict
the same total. An exact row's HEIGHT is fixed by definition, so only where the
content sits INSIDE the row can separate them.

`scripts/generate-sidedness-probe.mjs` puts a marker immediately before a single
`hRule="exact"` row and another inside it, with no `w:tblBorders` so only the
cell border under test is in play, and reports `top(MK) - top(REF)`:

    case  cell borders (sz=12 = 1.5pt = 2.00px)   Word    ours    vs N
    N     none                                   16.00   16.00   +0.00
    T     top only                               18.00   18.00   +2.00
    B     bottom only                            16.00   16.00   +0.00
    TB    top and bottom                         18.00   18.00   +2.00

    model        predicted T   predicted B   predicted TB
    half-share         +1.00         +1.00          +2.00
    own-side           +2.00          0.00          +2.00
    own-bottom          0.00          0.00           0.00

**`own-side` is Word's model and the other two are refuted.** The content box
loses the FULL border width at the edge the border is on: a top border pushes the
content down by its whole 2.00 px, a bottom border takes its width off the bottom
and does not move the content top at all, and `TB` matches `T` because only the
top edge bears on this measurement.

**Our engine already implements it, matching Word to 0.00 px on all four cases.**
No change is needed, and the earlier "half-share vs own-top vs own-bottom all
agree on current evidence" is now resolved rather than merely still-ambiguous.

### Word reserves nothing above ANY header, anchored shape included (#80 scope)

`scripts/generate-headeranchor-probe.mjs` exercises the branch the header-height
probe could not. That probe's table header has `anchors.length === 0`, so it
reaches the 22.5 pt clearance through `complexHeader`'s SECOND disjunct with
`hasOnlyUnwrappedAnchors` false by vacuity — the anchored-shape path was never
touched, and the `hasOnlyUnwrappedAnchors` carve-out exists because of #67's
pleading-paper rails, so 22.5 pt might have been right there.

Same geometry as the header-height probe, so the numbers compare directly. The
shape is 1 inch by 8 pt anchored at the paragraph's own top, ending ABOVE the
12 pt line's bottom so it cannot raise `contentBottom` — the only quantity in
play is the flat clearance itself.

    case  header                            Word     ours
    PB    plain paragraph, no anchor       64.64    64.59
    AN    anchor, wrapNone                 64.64    64.59
    AS    anchor, wrapSquare               64.64    64.59
    AT    anchor, wrapTopAndBottom         64.64    64.59

**Word reserves nothing above any of them.** Table (from #72), and all three wrap
settings here — every header composition measured gives the identical body top.
The flat 22.5 pt has no case left among those tested that justifies it.

**The probe genuinely exercised our anchor path**, which matters because a null
result is otherwise indistinguishable from a probe that missed. Our render paints
the shape on pages 2-4 and not on page 1: a 96.0 x 10.7 px black box at y=48,
exactly the authored 1 inch by 8 pt. So the shape is parsed, and we still apply
no clearance.

**That confirms #80's scope rather than widening it: the only OBSERVABLE defect
is the table disjunct.** Our anchor cases already agree with Word at all three
wrap settings, so removing the term changes the table case and nothing else in
practice.

**One sub-question left for whoever makes the change.** It is not established WHY
our wrapped-anchor cases decline the clearance, since the code as written should
add it for `wrapSquare`. The likely explanation is that the parsed shape carries
no `wrap`, which makes `anchors.every((s) => s.wrap === undefined || ...)` — and
so `hasOnlyUnwrappedAnchors` — vacuously true. If that is a parse gap rather than
a decision, a document whose shapes DO carry a parsed wrap could still hit the
clearance, so confirm it before assuming the anchor path is inert.

### Verification at likeoffice-finalcal (#80 landed, #62 closed, footer clamp refuted)

Measured in the browser against the worktree build, engine
`likeoffice-finalcal @ 4cdcba9`.

**#80 landed and both header probes now agree with Word.** The 22.5 pt clearance
is gone from `measureHeaderFooter`. `TB` moves 94.59 -> 64.59, and all ten cases
across the two probes sit -0.05 to -0.06 px from Word — the same constant every
case shares:

    probe-headerheight   P1 -0.05  P2 -0.05  P3 -0.05  S2 -0.05  S4 -0.06  TB -0.05
    probe-headeranchor   PB -0.05  AN -0.05  AS -0.05  AT -0.05

The anchor cases did not move, as predicted: `layoutFrame` consumes the header's
anchors before `collectAnchors` runs, so that half of the term was unreachable
and only the table disjunct ever fired. Full edit-roundtrip gate 17/17 with every
number identical to the pre-change run — no gate scenario has a table header.
`wild3-template-caed-pleading` cannot move (its `w:top` is -1325, and the header
height reaches `bodyTop` only inside `if (sp.marginTop >= 0)`); measured after at
30.90%, which is #67's known ~12.6 px leftover.

**#62 is closed.** The grid-snap demand fix returns the whole paragraph to
page 7, and our two lines land within ~1 px of Word's:

    line              Word              ours
    guseqotu 1    959.99 .. 980.66   959.36 .. 980.03
    guseqotu 2   1001.67 ..1022.33  1000.69 ..1021.36

Page 8 now opens on `B pohaladuh` in both renders. math-eq is 8 pages, Word's own
count. (The earlier reading that our page 7 held only the first line was a
measurement artefact: the scan filtered to `bottom <= 1010` and the second line
ends at 1021.36.)

**ca-agreement is still 23 against Word's 22, and the footer-clamp suspect is
REFUTED.** `engine.ts` names the `bodyBottom` footer clamp as the leading
suspect, needing a footer of ~62 px to put the bottom at 945.8. Measured per page
in the browser — `w:titlePg` means page 1 uses a different footer from the rest,
so this has to be read per page, not once:

    page   footer painted   implied footerH   implied bodyBottom
       1   996.22..1008.03            11.78              960.00
       2   990.34..1008.06            17.66              960.00
       3   990.34..1008.06            17.66              960.00
       4   990.34..1008.06            17.66              960.00
      20   971.94..1007.98            36.06              960.00

**Every page clamps to 960.00, and page 1 — the page that matters — has the
SHORTEST footer of all at 11.78 px.** The clamp `min(960, 1008 - footerH)` needs
`footerH > 48` to bite and the tallest anywhere is 36.06. So the browser has no
footer deficit either, and the browser/headless asymmetry the suspect rested on
does not exist.

That leaves ca-agreement's remaining page where eq-as-images' half turned out to
be: **in the DEMAND, not the bottom.** Its break-only paragraph sits at 930.49
with a 14.77 px line and a 960 bottom — 29.51 px of room for a bare-line demand
of 14.77 — and still spills, which no bottom in the table above can explain. The
paragraph is a textbook instance of the shape (`w:pPr` carrying only `w:rPr` and
the section's `w:sectPr`, one run whose only content is `<w:br w:type="page"/>`),
so it should take `pageBreakOnlyDemand`. Whether it does is the next thing to
measure, and it wants engine instrumentation rather than another browser probe:
if it is taking the ordinary demand instead, space-before plus line is 34.17 px
against 29.51 and the spill is explained exactly.

### ca-agreement's last page was a phantom line in the FLOW (#75, closed)

The instrumentation the section above asks for was written, and it exonerates
the break-only rule completely. `isPageBreakOnlyParagraph` ACCEPTS the shape,
`pageBreakOnlyDemand` fires in the browser exactly as it does headlessly, and
the demand is the bare 15.33 px line the rule intends. Nothing in the flag path,
the `pPr` `rPr` or the one-line precondition is wrong.

**The paragraph spilled because the cursor reached it at y=948.89 against a 960
bottom — 11.11 px of room, not 29.51.** That is the whole correction, and it
retracts the premise rather than the rule.

**29.51 px was measured below the last PAINTED item, and on that page our paint
and our flow disagreed by 18.40 px.** `browser-page-room.mjs` reports the room
under the deepest thing on the page, which is the right question for a paragraph
that paints something and the wrong one for a paragraph that paints nothing:
between the last painted item at 930.49 and the cursor at 948.89 sat four empty
paragraphs' worth of flow that no scan of the raster can see.

The 18.40 px is one rule. `layoutBlock` charged a document-opening empty
paragraph TWO mark lines when a table followed, and `applyOpeningFlowOverlap`
then lifted the whole painted first-page body back up by exactly the same line
height. So the reservation only ever moved the FLOW — invisibly, by construction
— and it moved it just far enough to spill the break-only paragraph and buy the
blank verso.

Word charges ONE line, and the current-build reference says so to 0.02 px. Its
letterhead table's first row is an `hRule="exact"` 260 tw = 17.33 px, and the
rule under that row sits at 131.71, so Word's table top is 114.38 against a body
top of 96 — an 18.38 px opener, where our mark line measures 18.40.

The two-line reading was PDF-measured on the 23-page export that #58 replaced,
so it is one more casualty of that reference and not a new defect. The
grown-header half of the rule (phase23's 2 x (13.4 line + 6 after)) keeps its
own evidence and is untouched; only the table disjunct and its paint-side
counterpart go.

**Scope was checked rather than assumed.** Exactly ONE corpus fixture opens with
an empty paragraph immediately followed by a table, and it is ca-agreement.

At engine `9a6b058`: ca-agreement is **22 pages** in the browser, the break-only
paragraph sits at 930.494 with 29.51 px of genuine room against its 15.33 px
demand, and the full edit-roundtrip gate is 17/17. **toc-insert's baseline half
goes from 99.080% worst to 0.650%**, its web side landing on 22 pages against
Word's 22 for the first time. Its edited half is 23/23 with 12 TOC entries on
each side, so that scenario is no longer the contaminated case this file warns
about.

The general lesson is the one the `word-reacted` section already makes in a
different key: **a number read off the raster describes the paint, and
pagination is decided by the flow.** When the two can disagree — and any rule
that reserves space without painting it makes them disagree — only engine
instrumentation settles which one a fit test saw.

### A rule's painted width lives in a transform, because Chromium snaps boxes (#19/#79)

Word paints `w:sz/8` points faithfully at every weight, so there is no Word-side
quantization to match and the whole error was ours. THREE sites produced it,
where the backlog named one:

1. the width snap in `renderEdge`, `max(1/dpr, round(w*dpr)/dpr)`;
2. the `deviceHairline` PAINT branch beside it, which hard-coded a 1 px strip
   with `scaleY(.5)` for every rule under 0.75 px — that, not the snap, is what
   painted `sz=2` and `sz=4`, so dropping the snap alone would have left Word's
   DEFAULT table rule untouched;
3. **Chromium snaps a painted box's SIZE to a whole CSS pixel.** Writing the
   declared width into `style.height` makes `sz=2` THREE times too heavy and
   leaves `sz=8` exactly as wrong as before. The old `deviceHairline` branch was
   working around this and never said so.

**Measure a rule as INK MASS off the raster, never as a DOM box.** The box lies
in both directions: `getBoundingClientRect` reported the snapped 0.5 px rule as
0.5, and reported a fractional height as 0.667 on a box that painted 1.0. Sum
`1 - luminance` down a column through the rule instead. Painted ink against
Word's `sz/8` pt, at pinned device scale:

    w:sz     before 1x   before 2x    after 1x   after 2x
       2     unpainted   unpainted       1.000      0.994
       4         0.753       0.753       1.000      1.003
       6         1.000       1.000       1.000      1.000
       8         0.750       0.750       1.000      1.002
      12..48     1.000       1.000       1.000      1.000

The 1x/2x sign flip is gone. The fix is to carry the width as `scaleY`/`scaleX`
on a 1 px box, the same trick the hairline branch used, applied at every weight;
position handling is separate and unchanged.

**`sz=2` was not painted at all, before or after, and a DOM measurement cannot
see that either.** `sameParagraphBorders` compared `width`, which floors at
0.75 px and therefore SATURATES below `sz=6`, reporting 0.25 pt and 0.5 pt rules
as the same edge. The probe's `sz2` and `sz4` paragraphs merged into one bordered
block and the shared boundary was suppressed. Harmless while both weights snapped
to one painted width; live the moment each is painted at its own. Any probe that
sweeps a property should check that every case in the sweep actually rendered.

Corpus spot-check, `tableRuleWeightErrorPct` against the same references:

    staging-tblextreme p1   27.52% -> 5.85%
    staging-tblextreme p2   25.66% -> 4.21%
    parity-tables p1        22.30% -> 2.48%
    coverletter-anon p1     39.24% -> 0.53%
    forsale p1              48.34% -> 22.81%
    benchmark p1            10.68% -> 12.75%

`severityPct` is unchanged on every fixture, and the three no-light-rules
controls (parity-dividers, parity2-dropcap, probe2-styleref-headers) do not move
at all. The full edit-roundtrip gate stays 17/17, and **toc-insert's edited half
goes to mean 0.000% / worst 0.000% over its 23 pages**, from 0.011% / 0.260% —
the ink these three sites were losing was the last thing left in it. **benchmark p1's rule channel RISING while its ink goes from 0.75x to
1.0x means our rule SET differs from Word's on that page** — more ink makes a
wrong set worse — so that is a separate defect this change reveals rather than
causes. forsale carries `w:sz="0"` borders, the one case with no Word
measurement behind it; they keep the old one-device-pixel floor.

### The chart legend sits still; the plot's right edge does not (#81)

`LEGEND_GAP = 16` in `chart-geometry.ts` was fitted to the LINE page of
probe-charts-basic, and the bar page contradicted it by ~13 px. Re-reading every
page of `parity/probe-charts-basic-word.pdf` with `fitz get_drawings()` — one
method for all five, chart-local CSS px, plot rect taken from the white plot fill
and its gridlines — shows the line page is the outlier and the bar pin was right:

    page      plot left   plot right   legend key   key w   legend text
    column        35.23       401.44       422.77    7.32        433.52
    line          35.21       375.63       413.68    8.00        433.48
    pie               —            —       439.13    7.32             —
    bar           37.86       394.31       422.77    7.32        433.50
    area          35.21       392.99       422.76    7.32        433.49

**The legend does not move.** Four of the five pages put the key at 422.77 with a
7.32 px swatch and the text at 433.5. Only the LINE page differs, and only in the
key: 413.68, 8.00 wide, because a line chart's key is a line-and-marker sample
rather than a plain swatch. The legend TEXT's left edge is invariant to 0.04 px
across every page that has one.

**So there is no constant clearance to widen or narrow.** The plot-right-to-key
gap is 21.33 (column), 28.46 (bar), 29.77 (area) and 38.05 (line) — four
different numbers for one alleged constant. `LEGEND_GAP` generalises the single
page whose key geometry is atypical, and the right edge is set by something other
than the legend's position. What sets it per chart type is NOT established here
and should be measured before any constant is re-fitted.

The old bar figure (264.8 pt plot width, commit 6669f9e) re-reads as 267.34 pt —
2.54 pt wider, which is the extraction offset that comment already admitted. The
bar page never contradicted anything; the method it was measured with did.

### A lines-grid section's opening paragraph is an ordinary paragraph (#82)

`generate-gridopen-probe.mjs` sweeps six openers on one
`w:docGrid w:type="lines" w:linePitch="312"`, varying only the first paragraph
of each section and then running the same plain body under it. Both Word
exports reproduce byte for byte on all 12 pages. Word's `L01` top, against the
plain case, beside our engine before and after:

    case                                    Word     ours before   ours after
    P   plain, no before                   +0.00           +0.00        +0.00
    S   plain, w:before=12pt              +16.00           +0.00       +16.00
    G   snapToGrid=0, no before            -2.33          +37.00        -4.67
    GS  snapToGrid=0, w:before=12pt       +13.67          +11.33       +11.33
    H   Heading1                           +0.00          +23.59        +0.92
    HG  Heading1 + snapToGrid=0            -2.33          +36.92        -4.75

**Three rules lived in `docGridDropBefore` and Word refutes all three.** A
`snapToGrid="0"` opener without a space-before dropped TWO grid rows — Word
drops none, and its -2.33 is simply the first line's grid snap not being taken,
which is what turning the grid off means. A `Heading1` opener took a grid row
plus 1.5pt of "grid leading" — Word puts it exactly where a plain paragraph
goes. And every other opener had its space-before DROPPED — Word applies it in
full, on top of the snap.

**The third was never suspected, and it is the same gap that produced the other
two.** `probe-docgrid` authors `w:before="0"` in all six of its cases, so it
could not see a space-before being dropped, and the change that read its result
took the silence for licence. A sweep pins a rule only as far as the values it
varies — the same lesson the sectPr sweep taught about advances.

In ABSOLUTE terms the three opt-out cases now land on Word to 0.13 px. What
remains is a constant +2.29 px in the first line's grid snap on the SNAPPED
cases and a further +0.92 on Heading1, both pre-existing.

**Still open, measured here and deliberately not fixed:** `snapToGrid="0"`
should also suppress the PER-LINE grid snap. Word runs those paragraphs at
their natural 18.33 px advance and we still snap them to 20.67. That is the
`minLineHeight` argument to `breakParagraph`, threaded through six call sites
including two lookahead simulations, so it wants its own change and its own
gate run.

### A both-nil boundary charges nothing inside an exact row too (#51, closed)

The one variant e186c1e's restriction left open. An exact row's HEIGHT cannot
answer the question — it is the authored value whatever the borders say — so
`generate-exactnil-probe.mjs` measures the content INSET instead: two
`hRule="exact"` 495 tw rows sharing one `tblBorders insideH` sz-12 rule (1.5pt =
2.00 CSS px), reporting `top(MK in row 2) - top(UP in row 1)`. Row 1's exact
height is 33.00 px, so anything above that is the boundary's cost.

    case  authored at the shared boundary      Word     ours   ours - Word
    N     no rule at all                      33.03    33.00         -0.03
    R     insideH sz=12, no nil               35.00    34.00         -1.00
    RN    insideH sz=12, BOTH cells nil       33.00    34.00         +1.00
    RO    insideH sz=12, only the LOWER nil   35.00    34.00         -1.00
    RU    insideH sz=12, only the UPPER nil   35.00    34.00         -1.00

**Word's #51 rule governs the inset exactly as it governs an atLeast row's
height.** `R` costs the full 2.00 px, `RN` costs zero, and a one-sided nil
suppresses nothing — `RO` and `RU` agree to the digit, so the rule is symmetric
and does not care which side declares it.

**We get all four rule cases wrong, and identically.** Our number is 34.00
everywhere: we charge HALF the rule where Word charges the whole of it to the
row below the boundary, and we do not read `nil` here at all. Note that
`exactRowCellBorderShare` only inspects `tcBorders`, so a `tblBorders insideH`
rule never reaches it — the 1.00 px comes from elsewhere and the fix wants that
path found before anything is changed.

### parity-hftemplates p2-p4: the reference is clean and the score is a regression (#89)

The 96.88% on p2 (192.72% semantic text, with p4 at 15.52% and p3 at 4.09%)
carried the classic stale-reference signature, so the screening discipline was
applied first: two fresh exports of the fixture on the current Word build are
byte identical to each other AND byte identical to the cached July 16 reference
on all four pages at 192 DPI. The reference needs no re-export, and the score
is real.

It is also new. Every full run from 18 to 30 July (engine `e2f94ad`) read this
fixture 0/0/0/0, and the 6 August run at engine `22451c2` reads 0/96.88/4.09/
15.52 — the regression landed in that window, which contains #80's removal of
the 22.5pt complex-header clearance.

The asymmetry, measured from the fresh Word PDF against our render: on pages
2-4 our BODY TOP sits 29.0 / 23.0 / 30.5 CSS px above Word's, and page 1
agrees exactly. Page 1's header is a plain paragraph; the other three are the
template headers — Ion Light (a table), Banded (a wrapSquare anchored bar),
Ion Dark (a wrapTopAndBottom anchored bar). On p2 the arithmetic closes to
0.1pt: Word's body top is `headerDistance 35.4 + table row 36.0 + the header's
trailing empty paragraph ~22.5 = 93.9pt`, above the 72pt top margin, while we
stop at `35.4 + 36.0 = 71.4pt` — under the margin, so the margin governs and
we sit 22.5pt high. **We drop the paragraph that follows the header's table**;
the old 22.5pt clearance was accidentally standing in for it, which is why #80
regressed this fixture and why the two header probes (whose table header has
no trailing paragraph) measured the clearance as pure overcharge. Both
readings are right about their own documents: the probes under-determined the
rule, exactly the way `probe-docgrid`'s `w:before="0"` hid the space-before
drop. The anchored-shape pages miss by 23.0 and 30.5px and their decomposition
is not pinned here — only that Word charges header height we do not.

### The dropped paragraph was the after-table collapse, applied outside cells (#95)

The mechanism behind p2 was already in the engine and already correct — for
cells. `layoutFrame` collapses the mandatory empty `<w:p/>` OOXML places after
a table to zero height, a rule calibrated on parity2-nestedtables (the
trailing `<w:p/>` after the L3 and L2 nested tables, inside a cell), and it
applied that collapse in EVERY frame context: cells, headers, footers, text
boxes alike. Ion Light's header is `tbl` + one empty pPr-less `w:p`, exactly
the collapse's trigger, so the header measured 22.5pt short. Word's two rules
were never one rule: in a cell the trailing paragraph collapses, in a header
story it is charged in full, and #89's arithmetic closing to 0.1pt with the
paragraph charged is the header-side evidence. No new probe was needed — both
sides of the distinction already had a Word-verified fixture.

Engine branch `header-trailing-para` (5bef1dd) gates the collapse on the
`inCell` flag cell layout already passes. Measured against that build:
p2 96.88% -> 0.00% (clean, matching p1), and every sentinel holds digit for
digit — probe-headerheight P1/P2/P3/S2/S4/TB at 64.59/80.59/96.59/77.92/
91.25/64.59, probe-headeranchor PB/AN/AS/AT all 64.59, benchmark
0.00/0.37/0.35/4.00, staging-tblextreme 0.00/0.00. So the fix is the trailing
paragraph, and no clearance came back.

p3/p4 stay at exactly 4.09% and 15.52% — a separate mechanism, still open.
What is measurable from the cached reference: calibrating text ascent on the
now-agreeing p2 (14.34px), Word's body top sits 119.03px on p3 against our
96.03 (we stop at the 72pt margin), and 181.39px on p4 against our 151.69
(we do charge the shape's height, 29.70px short of Word's charge). Neither
page closes to header arithmetic from the shape's extent, its page-offset
bottom, or the Normal paragraph advance; the Banded/IonDark pStyles the
fixture names are not defined in its styles.xml, so both resolve to Normal
and style spacing is excluded. #80's headeranchor probe cannot see this —
its shape ends above the line bottom, and these bars extend to page-relative
bottoms of 56.9pt (wrapSquare) and 96.0pt (wrapTopAndBottom). Pinning it
needs a probe that varies wrap mode and the shape's extent below its line,
measured against a fresh Word export.

### benchmark p1's 12.75% rule channel is raster quantization, not a rule defect (#83)

The rule SETS match. Word's page-1 vectors hold 9 horizontal + 7 vertical
table rules (two tables), every one 0.5pt, plus two 0.75pt text-border rules;
our render paints the same 16 rules at the same positions (a single ~0.6pt
whole-page shift, within the metric's global offset) and the same lengths.
Nothing is missing, nothing is extra, and no boundary class differs — so
there is no probe to write and no engine change to make.

What the channel measures is the REFERENCE raster. pdftoppm rounds both edges
of a thin filled rectangle to the nearest device pixel, so an identical 0.5pt
rule (1.333 device px at 192 DPI) lands as one or two whole device rows purely
by where it falls on the pixel grid. Predicted from each rule's vector edges by
`round(edge x 8/3)`, all 14 table rules match the raster exactly — 7 rules
read 2 rows (0.75pt of ink) and 7 read 1 row (0.375pt). Our render antialiases
the true 1.333, so every rule differs from the reference by -25% or +33% in
mass, and the channel sums that to 12.75%. The pre-#79 snap painted 1.0 device
px, which agreed with the majority phase by luck — that is the whole story of
10.68% rising to 12.75% when our ink went from 0.75x to 1.0x of Word's.

Read the rule-weight channel with this floor in mind wherever the authored
weight is under ~1pt: for sub-pixel rules it cannot reach zero for any correct
renderer. Set membership and position questions go to the PDF vectors, not the
raster.

### The chart legend right edge is fixed in pt, and the per-type gap is a label overhang (#84)

`scripts/generate-legendedge-probe.mjs` writes `fixtures-staging/
probe-legendedge.docx`: bar, line and pie, each with the standard right legend
at 360x216pt AND at 240x144pt (the wp:extent is the only thing the small cases
change). Both Word exports reproduce byte for byte on all six pages, and
`parity/probe-legendedge-word.pdf` is the reference. Everything below is read
from the PDF vectors with fitz `get_drawings()`, chart-local pt.

**Every right-edge quantity is identical at the two box sizes**, so the whole
band is fixed in pt at Word's default 10pt chart text and nothing on this edge
is a fraction of the box — which retires `PLOT_EDGE_INSET`'s share of it, and
the probe's left/top/bottom edges are size-invariant too (plot left 28.40/
28.40 on bar, 26.41/26.41 on line), so the fraction fitted in 6669f9e was a
one-size accident throughout.

The band, right to left: the legend label column's right edge sits 9.87pt off
the chart's right edge; labels are left-aligned at `chartRight - 9.87 -
maxLabelWidth`; the key sits 8.05pt left of the label column when it is a
swatch (5.49pt square + 2.56 gap) and 21.42pt when it is a line sample (19.2pt
segment + 2.2 gap). Then the plot's right edge:

    page      key left   plot right   gap     16pt + half last bottom label
    bar  L      317.07       295.74   21.33   16 + 5.30 ("12")  = 21.30
    bar  S      197.07       175.74   21.33   16 + 5.30         = 21.30
    line L      303.71       281.72   21.99   16 + 6.31 ("Q4")  = 22.31
    line S      183.71       161.72   21.99   16 + 6.31         = 22.31
    column (probe-charts-basic)       16.00   16 + 0            = 16.00
    area   (probe-charts-basic)       22.33   16 + 6.31         = 22.31

**The rule: the plot stops 16pt short of the legend key, plus half the width
of the bottom-axis label that centres on the plot's right edge.** A bar
chart's last value tick and a line or area chart's last category centre on
that edge and overhang it; a column chart centres its categories inside their
bands and overhangs nothing. Three of the four types close to 0.03pt and line
to 0.32. (#81's line-page gap of 38.05px was measured to the marker inside the
line sample, not the sample's left edge; re-anchored, that page reads the same
21.99pt as this probe.)

The pie has no plot rect. Its circle takes the vertical leftover (diameter
166.92pt in the 216pt box, 93.77 in the 144pt one, both against a ~40.7pt
title band and ~10.5pt bottom pad) and centres horizontally in the band the
legend leaves, with ~4.7pt of clearance folded in: centre x = (keyLeft -
4.7)/2 at both sizes.

Engine branch `legend-edge` (worktree, commits 2c516a4 and 7b43d65)
implements the right-edge rule, the measured band for right legends, and the
pie centring: plot right = legendLeft - 16pt - overhang, replacing the tick
allowance, the fractional inset and the flat LEGEND_GAP on that edge only.
Measured back in the browser, gridline right edges land on Word within
0.10-0.99pt at BOTH sizes on all four axis pages, the pie centre within
0.9-1.4pt at both sizes, and probe-charts-basic's line page falls 30.03% ->
4.78% with column/bar/area flat or better. The full edit round-trip gate is
17/17 on the branch build.

Do not read this probe's per-page severity as the rule's score: line-S READS
27.60% after the fix against 19.01% before, while its own gridlines land on
Word. The page is dense, its VERTICAL band is still wrong (our line-S plot
runs 39.75..121.88pt against Word's 40.80..119.03) and so is the left gutter
(ours 30.75 against Word's 26.41), and correcting the width rescaled the
series against those standing offsets. Both leftovers are size-invariant in
pt, neither is decomposed by a probe yet, and the small-box pages stay noisy
until they are.

### SmartArt we author must survive Word's re-evaluation (#94)

Word does not paint the cached `dsp:drawing` a SmartArt package carries; it
re-runs the `layoutDef` over the data model and paints what THAT produces. So
authored SmartArt has two descriptions of the same art, and they have to agree.
Ours did not, and the whole 65.66% on `word-interop-smartart-only` was that
disagreement: the engine cached a cycle family's ellipse ring in explicit
srgbClr colors while its layoutDef said `alg lin` + `shape roundRect` +
autofit, and its colorsDef said `schemeClr accent1` — which, in a package with
NO theme part, Word resolves against its own default theme. Word painted a row
of three #156082 roundRects with 24pt Aptos where our render showed the cached
ring. Real-Word-authored SmartArt was never affected (its two descriptions
agree by construction), which is why `wild2-med-phase23-protocol` p13 sat at
0.08% severity throughout.

Engine `b67b22a` (branch `smartart-consistency`) makes the cycle family
self-consistent, and two Word exports calibrated it:

- **Everything pinnable was pinned, and Word honours every pin exactly.**
  Child w/h as `refType` fractions (0.25*W, 0.3*H → 90.00x54.00pt measured),
  the cycle `diam` (0.6*H → circle radius 54.00pt measured), `primFontSz`
  val=12 with no autofit rule, and 12pt bold white Calibri written as explicit
  `a:rPr` on the data-model runs (Word paints Calibri-Bold 12.00 over its
  theme font). The cached node stroke is 1.5pt, matching what the styleDef's
  `lnRef idx="2"` makes Word draw.
- **`meth="cycle"` interpolates; `meth="repeat"` is the discrete one.** The
  first calibration export came back with nodes 2 and 3 blended to #75C38E.
  With `meth="repeat"` and the engine's six-color table as explicit srgbClr
  values, Word paints #4472C4/#ED7D31/#70AD47 per node, exactly the cache.
- **Word centers the arrangement's bounding box, not the circle.** For three
  children the ring's bbox is vertically asymmetric (top -81pt, bottom +54pt
  around the circle centre), and Word puts the BBOX centre at the canvas
  centre — the circle centre lands 13.5pt low. The cache formula reproduces
  this; measured agreement is 0.00-0.01pt on every ellipse edge.
- **Connectors are dropped, one-sidedly.** The layoutDef declares no
  connector nodes and the data model has no transition points, so Word draws
  none; the cache now draws none either.

Both references are fresh Word exports and both passed the self-reproduction
control (two exports, byte identical on the page). The fixtures were
regenerated with that engine build, the manifest is updated, and
`word-interop-smartart-embedded` — previously without a reference at all —
now has one:

| fixture | before | after |
| --- | --- | --- |
| word-interop-smartart-only | 65.66% severity | 0.21% |
| word-interop-smartart-embedded | (no reference) | 0.00% |

The other three families (list, process, hierarchy) still carry the shared
lin/roundRect layoutDef and remain inconsistent with their caches in exactly
the way cycle was; the explicit-color and pinned-text halves of this change
already apply to them, but their geometry is not calibrated. That is the
remaining phase of #94.

### staging-eastasian's 38% was the compat gate on the grid snap, not the CJK (#tail-A)

The reference screening discipline came first, and it matters here because the
repo held TWO copies: `parity/staging-eastasian-word.pdf` (July 16, the one
the corpus gate reads) reproduces byte for byte against two fresh exports of
the fixture; `fixtures-staging/staging-eastasian-word.pdf` (July 14) was a
stale sibling carrying a phantom empty heading line that pushed every band
64px down, and is now replaced by the clean copy. Diagnose against the
`parity/` copy only.

Measured by INK BANDS (raster row profiles - the DOM baselines of our flex
spans measure span bottoms, not baselines, and had invented two intra-line
"defects" the ink shows do not exist), all 14 bands matched in height and
every Normal<->Normal boundary agreed to 0.5px. The four boundaries touching
a Heading1 line missed by 7.5-13px each, summing to a 26.5px spread: Word
gives the heading line (Calibri Light 16pt, natural 26.04px) TWO grid rows
of the pitch-360 grid - 48px, glyph box centered - while we laid natural x
multiplier = 28.1px.

`generate-docgrid15-probe.mjs` -> `fixtures-staging/probe-docgrid15.docx`
(reference `parity/probe-docgrid15-word.pdf`, two exports byte-identical)
separates what staging-eastasian could not: at compat 15, in-paragraph
advances read directly per case:

    case   font                 pitch   natural   Word advance
    A360   Calibri 11pt          360      17.9      26.00  (= quantQuarterPt(24 x 1.0792))
    B360   Calibri Light 16pt    360      26.04     48.03  (2 rows)
    D360   Calibri Light 18pt    360      29.3      48.03  (2 rows)
    E360   MS Mincho 11pt        360       -        26.00
    F360   MS Mincho 16pt        360       -        48.03  (2 rows - EA snaps too)
    H360   Heading1 (2 lines)    360      26.04     48.03  (+ follower gap closes to 0.03px)
    A240   Calibri 11pt          240      17.9      32.00  (2 rows of 16)
    B240   Calibri Light 16pt    240      26.04     32.00  (2 rows of 16)

So the engine's `compatibilityMode < 15` gate on textSnap had NO discriminating
evidence behind it: it was calibrated on the fixture's Chinese-fallback lines,
where snap (2 x 24 = 48) and multiplier x natural (44.48 x 1.0792 = 48.0) are
numerically identical. probe-docgrid and probe-gridopen are both compat 12, so
the compat-15 side had never been probed. H360 also confirms the section-break
space-before rule (36.71 = 31.37 + max(0, before 16 - prev after 10.67))
composing with the snap.

Engine `tail-fixes` 821eac5 removes the gate for horizontal flow.
staging-eastasian p1: **38.01% -> 0.00% structural**. Deliberately NOT
implemented, filed as measured asymmetries:

- **EA oversized lines at compat 15** (F360): Word snaps them; our substituted
  Hiragino/PingFang raw profiles overstate the natural (ja 11pt raw box
  24.10px > pitch 24) and would false-snap lines Word lays at one row. The
  corpus EA lines land right through the auto path (zh baselines +0.09px), so
  the EA carve-out stays until the profiles carry Word-em snap metrics.
- **Vertical (tbRl) flow keeps the old gate**: enabling the snap there moved
  probe2-ruby-vertical p2 from 0.02% to 12.05%; Word's vertical column pitch
  does not take this rule, and nothing further is probed about it.
- The residual after the fix is a uniform-ish +0.5..+7px drift (band spread
  6.5px, worst at H2/combine where the PingFang re-sync interacts), inside
  0.00% structural.

Sentinels at 821eac5, digit for digit against the corpus log:
probe2-ruby-vertical 0.15/0.02, wild2-math-eq-as-images 0.00/1.69/1.23/1.25/
1.29/1.14/0.65/0.90, benchmark 0.00/0.37/0.35/4.00, staging-tblextreme
0.00/0.00. Core suite green, edit round-trip 17/17.

Tools this added: `internal/scripts/ink-bands.py` / `ink-bands2.py` (x-windowed
raster row-band comparison, web screenshot @2x vs reference @192dpi) and
`internal/scripts/capture-page.mjs`. Measure vertical drift questions with ink
bands, not DOM rects: `display:flex` span boxes bottom-align their glyphs, so
a span's rect bottom is not a baseline and cross-font comparisons built on it
are artifacts.

### caed-pleading's last 12.6px was w:suppressTopSpacing, not the empties (#tail-B)

Measured with the x-windowed ink bands: the pleading rail is EXACT (28 of 28
line numbers at 0.0 to -0.5px, both vertical rules within 0.75px), the body
top is Word + 0.4 (probe-negmargin re-measured at tip: markers 192.59/144.59/
96.59/88.92/48.59/24.59 against Word's 192.20/144.20/96.18/88.51/48.18/24.14),
and the WHOLE residual was the body text sitting +12.5..13.0px low from the
Court heading down - fully formed before Court's line, constant after it.

`generate-emptyexact-probe.mjs` -> `fixtures-staging/probe-emptyexact.docx`
(every export self-reproducing) went through two WRONG models before the flag
was noticed, and both retractions are worth keeping:

- "Word discounts ~12.7px across a run of empty exact paragraphs" - E6/E12
  refute it: the empty exact-480 run slopes at exactly 32.00px per paragraph.
- "Word top-anchors an undersized exact line at 0.8 x font size" - M9/M18
  (9pt and 18pt markers land at the SAME baseline as 12pt) and X24/X48
  (exact-360 and exact-720 bands land within 0.3px of each other) refute
  both the size-scaling and any height-fraction anchor.

What separates the fixture from pleading-anon (same 12pt-in-exact-480
construct, 0.00% all seven pages under the UNCHANGED engine) is one settings
flag: `<w:suppressTopSpacing/>` in w:compat. With it, Word charges the FIRST
line of a page min(exact, natural): the fixture's opening empty paragraph
costs 19.32px instead of 32, and every number on the page then closes -
Court's baselines, the caption cells' exact-16 anchoring (baseline 12.6-12.7
into the line, cell bottom rule at 730.56), and the caption table top at
567.83 with the Court style's spacing-after 660tw = 44.0px charged in full.
P12 shows the collapse at a POSITIVE top margin; Q6 shows an exact line
UNDER its natural stays put (the suppression only shrinks).

Engine `tail-fixes` cebe80d parses the flag and collapses the page-top exact
line. **wild3-template-caed-pleading p1: 29.34% -> 0.00% structural.** Open,
measured: our Arial 12pt natural is 18.40px where Word's collapsed line
measures 19.32 (baseline pairs only constrain the sum of first-line height
and the exact-line anchor, so the split is taken from our own metrics); the
break plan does not mirror the collapse (only this fixture and
probe-negmargin carry the flag, neither within 13px of a fit decision); the
flag's space-before half is unprobed and unimplemented. Sentinels:
pleading-anon 0.00 x7, benchmark/tblextreme/eastasian unchanged,
probe-negmargin markers unchanged on every pinned case. Core green, gate
17/17.

### A header bar pushes the body only as far as Word's wrap rules say (#97)

`generate-headeranchor2-probe.mjs` -> `fixtures-staging/probe-headeranchor2.docx`
(reference `parity/probe-headeranchor2-word.pdf`, exports byte-identical)
varies what the #80 probe could not: wrap mode x shape extent BELOW the
header's last line, paragraph-positioned like #80 so the numbers compare.
Word's marker (body top + 12.07 to its baseline):

    case                       Word base   shape bottom
    S8   square,  8pt             76.71    above the line
    N40  none,   40pt             76.71    101.97
    N72  none,   72pt             76.71    133.97
    S40  square, 40pt             76.71    101.97
    S72  square, 72pt             76.71    133.97
    T40  topAndBottom, 40pt      114.04    101.97  <- body top = shape bottom
    T72  topAndBottom, 72pt      156.73    144.64  (offset scales 1:1, both pitches)

**wrapNone and wrapSquare never move the body top, however far the bar hangs
below the header text. wrapTopAndBottom puts the body top exactly at the
shape's bottom edge (+distB).** And the carrier paragraph is exempt from its
own paragraph-positioned anchor: HDT40/HDT72 paint at the header top beside
the band.

parity-hftemplates p3/p4 are the same rules plus one more: their bars are
PAGE-positioned, and Word lays the single carrier paragraph's own line BELOW
the bar - p3 (Banded, full-width wrapSquare bar, behindDoc="1"): body top =
bar bottom 88.45 + line + after = 119.03, closing to 0.6px; p4 (Ion Dark,
wrapTopAndBottom): bar bottom 127.23 + distB 24 + line + after = 181.39,
closing to 0.2px. So the header story wraps like body text, except that a
shape anchored to a paragraph's OWN position does not displace that
paragraph.

Engine `tail-fixes` (this commit): frame stories bound their line breaker on
registered floats (the cell path generalized, page-origin-shifted);
page/margin-positioned wrapped anchors pre-emit so the carrier itself wraps;
behindDoc no longer suppresses a frame float when the anchor declares a wrap;
art (txbx-less wps) shapes carry wrap+dist and register floats in FRAME
stories only; and contentBottom counts only topAndBottom float bands (S72
pins that a square bar hanging below the text reserves nothing).

**parity-hftemplates p3 4.09% -> 0.00%, p4 15.52% -> 0.00%** (p1/p2 hold at
0.00). All seven probe cases land on Word (S/N at 64.59, T40 101.92 vs
101.97, T72 144.59 vs 144.67). The ten #80 sentinels are digit for digit:
probe-headerheight 64.59/80.59/96.59/77.92/91.25/64.59, probe-headeranchor
64.59 x4. wild2-med-phase23-protocol (margin-anchored header art, 69 pages)
is identical to the corpus log page for page; staging-anchors2,
parity2-textboxes, parity-wrapmodes, parity-headerfooter, staging-hf2,
benchmark, staging-tblextreme all at their corpus values. Core suite green,
edit round-trip 17/17.

### Two pre-existing findings from this wave, not caused by it

- **wild3-template-us-courts-answer p1 reads 24.94% against the corpus log's
  0.08%** at the UNTOUCHED likeoffice tip 8dba4e9 (verified by scoring on the
  main checkout build before any tail-fixes change). The regression landed in
  the #94/#95 window (3b49cc8..8dba4e9). Unattributed further; whoever owns
  that merge window should bisect it.
- The header probes (probe-headerheight, probe-headeranchor) and several
  other fixtures-staging probes were never copied into
  `apps/demo/public/fixtures`, so any browser measurement of them silently
  timed out on a 404. They are copied now (probe-headerheight,
  probe-headeranchor, probe-headeranchor2, probe-emptyexact,
  probe-docgrid15). A `.dxw-page` timeout on a probe fixture means CHECK THE
  URL FIRST - the demo serves only what that directory holds.

### The us-courts p1 regression was never the nil boundary: a pre-15 exact row charges its BOTTOM cell margin (#86 follow-up)

wild3-template-us-courts-answer p1 went 0.08% -> 24.94% in the engine window
`22451c2..72435fd`. The briefed suspect — #95's in-cell scoping of the
after-table empty-paragraph collapse — is excluded twice over: `5bef1dd` is not
an ancestor of the first regressed engine, and re-widening the collapse to all
contexts on tip reproduces 24.94 digit for digit (the fixture's after-table
`<w:p>` carries a pPr, which the collapse never matched anyway). Bisecting the
three-commit window pins `72435fd` (#86, the exact-row rule arithmetic);
`414f9da` still reads 0.08. The fixture is compatibilityMode 11, and every
probe behind #86 was compat 15.

Five probes, every package exported twice with ink-identical results:

- `generate-exactnil11-probe.mjs` — probe-exactnil's five cases in the same
  ca-agreement package with the ONE variable `compatibilityMode` rewritten
  15 -> 11, plus exact/content, content/content and full-border (D) variants.
  Verdicts identical to compat 15: a live insideH sz-12 insets the row below
  2.00 px, a both-nil boundary insets zero, one-sided nil suppresses nothing.
  END-REF is CONSTANT across N/R/RN/D1/D2 — an exact row's FLOW takes no rule
  charge at all, in either compat regime.
- `generate-exactnil11p-probe.mjs` — the same thirteen cases in the PLEADING's
  own package (its Word-95-era compat flag pile, styles, fonts). Every number
  identical to the ca-agreement-based probe: the flag pile is irrelevant.
- `generate-exactclip-probe.mjs` — exact/exact pairs at 115/144/361/495 tw
  (the fixture's problem rows are 115 tw = 7.67 px, SHORTER than their line):
  flow = 2 x authored everywhere, rules or nil or neither. Clipping is
  irrelevant.
- `generate-exactmar-probe.mjs` — rows 3-5 of the caption table VERBATIM, then
  one property stripped per case. tcMar is the whole difference: Word spaces
  'for the' -> 'Rewugofi of' 5.75 pt wider with the fixture's margins
  (top 58, bottom 29 tw) than without, and the nils still charge zero
  (V0 55.03 px vs V2 58.99 with the rules live).
- `generate-exactpad-probe.mjs` — the exact-115 spacer's tcMar varied one side
  at a time, content rows untouched, against the (0,0) control:

      (top, bottom) tw    gap moves
      (58, 29)            +1.52 pt
      (29, 58)            +3.00 pt
      (58,  0)            +-0.00 pt
      ( 0, 29)            +1.25 pt

  **A pre-15 hRule="exact" row's flow is trHeight + its BOTTOM cell margin;
  the top margin adds nothing.** The compat-15 regime charges the TOP margin
  instead (probe-trheight), and neither regime charges rules to exact-row
  flow. #86's both-nil suppression was correct in every compat mode; what its
  guard removal deleted was a compensating half-rule-per-boundary charge that
  had been standing in for the missing bottom margins (0.5 pt per boundary
  against the true 1.45 pt per row).

One more term, measured from the fixture's own Word PDF rather than a probe:
the caption table's rows 0-1 are `tblHeader`, and the REPEATED instance of the
exact-144 row charges only HALF its bottom margin — Word's 'Vop' -> first-data
gap is 22.808 pt on p1 and 22.077 pt on p2/p7, bottomPad/2 = 0.73 pt apart.
p7 exposes this directly because it is the only continuation page that opens
on a text row instead of inside a split spacer row.

After the engine change (rowHeightFromTrHeight pre-15 exact branch + the
repeated-header path): all seven pages read 0.00, including p4's longstanding
1.97 — the same missing term, entering through the repeated header stack.
References: `parity/probe-exactnil11-word.pdf`, `probe-exactnil11p-word.pdf`,
`probe-exactclip-word.pdf`, `probe-exactmar-word.pdf`,
`probe-exactpad-word.pdf`.

### A table's OUTER edge takes one declarer: the nil zeroes it, and an exact first row absorbs even a live rule (#100)

probe-exactmar's V1 case carried the evidence sideways: lifting the caption
rows out of their table turned two interior boundaries into the table's OUTER
top and bottom edges, and against Word our V0/V1 (all-nil tcBorders under the
fixture's live sz-8 tblBorders) read +1.97/+2.00 pt — one full rule per outer
edge — while the nil-stripped V2/V3 agreed to 0.02 pt. Every earlier nil probe
(#51, #86) measured SHARED boundaries, where suppression needs both cells; the
outer edge has only one cell facing it and had never been isolated.

`generate-exactouter-probe.mjs` isolates it: two packages from the exactnil
base differing only in `compatibilityMode` (15 as authored, 11 rewritten),
target rows at the table's FIRST row (outer top) and LAST row (outer bottom),
one authored thing varied per case — no tblBorders (0), a live outer sz-12
rule (R), the same rule plus the target cell's nil on that edge (N) — across
exact 115 tw (X), exact 495 tw (Y, the scaling control) and content rows (C,
the V1 construct). Zero cell margins throughout so the pre-15 bottom-margin
charge cannot confound. Word END-REF spans, DIGIT-IDENTICAL in compat 11 and
15:

    case            0        R        N      charge of the live rule
    X (exact 115)  F 29.75  29.75  29.75    top: NOTHING
                   L 29.75  31.27  29.75    bottom: full 1.52pt
    Y (exact 495)  F 48.77  48.78  48.75    top: nothing (height-invariant)
                   L 48.75  50.28  48.77    bottom: full 1.53pt
    C (content)    F   —    37.52  36.00    top: full 1.52pt (CL0 36.00 control)
                   L 36.00  37.52  36.00    bottom: full 1.52pt

Three rules, all compat-invariant: **a live outer rule charges the flow its
full width — except above an exact FIRST row, which absorbs it into its fixed
height (the #51b own-side inset seen from outside); and a nil on the outer
cell edge alone zeroes the rule entirely, every variant reading the no-border
control to the digit.** The scaling sweep pins the charge as a border width,
not a height fraction.

The self-reproduction control earned its keep again: the compat-15 package's
export b laid the document-opening XF0 control ~31.5pt taller than exports
a/c/d (which are ink-identical at 192 DPI and arithmetic-correct), a
doerfp-class one-case wobble on the FIRST table of the document. Reference is
export a; compat-11's pair reproduced ink-identically first try.

Our engine before the change: nil ignored at outer edges (the V1 overcharge),
and pre-15 exact rows took the half-lead/half-inset convention (+0.76pt at the
top edge, -0.76 at the bottom against Word) that `exactInsetRow`'s compat gate
preserved. The change: `rowBorderWidths` honors a first/last row's all-nil
declaration at k=0/k=rows.length, and `exactInsetRow` drops the compat gate —
probe-exactnil11's full 2.00px inset had already pinned the inset half in
compat 11. After it, all 36 exactouter cases and all four exactmar variants
land on Word within 0.03pt.

**One scoped exception, fixture-calibrated.** us-courts-answer's body — one
124-row compat-11 table whose row 0 declares top nil under a live sz-8
tblBorders, repeating two tblHeader rows on every page — regressed p4 to 2.48%
when the nil was honored at CONTINUATION segment tops: pages 2-7 hold 0.00
only with the old 1.0pt charge kept there. No probe measures a
repeated-header segment top (probe-exactouter's tables are single-page), so
`nilSuppressedOuterTop` keeps the pre-nil charge on that path only, split
half-lead/half-row-0-instance exactly as the old arithmetic did. Word's model
at a repeated-header page top is measured only through this fixture; a probe
with a header-repeating table crossing pages would pin it properly.

Measured leftovers, not this change's: probe-exactnil11/-11p p2 read 67% in
ANY build (baseline 67.16, this branch 67.38) because Word spills C1's
trailing 12pt marker to p2 at a knife-edge p1 foot and we keep it — a page-fit
divergence, not a border one; and B1's mixed exact/content interior boundary
takes the full sz-12 rule in Word (MK-UP 26.25, flow +1.52) where we charge
half (25.50/+0.75) — the mixed case `exactInsetRow`'s comment calls unmeasured
now has a measurement, on the row BELOW the exact row.

Sentinels at the branch build, digit for digit against
`corpus-full-20260807-1116.log`: us-courts-answer 7x0.00, hftemplates 4x0.00,
caed-pleading 0.00, benchmark 0.00/0.37/0.35/4.00, tblextreme 2x0.00;
probe-exactmar/exactnil at 0.00. References:
`parity/probe-exactouter11-word.pdf`, `probe-exactouter15-word.pdf`.

### A repeated-header continuation top is the true outer top, nils included (#108a)

`generate-repeathdr-probe.mjs` measures the edge the #100 exception was
guarding: a tblHeader row repeated at the top of a continuation page, with
the outer-top declarations varied one at a time — no border, a live sz-12
top rule, the rule plus row-0 cell nils (the us-courts construct), sz-24 for
width scaling — over exact-115, exact-495 and content header heights. Two
packages differing only in compatibilityMode (15 / 11), each exported twice,
marks digit-identical across exports AND across compat modes.

Word's continuation page reads the FIRST page's arithmetic to the digit in
every one of the eleven cases: a live rule charges the flow its full width
above a content row 0 and nothing above an exact row 0 (full width as the
exact row's content inset instead), and a row-0 all-nil zeroes it — at the
continuation top exactly as at the true start. The engine's
`nilSuppressedOuterTop` charge contradicted all four nil cases by exactly
its own width and is removed; after the change all eleven land on Word to
the 0.05px text constant.

**What the fudge was hiding.** Removing it regressed us-courts p4 to 2.48%,
and pulling that thread surfaced SIX latent defects, each of which had been
canceling against another:

- **p4's ten uniform blocks ran 0.69px/block short.**
  `generate-uscourtsblock-probe.mjs` rebuilds the fixture's
  content-row + atLeast-spacer block VERBATIM and strips one authored thing
  per case; round 2 (`generate-uscourtsblock2-probe.mjs`) isolates
  single-kind stacks. Word's pitch is INVARIANT to the tblBorders insideH
  (V2), the spacer's sz-1 cell rule (V1), the tblPrEx (V3), the run mix (V5)
  and the FORMTEXT (V6); it moves only with the leading empty paragraph (V4)
  and the trHeight floor (V7). The mechanism: **a boundary resolves per grid
  column, and the table-wide rule enters only through a side whose cell is
  SILENT — a cell that declares its edge (nil or a width) replaces insideH
  on its side** (SN 21.68 vs CN 21.67 under a live sz-8 insideH, where
  charging insideH reads +1.33). And **the pre-15 atLeast floor is
  trHeight + topPad + bottomPad, no haircuts** (S0: 624tw + 58 + 14tw reads
  46.40px in compat 11 and 15 alike; the old topPad−0.25pt / drop-sub-1pt-
  bottomPad haircuts carried no probe). With both, the fixture block reads
  85.01px against Word's 84.69–85.04 quantization band, and p4 goes to 0.00.
- **A cantSplit row gets ZERO overhang allowance.**
  `generate-rowfit-probe.mjs` / `-rowfit2-` sweep the room under a cantSplit
  row in 2px steps on the us-courts package: Word moves it at room 30 and
  keeps it at 32 for a 32.00px row — exact-rule lines AND natural 11pt
  lines, with and without a footer. The old allowance (the FOOTER HEIGHT,
  compat-11 cantSplit only) was compensating the next defect exactly.
- **A page-framed footer paragraph consumes no footer flow, and an
  effectively empty footer still reserves its w:footer distance.** The
  fixture's footer is a page-anchored 'Page N of M' frame (vAnchor="page",
  y=15264tw) plus one empty paragraph; charging the frame clamped our body
  at 986.9px where Word's row decisions bracket the bottom in [992, ~1005]
  — margin bottom 1027.2, footer distance floor 1017.6, minus the empty
  paragraph's line. The measure now excludes page-framed paragraphs, and a
  footer PART floors the body at pageHeight − footerDistance − footerH even
  when footerH is 0.
- **The bottom cell margin is charged in full.** The signature rows
  (trHeight 20tw atLeast, tcMar 58/43) measure 23.6px in Word = line + both
  margins exactly; a compat-11 haircut (half over 2pt, quarter under, on
  sub-2pt atLeast rows) was canceling the insideH overcharge.
- **A run whose visible text is entirely w:sym-drawn takes the sym's font
  metrics.** The fixture's Symbol-font minus signs make Word's 11pt Times
  lines ~18.0px (Symbol's 1.225em box) where a plain line reads ~16.7;
  we mapped the glyph but kept Times metrics, so sym-bearing pages drifted
  ~1.3px/line.

**Filed, not closed: the two-row repeated stack.** us-courts p6 reads 0.91%
and p7 8.01% at the branch build — one knife-edge row fit (≤1px on BOTH
engines' side of the decision) that traces to the fixture's TWO-row header
stack (content caption + double sz-8 cell bottom + exact-144 all-nil row).
`generate-repeathdr2-probe.mjs` measured it verbatim: Word's continuation
stack equals its first-page stack exactly (contradicting the repeated-
instance bottom-margin halving probe-exactpad read off the fixture — the
discriminating variable is likely the following row's top margin, 58tw in
the fixture and 0 in the probe), the double border charges NOTHING to flow
(W1 = W0), margins enter at exactly 7.66px (W3), and the exact-144's base
charge reads ~5.15px against the authored 9.6 — a number no model tried
(trHeight, +margins, collapse) reproduces. The stack needs its own
decomposition with instrumented line boxes before anything else moves.

Other measured leftovers: Word splits a splittable two-line row 1+1 at
rooms below its height (rowfit P26–P30) where splitLaidRow rejects one-line
fragments; non-cantSplit rows of exact-rule lines keep a 1.78px overhang
under ROW_OVERHANG_TOL=3 that Word moves (probe-repeathdr Y p1, 51 vs 50
rows); and rowfit2's DI case reads +2.37 for a declared sz-12 against nil
UNDER a live insideH sz-8 where DW (no insideH) reads exactly +2.00.

Sentinels at the branch build: us-courts 0.00 ×5 / 0.91 / 8.01 (p4's
longstanding deficit closes; p6–p7 blocked on the stack above), hftemplates
4×0.00, caed-pleading 0.00, staging-eastasian 0.00, benchmark
0.00/0.37/0.35/4.00, tblextreme 2×0.00, probe-exactmar/nil/nil11/nil11p
0.00 — **including probe-exactnil11 p2, whose 67% any-build page-fit
divergence closes to 0.00** — exactouter11/15 0.35/0.00 (weight channel
only; line and align 0.00). References:
`parity/probe-repeathdr11-word.pdf`, `probe-repeathdr15-word.pdf`,
`probe-uscourtsblock11/15-word.pdf`, `probe-uscourtsblock2-11/15-word.pdf`,
`probe-rowfit11-word.pdf`, `probe-rowfit2-11-word.pdf`,
`probe-repeathdr2-11-word.pdf`.

### A mixed exact/content boundary gives the whole rule to the row below (#108b)

`generate-mixedbound-probe.mjs` varies the interior boundary one side at a
time, both orders — exact-495 above / content below (EC), the inverse (CE),
and same-kind controls (CC, EE) — under no rule, a live insideH sz-12,
both-nil, each one-sided nil, and sz-24 width scaling. Two packages
(compat 15 / 11), two exports each: marks digit-identical across exports
and compat modes.

    case          UP-REF   MK-UP   END-MK  END-REF     what it pins
    EC0/ECR        16.00  33→35     16.00   65→67   full rule BELOW, flow +full
    ECW            16.00   37.00    16.04   69.04   scales: sz-24 charges 4.00
    ECU/ECL        16.00   35.00    16.00   67.03   one-sided nil suppresses nothing
    ECN            16.00   33.00    16.00   65.00   both-nil charges zero
    CE0/CER        16.00  16→18    33→31    65.00   exact row ABSORBS it: inset +full, flow ZERO
    CEW            16.00   20.00    29.00   65.00   scales, still absorbed
    CC0/CCR        16.00  16→18     16.00   48→50   content/content: half/half already right
    EE0/EER        16.00  33→35    33→31    82.00   exact/exact: already right (#100)

The rule: **an interior boundary's full painted width belongs to the row
BELOW — inset into a content row (flow grows by it) and absorbed by an
exact row's fixed height (flow unchanged). We charged half in both mixed
orders**; `rowBorderShare` and the cell-inset path now take the full width
below an exact row and nothing above one, and content/content keeps the
measured half/half split. After the change every case lands on Word to
0.04px. This closes the B1 leftover #100 filed (Word 26.25 / flow +1.52pt
where we read 25.50 / +0.75).

### list, process and hierarchy pinned to their caches by construction (#94 phase 2)

The three non-cycle families shared a lin/roundRect layoutDef that
contradicted their cached drawings exactly the way cycle's did. Rather than
calibrate three more Word evaluation models, phase 2 uses what the cycle
campaign proved — Word honors refType w/h constraint fractions to 0.01pt —
plus one structural fact: the engine regenerates the layoutDef WITH the data
on every insert and edit, so the node count is known. Each family now emits
a COMPOSITE layoutDef whose per-node l/t/w/h constraints are fractions
computed from the same `diagramShapes()` the cache is built from; the two
descriptions agree by construction. `primFontSz` is pinned at 12 (no autofit
rule), and process/hierarchy connectors are dropped from the caches — the
layoutDef declares none, so Word draws none (cycle's one-sided-drop rule).

Calibration: one document per family (3 nodes), each exported twice by
desktop Word, byte-consistent (`internal/scripts/gen-smartart-cal2.mjs`,
`export-smartart2.mjs`, `read-smartart2.py`). Word's re-evaluation
reproduces every cache digit for digit:

    family      cache node sizes (pt)        Word painted     offsets
    list        345.00 x 50.00 (x3)          345.00 x 50.00   all (72, 88.97)
    process     110.00 x 66.93 (x3)          110.00 x 66.93   all (72, 88.96)
    hierarchy   125.98 x 44.09 root,         identical        all (72, 88.97)
                168.75 x 48.82 children

— the (72, 88.97)pt offset is the inline anchor (page margin + first line),
constant across every node. Colors land per node with no cycle-meth
blending, and the text is Calibri-Bold 12.0 white on every node in all
three families. Sentinels: word-interop-smartart-only 0.21%,
-embedded 0.00% — their recorded values; the cycle branch is untouched and
the fixtures carry only the cycle family, so no fixture regeneration was
needed.

### us-courts p6/p7: the "~5.15px exact-row charge" was three charges canceling (#109a)

The instrumented line-box decomposition #108a asked for was run (temporary
per-component logging in the worktree engine: per-row contentH / borderShare /
trHeight branch, per-line metric spans), and the mystery number dissolves
without any new exact-row model. **Word's exact-144 row charges exactly what
the pre-15 rule says — trHeight 9.6 + its 1.93px bottom cell margin.** The
"~5.15px base charge" was an artifact of assuming the caption row's line is
~16.9px (11pt): Word's caption line measures **12.37px = the 8pt Times
natural (8 x 1.1597em = 9.2776pt) exactly**, and with that one number the
whole probe-repeathdr2 table closes to 0.1px — W3 stack 22.00 = 12.37 + 9.63,
W0 29.66 = (3.87 + 12.37 + 1.93) + (9.6 + 1.93), W2's border-to-content
control 41.02 = row0 18.17 + (3.87 + 17.0 + 1.93), and the double border's
painted band top sits at row0's bottom edge on every case.

Three engine charges were mutually canceling on this fixture, and each is now
replaced by a measured rule:

1. **A hidden field's zero-width strut sized the line.** The caption
   paragraph opens with `SEQ CHAPTER \h` (hidden, empty result), whose
   metricsStrut atom carried the default-12pt font and sized the line at
   18.40px against Word's 12.37. The strut now sizes a line only when
   nothing visible shares it — same shape as the existing whitespace-run
   rule.
2. **The repeated exact row's bottom-margin halving is refuted.**
   `generate-repeathdr3-probe.mjs` sweeps the follower row's tcMar top
   (0/29/58tw) against the exact row's tcMar bottom (0/29/58tw) on the
   fixture's own stack; both exports ink-identical. Word's continuation
   stack equals its first-page stack **to 0.03px in all five cases**, and
   B58-F58 pins the bottom margin charged in FULL (+2.00px for +29tw). The
   halving read off the fixture (probe-exactpad p1 vs p2/p7) was a
   confound: the continuation page's first data row is a different row from
   p1's. Reference: `parity/probe-repeathdr3-11-word.pdf`.
3. **The flat 7pt empty-header body-top charge carried no probe.** Word's
   us-courts body top is 70.87 - 3.87 (tcMar) - 0.54 (8pt glyph offset) =
   66.5 = headerDistance 48 + the empty header paragraph's FULL 18.55px
   default-12pt line, to 0.1pt. compat<15 all-empty-paragraph headers now
   charge their measured height like every other header. (nccih and
   chem-omml, the other compat<15 empty-header fixtures, have top margins
   that govern either way — the branch was live only on us-courts.)

wild3-template-us-courts-answer: **7 x 0.00** (p6 0.91 -> 0.00, p7 8.01 ->
0.00 from the 25d7298 baseline). The p7 shape was one knife-edge cantSplit
signature row ("Tosutuveh Manive") we kept at the p6 foot where Word moves
it; with the stack corrected it moves. probe-repeathdr2 reads 8 x 0.00 and
probe-repeathdr3 lands on Word within 0.2px in every case, first page and
continuation. The two #108a side-filings (the 1+1 two-line row split, the
1.78px non-cantSplit overhang) were NOT reached by this decomposition and
remain filed. probe-uscourtsblock/rowfit/mixedbound severities are
digit-identical to the untouched 25d7298 build (their nonzero pages are
pre-existing sweep scaffolding), and the full sentinel set holds.

### Word's MS Mincho snap natural is 1.296em; ja lines now take the compat-15 grid snap (#109b)

`generate-docgrid15b-probe.mjs` -> `fixtures-staging/probe-docgrid15b.docx`
(reference `parity/probe-docgrid15b-word.pdf`, two exports ink-identical)
brackets the quantity the lines-grid snap tests: one 3-line MS Mincho
paragraph per case, ja 10/11/12/16pt over pitches 240/300/330..480tw, each
case's 1-vs-2-row answer bracketing Word's natural em against
pitchPx/sizePx:

    case      threshold   Word      case      threshold   Word
    J11P240     1.091     2 rows    J11P300     1.364     1 row
    J16P360     1.125     2 rows    J10P300     1.500     1 row
    J10P240     1.200     2 rows    J12P360     1.500     1 row
    J12P300     1.250     2 rows    J16P480     1.500     1 row
    J16P420     1.3125    1 row     J11P360     1.636     1 row

**Every threshold <= 1.25 snaps and every threshold >= 1.3125 does not, so
Word's MS Mincho snap natural lies in (1.25, 1.3125)** — consistent with the
1.296em the natural-pitch contexts measured (probe2-ruby-vertical's 20.5px
vertical columns = 1.296 x 1.0792 at 11pt) and excluding both our raw
hiragino profile's 1.643em and the 1.4em a vertical reading would suggest.
The 1.643 was reverse-engineered from staging-eastasian's 19.5pt/line
advance, which is that fixture's grid pitch (360tw x 1.0792), not a natural
— the browser's true Hiragino fontBoundingBox reads ~1.0em, so the
overstatement lived entirely in the calibrated profile.

Engine (close-wave): the hiragino profiles now carry 1.296em directly
(split proportional), the ja targetEm rescale goes (identity), and the
textSnap EA carve-out narrows to the unmeasured zh/ko fallback faces. The
vertical (tbRl) lines-grid flow is pinned at the old 1.643em column pitch —
probe2-ruby-vertical p2 was measured good there and vertical remains
unprobed. Verified: **probe-docgrid15 F360 38.57% -> 0.00** (the filed EA
false-snap case), all 10 docgrid15b cases land on Word's row counts in the
browser, staging-eastasian 0.00, ruby-vertical 0.16/0.00 against the corpus
0.15/0.02.

Still open (unchanged): the zh fallback's snap metric. PingFang/Songti keep
both the 1.733em natural-pitch rescale and the snap carve-out; a
docgrid15b-shaped sweep with zh text would pin them the same way.
