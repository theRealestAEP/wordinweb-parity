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

`wild2-legal-ca-agreement` is the same story and its reference is stale too. It
carries 17 hints; strip them — 442 bytes, nothing else — and Word exports 22
pages where the cached reference has 23. The extra one is a blank verso at
page 2 that only the stored pagination contains. Word's computed layout of the
untouched fixture is page-for-page identical to Word's layout of the
TOC-inserted save, so the toc-insert edit changes Word's pagination not at all;
it only disturbs the file enough to make Word recompute. Our 23rd page is ours,
and it is there before any edit.

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
it on the current page whatever the room — 18 px was still enough — and starts
the new page after it. We apply the ordinary test and spill it.** In the fixture
that is 22.1 px of room against our 32.7 px demand, so we spill, and the spill
costs the page.

Both this and the two #38 rules are the same shape of defect: a quantity Word
declines to charge at a boundary, charged in full by us.

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
