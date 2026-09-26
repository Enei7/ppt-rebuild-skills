# Reconstruction and QA recipe

Read this when starting a PDF run, writing formula inventories, or auditing a final deck. This skill bundles the `editppt` CLI and its manifest schema; use the bundled runtime instead of inventing run-state files.

## Source-page inspection

For every page, capture source page number, pixel dimensions, legible text, formula regions, and foreground figures. Render/crop at sufficient resolution to read small symbols and QR modules. Prefer extraction of an embedded PDF image when it is the exact visual; otherwise crop the original PDF page rendering. Keep the crop file plus PDF page number and crop rectangle. Check the crop's pixels against the source before using it. A source crop is a selectable bitmap, not an editable drawing.

Do not convert a whole page or major content block into a single image merely because the source is complex. Text, simple layout geometry, and tables stay native under the page decision tree. A user-requested exact QR/logo/hand-drawn diagram is a narrow exception to the generated-asset default: mark its manifest asset provenance as `user-provided`, set `exact_source_crop: true`, `source_page`, and `source_crop_px`, point `source` at the page's `source.png` (or another existing file from which the crop was made), and explain the crop in `provenance_note`. Keep positioned `images[].box_px` and the source-pixel coordinate system required by the manifest schema. The bundled validator accepts this documented exception; if evidence is missing, fix it rather than forging validation or silently substituting a generated figure.

## Formula inventory contract

For each standalone mathematical expression, create a unique per-page id and a source-pixel box. Preserve exact variables, indices, matrices, delimiters, and display alignment in its LaTeX source. A sample inventory entry follows; adapt its `image` path to the actual intermediate formula object.

Keep the source box formula-only, with modest padding. Do not include the next prose line to make the box taller: source-ink measurement will then calibrate to that unrelated line and can move an otherwise correct inline equation into the surrounding text. Check the crop itself, not just the page preview.

PDF text extraction is only a hint for inventory coverage: font-name heuristics can miss whole equations, and reading-order text can flatten superscripts (for example `e^{3x}` into `e3x`). Check the visible source even when extraction succeeds. Do not mark mathematical text as fully editable merely because its flattened characters occupy a native text box; required equations belong in the native-math inventory.

```json
{
  "id": "formula_03",
  "box_px": [118, 264, 430, 78],
  "tex_source": "assets/formula_03.tex",
  "image": "assets/formula_03.svg",
  "decision": "latex-rendered-image",
  "editable": false
}
```

The page-stage image is temporary and must be recorded honestly as non-editable. After `editppt run finalize`, `equationize.py` reads the inventory, transforms LaTeX → MathML → Office Math via `MML2OMML.XSL`, and replaces the picture or a uniquely named candidate with a native equation at the manifest box. `editable: false` describes the page-stage asset; the delivered PPTX's native equations are verified separately. If an expression was not converted, report it as non-editable. Do not fake a formula with Unicode superscript fragments or an invisible equation over a visible picture.

Inline expressions may stay in editable body text when that is both legible and within the user's editability requirement. If the user asks that **every** expression open in the equation editor, inventory and convert inline expressions too, then recalibrate their baselines and boxes individually. Transcribe the source's *math layout*, not just its value: `1/(z-z_0)` is a linear inline expression, whereas `\frac{1}{z-z_0}` becomes a stacked fraction that may overflow a prose line. Keep parentheses around the denominator when using a slash.

Use `\tfrac` for genuinely compact stacked source fractions; the converter preserves them with a smaller native math argument, not by shrinking the entire equation. The supported argument-size mechanism is documented by [Microsoft](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.math.argumentsize). Plain `\frac` remains a full-size fraction. Keep punctuation attached to the adjacent equation/prose when possible: a narrow standalone period box may be shrunk by text fitting.

Bare `\int` must become a native n-ary operator even without limits; an ordinary integral character has the wrong height. Explicit `\left...\right` fences are normalized to native delimiters so they grow around fractions. The converter's self-check covers both structures, but a PowerPoint save/reopen and visual check remain required.

Matrix environments (`pmatrix`, `bmatrix`, `Bmatrix`, `vmatrix`, `Vmatrix`) also need structural native delimiters. Some MathML converters emit their fences as ordinary adjacent characters, which PowerPoint displays only at the middle row's height. The converter normalizes the matched pair around the matrix before transformation; verify all matrix rows stay inside the fences after a PowerPoint save/reopen. Retain the native matrix rather than covering a broken fence with a bitmap.

Treat source-positioned math blocks as separate objects even if they share a row. A left-hand equation and a far-right condition require independent formula-only crops, ids, and `box_px`; do not merge them with `\qquad` to imitate horizontal page spacing. The page-stage crop can hide this error because its pixels already contain the gap; after native conversion Office Math recomputes TeX space and pulls the condition left. Render at least one converted representative before finalization, then fit each equation to its own source box.

## Font and geometry

For a uniformly colored formula, record `native_color_hex` (six hex digits, e.g. `FF0000`) in its inventory entry; temporary crop colors are not automatically inherited by Office Math. Mixed-color expressions need separately inventoried source blocks or explicit verified math-run styling.

Bold mathematical letters must survive a PowerPoint save/reopen round trip. The converter represents bold Latin letters as ordinary Unicode letters plus explicit math style, avoiding a PowerPoint font-edit bug that can split supplementary-plane glyphs into invalid surrogate characters. Calibration audits the saved XML immediately; a failed audit is a blocker, not a warning to ignore.

Inspect function-name spacing in the native render (for example `\sin x`, `\cot x`, and `\sec x\tan x`). Office's MathML stylesheet can discard TeX thin-space nodes even when the source crop looks correct. The converter inserts an explicit thin-space Office Math run at adjacent function/letter boundaries; keep this covered by `equationize.py --self-check`. Do not compensate by moving the entire equation or baking the formula into a picture.

Use Microsoft YaHei for ordinary visible text, including explicit run fonts, end-paragraph properties, and theme aliases where applicable. Do not force the equation math run to YaHei; Office math glyphs use a math font, normally Cambria Math. Match **visible glyph bounds and line baseline**, not just nominal point sizes. A `16 pt` Office summation can look less than half the height of its PDF counterpart. For each page, prefer source text boxes and `text_hints` as a starting point, then inspect the rendered slide. Small OCR noise does not justify changing a clearly visible word or formula.

Dense formula tables especially need a source-supported `native_font_pt` starting size. Crop height includes numerator/denominator and padding; deriving independent point sizes from that height can make fractions much larger than adjacent one-line entries and collide across rows. Keep one visual math text level where the source uses one, then calibrate and inspect the native result. This is not a universal point-size rule.

For remaining size outliers, render the affected equation alone on its original slide, compare its full ink bounds against the formula-only source crop, then resize and rerender before translating its ink center. This avoids measuring clipped native ink inside an overly tight source rectangle. If a tall integral alone forces all adjacent letters too small, calibrate the n-ary control glyph separately from the math runs instead of shrinking the whole expression; preserve the editable operator, inspect the full-slide result, and audit again after saving.

Run `calibrate_math_geometry.ps1` after equation conversion, using a fresh output file. It first shrinks native equations whose PowerPoint bounds exceed their source boxes (including fractions), then measures equation ink in the original `source.png` and PowerPoint-rendered formula-only slides, adjusts nary/wide-display equations from Office `BoundLeft/Top/Width/Height`, and writes a report under the run directory. A source crop touching an edge is not automatically invalid; a partial glyph with too little ink is. Do not blindly enlarge every inline formula: stacked Office fractions may need a different LaTeX structure when the source uses compact linear notation, and split text/formula boxes may have inaccurate source bounds. Inspect report warnings/outliers, then compare full-slide renders for collisions and baselines. Run `check_formula_text_spacing.ps1 -ReportTightGaps -ReportBaseline` on the actual final candidate; its ≥4 pt single-line overlaps are blockers, near-zero word junctions and >5 pt compact-inline center differences need visual review. Repair adjacent text and formula objects together so that widening a gap does not create a collision with the next run. A bounded vertical correction is reasonable only after checking the rendered line; nary operators and fractions have inherently taller ink and must not be centered like letters. Rerender and run the checker with `-FailOnCollision -ReportBaseline` after repairs. Preserve the native equation object; do not hide a source bitmap over it. Legacy `normalize_math_font.ps1`/`check_math_font.ps1` are useful only when no source crop is available; their mean-point-size rule is not a valid acceptance threshold after source-geometry calibration.

## Validation and recovery

Before equation conversion, require every page `validation.json` to have top-level `passed: true`, then `editppt run record` and `editppt run finalize`. Keep its output unchanged as the recovery deck. Run equation conversion and source-geometry calibration to separate output paths. The later passes mutate only copies of the finalized PPTX; running `editppt run finalize` again would replace them with the base output. Run `equationize.py --audit-existing` on the **actual final file after its last save**; its empty-nary gate must pass before delivery. If PowerPoint COM is unavailable, run that same structural audit and explicitly report the missing visual-baseline calibration.

Audit the final file on three levels:

1. **Structure:** valid/openable PPTX, correct page count/order, expected native `m:oMath` count, zero empty `m:nary/m:e` bodies (sum/integral dotted placeholders), all exact source assets present, and no full-page source raster duplicated under editable overlays.
2. **Typography/geometry:** YaHei body text, source-visible formula size/position, no clipped or occluded text, and no major baseline or line-wrap drift. Inspect every formula sharing a line with prose, including 1–2 pt gaps that render as words glued to math; a formula-only render cannot expose collisions.
3. **Visual fidelity:** render all slides and compare each against its PDF page; inspect QR scanability, crop edges, diagrams, tables, symbols, superscripts/subscripts, and layout. Specifically inspect slides with summations/integrals for dotted empty boxes. Verify one representative native formula in PowerPoint's equation editor.

Use `python <skill-root>/scripts/audit_deck.py --run <run> --deck <final.pptx> --report <new-report.json>` for the inventory gate. It checks formula IDs per presentation-order page, not only the aggregate equation count, and reconciles remaining pictures after formula replacement. A deck with zero equations can pass the standalone empty-nary check while still missing every required formula; the inventory gate rejects that case. A report path must be new to preserve earlier evidence.

For multiple input decks, keep one run and one final file per input unless the user requests a combined deck. Preserve filenames with spaces when resolving inputs. Page authoring may run in parallel across independent runs, but serialize desktop PowerPoint operations: its shared COM application can otherwise mix selections or be closed by another worker. Track and validate each deck separately; a successful chapter does not imply the entire batch passed. Do not publish source PDFs or generated decks along with reusable skill changes unless specifically authorized.

For a failed formula: inspect the source PDF crop and LaTeX, correct the inventory entry, rerun equation conversion from the unchanged finalized deck, and recalibrate. For a failed page object: repair its page-owned manifest/assets, rebuild/validate/record that page through `editppt`, re-finalize, then rerun post-processing. Do not repeat an unchanged failing command. If a library, Office stylesheet, PowerPoint COM, worker, or source detail is genuinely unavailable, preserve successful artifacts and report the limitation rather than calling the deck finished.
