# Source-reviewed mixed-line reflow

Use `scripts/reflow_mixed_lines.ps1` only after visually reviewing a prose/math
line against the source page. The script does not discover lines or infer a
baseline. It resolves every explicit selector, normalizes only prose to the
line's `text_font_pt`, measures live PowerPoint `TextRange2.BoundWidth`, fails
on overflow, and then moves the named objects. Native equations keep their
existing font, color, AutoSize setting, and Office Math structure.

Each recipe line requires `page`, `text_font_pt`, `center_y_px`, `start_x_px`,
`right_limit_px`, and ordered `items`. An item selects either `shape`, or exact
`text` with optional zero-based `match_index` in slide order. Use
`expected_text` (or `exact_expected`) when the selected content must match
exactly. Optional `gap_before_px` and `center_offset_px` are source-reviewed
adjustments.

For example, after checking this three-object sentence against the source:

```json
[
  {
    "id": "p4-condition",
    "page": 4,
    "text_font_pt": 18,
    "center_y_px": 680,
    "start_x_px": 140,
    "right_limit_px": 2240,
    "items": [
      {"text": "When", "gap_before_px": 0},
      {"shape": "formula_condition", "gap_before_px": 12},
      {"text": ", we obtain", "gap_before_px": 0}
    ]
  }
]
```

These numbers are examples, not defaults. Coordinates are in the source
image's pixels and are mapped through the manifest's `content_box`, including
its slide offset. `center_y_px` is a reviewed visible line center, not an Office
typographic baseline. Tall sums, limits and stacked fractions can need a
source-supported `center_offset_px`; never center all tall math mechanically.
Keep punctuation close to its preceding expression instead of inserting the
normal word gap before a comma, period or closing bracket.

Prose shape names need not match manifest text-box IDs. An exact text selector
must be unique unless `match_index` explicitly selects its zero-based occurrence.
Native equations are selected by their real PowerPoint shape name. The
preflight resolves PowerPoint's active `mc:Choice` rather than counting its
fallback representation as a second object. Repeated selectors, incorrect
expected text, unknown fields and non-finite coordinates fail before mutation.

Optional item `font_pt` changes prose only; a math override is ignored with a
warning. Calibrate a formula's glyph size separately before reflow. After live
width measurement, any overflow aborts without saving or silently shrinking
the sentence. Revisit the source line break or intended span; do not reduce
all font sizes until the error disappears. The original input is never saved
over. `-Pages` can limit application to selected pages after recipe validation.

The width check reports all overflowing lines in one pass. A failed live-width
preflight preserves `<output>.preflight-plan.json` and
`<output>.preflight-widths.json` without saving a changed deck. Use these exact
measurements to repair source-supported line breaks or available spans together;
do not blindly expand every limit to the slide edge. Choose a fresh output path
for the next attempt so earlier failure evidence is retained.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File <skill-root>/scripts/reflow_mixed_lines.ps1 `
  -Deck <native-input.pptx> -Run <run-dir> -Lines <recipe.json> `
  -Output <fresh-output.pptx> -Python <same-python>
```

The output geometry report is written beside the fresh deck. Run the native
equation audit and full-slide visual QA afterward; this helper does not replace
either gate.
