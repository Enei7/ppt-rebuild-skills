---
name: ppt-rebuild
description: Rebuild existing PDF slide decks, especially mathematical lectures, as visually faithful and object-editable PowerPoint with native editable Office equations, calibrated math/text alignment, and exact source-figure crops. Not for creating a new image-only presentation from an outline.
---

# PPT Rebuild

Reconstruct an existing PDF deck as an editable `.pptx`. This skill contains its own `editppt` page runtime, page-worker prompt, manifest rules, equation conversion, and visual calibration scripts. **Do not require installation of another skill.** Ordinary text and layout should be native PowerPoint objects; required formulas must become native Office Math (`m:oMath`), not pictures with hidden editable overlays. Exact source figures and QR codes may remain individually selectable bitmap objects; never call their internal pixels editable.

## Start here

1. Confirm the source PDF, output path, page count/order/aspect ratio, legibility, formula density, and figures. Inspect representative pages and record unreadable source details rather than guessing.
2. Handle runtime setup on first use; do not send the user a separate skill-installation checklist. Select Python 3.10 or newer (if the default `python` is older, locate a compatible interpreter). If the bundled CLI dependencies are missing, install them from `<this-skill-root>/cli` with that interpreter: `python -m pip install -e <this-skill-root>/cli`. Check imports with `python -c "import fitz, PIL, numpy, lxml, latex2mathml"`, then run `python <this-skill-root>/scripts/editppt.py doctor`. This installs Python packages, not another skill; if no compatible interpreter is available, explain the prerequisite instead of claiming the conversion can run. On Windows, locate Office's `MML2OMML.XSL` and PowerPoint for the full geometry pass.
3. Read [references/cli-helper.md](references/cli-helper.md) for the commands you are about to run. Read [references/manifest-schema.md](references/manifest-schema.md) before writing a page manifest, [references/page-decision-tree.md](references/page-decision-tree.md) for page object decisions, and [references/reconstruction-and-qa.md](references/reconstruction-and-qa.md) for PDF crops, formulas, and final QA. The generated page prompt routes the page owner to the relevant sections.

In all commands below, `editppt` means **`python <this-skill-root>/scripts/editppt.py` using that same interpreter**. Generated page prompts contain the exact bundled command. Do not call a global `editppt` executable merely because it is on `PATH`.

The user may choose offline processing. Respect that choice: do not upload source pages for OCR or image editing. PaddleOCR is optional; if no token is configured, explain the offline text-hint limitation once before page work, then proceed if the user declines or already chose offline. If a token is configured but network approval fails, follow the user's data-handling choice and the runtime's approval mechanism. Send only task-local page images/prompts to configured services, never credentials or unrelated files.

## Page reconstruction and state

1. Run `editppt prepare <pdf>` to create an isolated run directory. It renders pages and produces `deck_manifest.json`, `page_jobs.json`, and `pages/page_NNN/source.png` with text hints. Choose the recorded image backend according to available tools: built-in `image_gen.imagegen` first when callable, then the bundled `editppt image` fallback only for its documented fallback conditions. Source-exact PDF crops and formulas do **not** need generated images.
2. Advance state with `editppt run next <run>`. For multi-page runs, dispatch page workers when available. If page-worker tools are unavailable, use `editppt run next <run> --local` and rebuild pages sequentially. Never hand-edit run-state JSON.
3. For each suggested page, generate its prompt with `python <skill-root>/scripts/build-page-worker-prompt.py <run> --page <page_id> --out <absolute-page-dir>/worker-prompt.md`. Spawn the worker before `editppt run dispatch` in worker mode; in local mode claim with `editppt run dispatch <run> --page <page_id> --agent-id main --prompt-file <absolute-prompt> --local` before writing any page artifact. Only the page owner writes that page's manifest/assets/PPTX/validation. In sequential local mode, record the current page before claiming the next.
4. Inventory background, foreground visuals, editable text/shapes/tables, and every required formula before building. Use source-pixel geometry and measured text hints as starting evidence, then verify against the visible PDF. Write the page manifest and formula inventory. Build with `editppt page build <page-dir>`, compare the preview, run `editppt page validate <page-dir>`, and repair page-local failures. Required formulas may be temporary page-stage images only if their inventory has accurate LaTeX and source box.
5. Record a passing page with `editppt run record <run> --page <page_id> --agent-id <id>`. Do not reset a slow active worker. Retry only after changing the failed input/condition. When all pages are recorded, run `editppt run finalize <run>` and preserve its output unchanged as the pre-equation recovery deck.

## PDF source fidelity

- Preserve page order, layout, mathematical meaning, and significant visual elements. Aim for close correspondence, not a false pixel-identity claim: Office font rendering and editable reflow can differ.
- Never use a full PDF page as a raster background behind editable text. Native text, simple geometry, and regular tables remain native.
- When the user wants an **exact original** logo, QR code, photograph, or complex diagram, extract the embedded image or crop the PDF page render. The manifest's `user-provided` provenance must include `exact_source_crop: true`, one-based `source_page`, source-pixel `source_crop_px`, an existing `source` file, and a clear note. This is the documented exception to the ordinary foreground asset-sheet route. A source crop is a movable/replaceable image, not an editable drawing.
- For other foreground visuals, follow the backend and source decisions in `page-decision-tree.md`. Do not generatively redraw an explicitly source-exact figure.

## Native formulas, fonts, and alignment

- Transcribe LaTeX from the visible source; OCR hints do not prove symbols, indices, matrices, or delimiters. Every standalone formula, and every inline formula the user requires in the equation editor, needs a unique `formula_inventory` id, accurate source `box_px`, and `latex` or `tex_source`. Reference the temporary image with `image`/`replace_image_id` so conversion replaces it rather than overlaying it.
- When one visual row has spatially separate math blocks (for example a series on the left and its convergence condition on the right), give each block its own source crop, inventory id, and source box. Do not encode a page-scale gap with `\qquad` or other TeX spacing: Office Math may collapse it even though the temporary image looks exact. Check the native render, not only the page-stage preview.
- Do not encode multiline prose annotations under `\int`, `\sum`, or other n-ary operators with `\substack`/`\text`. The resulting Office Math can make PowerPoint refuse to open the entire deck. Keep the operator/formula native, and place its multiline annotation in separate editable text/math objects at the source position. Test the native page in PowerPoint before whole-deck conversion.
- If no local TeX renderer is available, use an **exact, formula-only source crop** as a temporary page-stage image. Record truthful `user-provided` provenance with `source: source.png`, `exact_source_crop: true`, one-based `source_page`, and `source_crop_px`; keep the verified LaTeX and image reference in `formula_inventory`. This fallback is only for building the intermediate deck: native equation conversion and the final image-replacement audit remain mandatory.
- Use Microsoft YaHei for editable ordinary text when requested. Office equations normally render with Cambria Math. Match their **visible glyph height, width, and baseline** to the nearby source text; matching nominal point size alone is insufficient. Preserve compact inline `1/(z-z_0)` when the source is linear; do not turn it into a tall stacked fraction merely for semantic equivalence.
- Keep each uninterrupted prose span as one editable text box (or a few intentional line boxes), not one box per OCR word. Split only around actual inline formulas or layout boundaries, and reserve a visible gap on both sides of each formula. Per-word positioning based on a serif PDF often collides after the text becomes Microsoft YaHei; inspect the rendered junctions.
- Convert **a copy** of the finalized deck. On Windows with PowerPoint installed, run:

```powershell
python .\scripts\equationize.py --self-check
python .\scripts\equationize.py --run-dir <run-dir> --input <editppt-final.pptx> --output <native-math.pptx>
python .\scripts\measure_math_geometry.py --self-check
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\calibrate_math_geometry.ps1 -Deck <native-math.pptx> -Run <run-dir> -Output <calibrated.pptx> -Python <same-python>
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\check_formula_text_spacing.ps1 -Deck <calibrated.pptx> -Run <run-dir> -ReportTightGaps -ReportBaseline -FailOnCollision
python .\scripts\equationize.py --audit-existing --input <calibrated.pptx>
```

Use a compatible `MML2OMML.XSL` with `--mml2omml <path>` if not auto-detected. The process-scoped PowerShell execution-policy flag changes no machine policy; inspect scripts before running them. Geometry calibration first fits oversized native equations, including fractions, inside their source boxes, then calibrates n-ary/wide math against source ink; it is not a blanket point-size rule. Ambiguous compact inline formulas still need visual review. After the **last** PowerPoint save, repeat `--audit-existing` on the actual delivered file. If PowerPoint COM is unavailable, run the structural audit and explicitly disclose that visual baseline calibration was not performed; do not present it as fully verified.

## Acceptance and recovery

Render and compare **every** output slide with its source PDF page. Inspect every prose/formula junction, especially sums, integrals, fractions, subscripts, multiline text, and colored text. Check for missing symbols, tofu/empty-square glyphs in ordinary text (especially offline-OCR Chinese titles), bad crops/QR codes, clipping, overlapping boxes, glued word junctions, size drift, and baseline drift. Manually transcribe legible source text when offline hints omit it; never ship tiny placeholder text. Formula-only crops cannot reveal collisions with nearby prose. Resolve each checker warning or report it explicitly; a structural pass is not visual acceptance.

The final `.pptx` must open, match source page count/order, pass page/final runtime validation, contain the expected native formulas with **zero empty `m:nary/m:e` bodies or dotted placeholder squares**, and have no full-page source raster overlay. Open at least one formula in PowerPoint's equation editor. If a required formula remains an image, the deck is not a completed editable-formula delivery. Repair its source LaTeX/inventory and rerun from the preserved pre-equation deck; do not overwrite the PDF or successful intermediates. Report final file path, page/formula counts, QA result, non-editable image classes, and unresolved limitations.

When post-finalization QA finds a page defect, repair that page's manifest/assets, rebuild and validate it, then call `run record` again with the same page-owner agent id. This refreshes an accepted page's hashes and reopens the run for `run finalize`; do not hand-edit `page_jobs.json`. Keep the previous native/calibrated files as recovery and write each new conversion to a fresh path.

## Bundled resources

- `cli/`: bundled `editppt` runtime and Python package; install it from this skill folder.
- `prompts/page-worker.md`, `scripts/build-page-worker-prompt.py`: page ownership and prompt generation.
- `references/cli-helper.md`, `references/manifest-schema.md`, `references/page-decision-tree.md`: commands, page contracts, and object decisions.
- `references/reconstruction-and-qa.md`: source crops, formula inventory, calibration, and recovery.
- `scripts/equationize.py`, `scripts/calibrate_math_geometry.ps1`, `scripts/measure_math_geometry.py`, `scripts/check_formula_text_spacing.ps1`: native math conversion and visual QA. Legacy font-size scripts are diagnostic only, not final source-calibrated gates.

## 致谢

The bundled `editppt` runtime and core page-reconstruction workflow are adapted from [ningzimu/image-to-editable-ppt-skill](https://github.com/ningzimu/image-to-editable-ppt-skill), under its MIT license; retain [THIRD_PARTY_LICENSES.md](THIRD_PARTY_LICENSES.md) when distributing. [ningzimu/codex-ppt-skill](https://github.com/ningzimu/codex-ppt-skill) inspired phase boundaries and presentation QA; its image-only assembly code is not used here. This acknowledgement does not require users to install either upstream skill.
