# Page Reconstructor Prompt Template

Placeholders of the form `{{NAME}}` are filled by `scripts/build-page-worker-prompt.py`.

```text
Rebuild one page for ppt-rebuild, preserving formula inventory for native Office Math conversion after deck finalization.

Run dir: {{RUN_DIR}}
Page id: {{PAGE_ID}}
Page dir: {{PAGE_DIR}}
Source image: {{SOURCE_IMAGE}}
Bundled runtime command: {{EDITPPT_CMD}}

You own only this Page dir. Do not edit deck_manifest.json, page_jobs.json, notes_manifest.json, final outputs, the original input, or any other page directory.

Read references by task, before acting on their rules. All reference paths below are under {{SKILL_ROOT}}/references/; do not reread unchanged sections already available in this execution.
- Start with page_request.json and page-decision-tree.md's introduction, "Common Failure Mode: False Progress", and "Pre-Decision: Page Inventory". Inspect the source and build the inventory, then read sections 1 and 2 for object-source decisions.
- Before native reconstruction, read section 3.1 for text; read sections 3.2-3.6 when the inventory contains their formulas, structural objects/tables, corners, decorated text, or groups. Read "Final Self-Check" and "Fix versus Warning" before acceptance or a repair decision.
- Read manifest-schema.md's page_request.json and pages/page_NNN/manifest.json contracts before writing the manifest, loading applicable object contracts such as "Native tables". Before image jobs read the image_backend and imagegen-jobs.json contracts; for region splitting read "Asset sheet regions and split reports". Before returning read the validation.json and page_result.json contracts. Deck-state and notes contracts are outside your ownership.
- Use cli-helper.md for the commands needed now: "Page Build Commands", and "Image Backend Commands", "Asset Processing Commands", or "Formula Commands" when applicable. Do not load unrelated install or orchestration instructions.

Hard rules (reminders; authoritative details remain in the references):
1. Separate foreground visuals through image editing, except user-requested exact PDF figure/logo/QR crops with documented `user-provided` provenance — page-decision-tree.md section 2. Never use a whole-page crop as a shortcut.
2. Decide background, then foreground, then native elements; consume text hints only after steps 1-2 decisions — page-decision-tree.md introduction and section 3.1.
3. Build from manifest.json with the deterministic runtime; preserve source-pixel geometry and request canvas — manifest-schema.md's page_request.json and manifest.json contracts.
4. Execute the recorded image backend contract, including built-in first, valid result import, and permitted fallback events — manifest-schema.md's image_backend contract.
5. Structural validation never waives object-source rules or the later native-equation conversion — page-decision-tree.md "Common Failure Mode: False Progress".
6. Keep uninterrupted prose as one editable box or intentional line boxes, never one box per OCR word. For mixed prose/math lines, split only at each formula box and reserve a visible gap on both sides; fit each prose segment to its source-pixel span. Microsoft YaHei is often wider than the PDF's serif face: adjust segment widths and point size only enough to match source line breaks **and visible glyph height**; shrinking prose to tiny text is also a page failure. Do not let a full-line text box extend behind a formula crop, highlight, or figure. Compare the rendered PowerPoint preview, not just preview.png; any visible text/formula collision is a page failure even when structural validation passes.
7. Give spatially separate math blocks on the same row separate formula ids/crops/boxes; do not use TeX spacing commands to position a far-right condition. Offline text hints may omit Chinese characters: transcribe legible source text manually in Microsoft YaHei, and reject tiny placeholder prose or empty-square glyphs in the preview.

Recovery: read any previous validation failure before editing. Verify reusable artifacts against the current source, page request/backend contract, manifest links, and imagegen-jobs.json provenance (paths/hashes); inspect their visual content where relevant. Reuse compliant artifacts and repair only failed or dependent parts. Rebuild from the inventory only when the source or object-source decisions are invalid. Never flip leftover validation to passed or return stale outputs without validating the current artifact set.

Execution:
1. Record the inventory and background/foreground decisions, then execute necessary image jobs under page-decision-tree.md sections 1-2. Page-local image jobs remain serial. Import and process selected image outputs using the image-job contract.
2. Reconstruct native elements under the applicable section 3 rules and write manifest.json using its field contract.
3. Run `{{EDITPPT_CMD}} page build {{PAGE_DIR}}`, then `{{EDITPPT_CMD}} page contact-sheet {{PAGE_DIR}}`.
4. Perform the Final Self-Check against the source and run `{{EDITPPT_CMD}} page validate {{PAGE_DIR}}`. Fix page-local issues yourself; after a change rebuild the affected outputs and verify the changed result. Once the current outputs pass, return without another unchanged build/QA cycle.

Required outputs: manifest.json, imagegen-jobs.json, page.pptx, preview.png, split_assets_contact.png, validation.json, page_result.json. Validation and page-result shapes are owned by manifest-schema.md; success requires top-level validation `passed: true` and the required outputs.

Routine repairs and recorded backend fallbacks require no extra user confirmation. Respect runtime-required approvals; the parent owns the OCR confirmation exception. Do not retry the same failing tool with unchanged inputs and conditions. If a hard requirement remains blocked after applicable fallbacks, write validation.json with `passed: false` and the concrete cause/error, plus page_result.json referencing only artifacts that exist. Preserve progress; do not fabricate outputs, substitute an approximate page, or request an unconditional fresh reconstruction.

Return only available artifact paths (omit missing artifacts on failure):
page_manifest=`<absolute path>`
page_pptx=`<absolute path>`
preview=`<absolute path>`
contact_sheet=`<absolute path>`
validation=`<absolute path>`
page_result=`<absolute path>`
```
