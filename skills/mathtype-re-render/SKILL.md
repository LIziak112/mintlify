---
name: mathtype-re-render
description: "Run a complete, copy-first DOCX MathType repair pipeline: audit an EQP profile, patch the 38-value MTEF EQN_PREFS/font records inside editable Equation.DSMT4 OLE Equation Native streams, prove non-target bytes and OLE streams are unchanged, render patched OLE objects to placeable WMF with a 32-bit worker, synchronize Word geometry, and audit the final package. Use when MathType binary configuration and displayed formula previews must both be updated without losing editability."
---

# Mathtype Re-render

Apply the full chain:

```text
EQP profile -> MTEF binary audit/patch -> binary integrity proof
            -> unique WMF mapping -> x86 OLE rendering
            -> WMF replacement + Word geometry -> final audit/render QA
```

Never treat “MTEF has been changed” as equivalent to “Word display has been updated.” They are separate layers and separate acceptance gates.

## Required inputs and environment

- Use a `.docx` containing editable `Equation.DSMT4` objects.
- Use a MathType `.eqp` profile. Prefer the user's file; otherwise select an explicitly matching file from `assets/presets/`.
- Use Windows with MathType's 32-bit OLE server installed and registered.
- Always write to a new DOCX. Refuse source=output and do not overwrite without explicit authorization.
- Do not close unrelated Word/MathType instances. This pipeline does not need Word for formula rendering.

## Preferred one-command workflow

Run from `scripts/` or use absolute paths:

```powershell
python .\run_full_pipeline.py target.eqp source.docx refreshed.docx `
  --work-dir .\work\job-name `
  --cache .\cache `
  --batch-size 25
```

The runner performs all mandatory structural stages and stops on any failed gate. Add `--overwrite` only when the user explicitly authorizes replacing an existing output/intermediate.

After the runner succeeds, render the final DOCX and inspect all pages using the documents skill. Structural success alone is not a visual QA substitute.

## Complete workflow and gates

### 1. Preflight the EQP and source

```powershell
python .\find_deviating.py target.eqp source.docx --out .\work\source-mtef-audit.json
```

Require exactly 38 numeric settings: 8 `[Sizes]` and 30 `[Spacing]`. Record all Equation.DSMT4 objects across every internal `word/*.xml` owner part—not only `word/document.xml`. A preflight exit code of `2` means deviations exist and is expected when repair is needed.

Read [references/mtef-binary-patching.md](references/mtef-binary-patching.md) before changing the patcher, supporting a new font drift/profile format, or diagnosing a binary failure.

### 2. Patch Equation Native conservatively

```powershell
python .\patch_mathtype_mtef.py target.eqp source.docx patched.docx `
  --report .\work\mtef-patch-report.json
```

Patch only the named `Equation Native` stream inside mapped Equation.DSMT4 OLE parts:

- locate a parseable `EQN_PREFS` record, not merely the first `12 00` bytes;
- replace the nibble-packed 38-value block;
- round-trip parse the replacement before publishing it;
- normalize only exact, verified font records;
- update the MTEF internal length and OLE mini-stream metadata when the stream size changes;
- preserve every document XML, relationship, preview and non-MathType part at this stage.

Stop rather than guess when preferences are unparseable, the replacement ends on an unsupported half-byte boundary, the target font is unsupported, `Equation Native` is not a mini-stream, or safe mini-sector capacity is insufficient.

### 3. Prove the binary patch

```powershell
python .\verify_mtef_patch.py target.eqp source.docx patched.docx `
  --report .\work\mtef-patch-verification.json
python .\find_deviating.py target.eqp patched.docx --out .\work\patched-mtef-audit.json
```

Do not proceed unless:

- MTEF deviations are zero;
- Equation object/relationship mapping is unchanged;
- every non-target DOCX part is byte-identical;
- every non-`Equation Native` OLE stream is byte-identical;
- each MTEF internal length equals `stream length - 28`;
- ZIP CRC passes.

This is the missing proof between “some bytes changed” and “the MathType configuration was safely changed.”

### 4. Map formula OLE objects to previews

```powershell
python .\ole-preview-bridge\map_equations.py patched.docx .\work\equation-map.json
```

Map by OOXML relationship IDs. Never infer that `oleObject17.bin` corresponds to `image17.wmf` from numeric suffixes.

If `unique_preview_parts < equation_count`, split shared preview relationships:

```powershell
python .\ole-preview-bridge\split_shared_previews.py `
  patched.docx .\work\equation-map.json unique-preview-source.docx
python .\ole-preview-bridge\map_equations.py `
  unique-preview-source.docx .\work\unique-equation-map.json
python .\ole-preview-bridge\verify_equation_map.py `
  unique-preview-source.docx .\work\unique-equation-map.json
```

Splitting is mandatory when different equations share one WMF; otherwise one rendered formula would overwrite another object's display. It may also split safe same-equation sharing for deterministic one-object/one-preview replacement.

### 5. Render patched OLE to WMF

Build the included x86 worker once per machine or after source changes:

```powershell
powershell -ExecutionPolicy Bypass -File .\ole-preview-bridge\build_renderer32.ps1
```

Then refresh:

```powershell
python .\ole-preview-bridge\refresh_previews.py `
  unique-preview-source.docx refreshed.docx `
  --manifest .\work\unique-equation-map.json `
  --renderer .\ole-preview-bridge\renderer32\OlePreviewRenderer.exe `
  --cache .\cache `
  --batch-size 25
```

The STA worker opens the patched OLE with Windows OLE, calls `OleRun` and `IOleObject.Update`, obtains `CF_METAFILEPICT`, and emits a validated placeable WMF. Restart it after 25 cache misses by default to release MathType memory. Lower the batch size if the OLE server becomes unstable.

The refresh stage must not change OLE bytes. It may change only mapped WMFs, necessary owner XML geometry, relationships introduced by preview splitting, and package metadata resulting from repacking.

### 6. Synchronize and audit geometry

For every formula, derive physical width/height from the placeable WMF header and synchronize:

- `w:dxaOrig` and `w:dyaOrig` in twentieths of a point;
- VML `width` and `height`, quantized to 0.75 pt like Word/MathType;
- preserve baseline `w:position` rather than vertically recentering formulas.

Run:

```powershell
python .\ole-preview-bridge\audit_geometry.py refreshed.docx
```

Require `failures=0`, no missing VML dimensions, valid WMFs, unchanged patched OLE hashes, and a valid DOCX CRC.

### 7. Render the whole document

Use the documents skill's `render_docx.py`, then inspect every page at 100%. Check for missing formulas, stale preview boxes, compression, stretch, clipping, overlap, baseline drift and page-flow changes. Do not alter mathematical content merely because the source formula itself looks unusual.

## Shortest safe variants

- Configuration wrong and preview stale: run the entire pipeline.
- Configuration already patched and verified: start with relationship mapping and OLE re-render; do not patch again.
- Only previews stale and EQP configuration is already compliant: skip patching, but still map, split, refresh and audit.
- Unknown/new font mapping: stop after audit and build a control-equation byte comparison before extending the patcher.
- Renderer unavailable: diagnose x86 build and MathType registration. Never silently replace editable equations with PNG.

## Bundled resources

- `scripts/run_full_pipeline.py`: complete structural pipeline entry point.
- `scripts/equation_parts.py`: enumerate Equation.DSMT4 objects across Word XML owner parts.
- `scripts/find_deviating.py`: EQP-versus-MTEF audit with JSON output.
- `scripts/patch_mathtype_mtef.py`: conservative nibble/font/mini-stream patcher.
- `scripts/verify_mtef_patch.py`: proof that only allowed Equation Native streams changed.
- `scripts/cfb_ole.py`: dependency-free OLE2 reader used by audit and patch code.
- `scripts/ole-preview-bridge/`: mapper, verifier, shared-preview splitter, x86 worker, refresh engine and geometry audit.
- `references/mtef-binary-patching.md`: binary encoding, allocation, validation and extension rules.
- `assets/presets/`: adapted profiles for 教辅格式11.7pm、正文11.9pt、一级标题16pm and 二级标题13pm.

## Deliverable

Return the final refreshed DOCX, not intermediates, unless requested. Report the selected EQP, equation count, MTEF changed/unchanged count, shared-preview splits, rendered/cache-hit counts, geometry result, OLE-preservation result and full-page QA status. Keep JSON reports in the work directory for reproducibility.
