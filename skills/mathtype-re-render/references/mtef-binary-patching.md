# MTEF binary configuration patching

Read this reference before changing `patch_mathtype_mtef.py`, adding a new EQP profile, supporting a new font drift, or diagnosing a binary-patch failure.

## First-principles model

A MathType formula in DOCX is not one object. The editable data is an OLE2 Compound File under `word/embeddings/*.bin`. Its named `Equation Native` stream contains the MTEF payload. The visual formula in Word is a separate WMF part. Therefore:

1. Patch `Equation Native` to change MathType configuration.
2. Prove that the patch is structurally and semantically correct.
3. Regenerate the WMF preview from the patched OLE.
4. Synchronize Word's object box to the regenerated WMF.

Changing the MTEF bytes alone cannot update the image Word currently shows.

## EQP to MTEF preference encoding

The supported `.eqp` file is a MathType Equation Preferences INI. Parse:

- 8 entries from `[Sizes]`;
- 30 entries from `[Spacing]`;
- base font names from `[Styles]`.

Exactly 38 numeric values are required. Each value is encoded as a nibble sequence:

```text
unit digit(s) F
```

- unit `2` means `pt`;
- unit `4` means `%`;
- digits `0..9` are literal decimal digits;
- nibble `A` is the decimal point;
- nibble `F` terminates the value;
- after the eighth value, `1E` separates Sizes from Spacing.

Some MathType writers end the eighth value on an odd nibble boundary and insert one `0` padding nibble before `1E`. Accept only these two forms:

```text
... F 1 E
... F 0 1 E   # only at the known odd-nibble alignment
```

Reject other markers or half-byte replacement boundaries. Do not guess alignment.

## Locating EQN_PREFS safely

Do not search the whole OLE file for `12 00`. Directory entries, unused sectors, and padding can contain false matches.

1. Parse the OLE2 container.
2. Read the named `Equation Native` stream.
3. Scan candidate `12 00` record markers.
4. Starting at `record + 3`, attempt to parse exactly 38 values and the required separator.
5. Accept the first candidate that round-trips as a complete preference block.

After replacement, parse the newly built stream again and require all 38 values to equal the EQP target before writing the OLE container.

## Font records

Font definitions before `EQN_PREFS` are length-coded legacy-code-page records. Font replacement is more fragile than numeric preference replacement. Apply only an exact, previously verified byte pattern.

The current conservative normalization supports a missing `宋体` TextFE definition when the equation contains one of these known exact drifts:

- `Arial -> 宋体`;
- `华文楷体 -> 宋体`.

Do not generalize this into arbitrary string substitution. A new target font or a new drift pattern requires:

1. a manually formatted MathType control equation;
2. before/after `Equation Native` comparison;
3. identification of the complete length-coded record, not only the font-name bytes;
4. tests for shorter, equal-length, and longer replacement records;
5. post-patch MathType editability and re-render validation.

If an EQP demands an unsupported font and the audit reports `missing_font`, stop. Do not claim full configuration compliance.

## OLE mini-stream rewriting

Most `Equation Native` streams are smaller than 4096 bytes and live in the OLE mini-stream. A one-byte change in the packed preference block may change the logical stream size and cross a 64-byte mini-sector boundary.

The patcher must keep these structures consistent:

- the `Equation Native` directory entry size;
- the MTEF payload-length field at stream bytes `8:12`, equal to `len(stream) - 28`;
- the miniFAT chain for the stream;
- the Root Entry's mini-stream logical size when a free mini-sector beyond the previous logical end is used.

The conservative allocator may append only free mini-sectors already covered by the Root Entry's physically allocated regular-sector chain. It does not grow the OLE file with new regular sectors. Stop with an explicit error when there is insufficient safe capacity or when `Equation Native` reaches the 4096-byte cutoff.

Never rewrite the OLE object through a generic library merely to make the operation easier. Container canonicalization can change unrelated streams or metadata and makes “only Equation Native changed” harder to prove.

## Patch integrity gates

Run `verify_mtef_patch.py` immediately after patching. Require all gates:

1. Source and patched DOCX have identical member names and order.
2. Equation object order, owner XML part, relationship ID, and OLE part mapping are identical.
3. Every non-target DOCX part is byte-identical.
4. Every non-`Equation Native` stream inside each MathType OLE is byte-identical.
5. Every patched `Equation Native` has a correct internal MTEF length.
6. All 38 values and required fonts pass the target EQP audit.
7. ZIP CRC passes.

The complete OLE container hash is expected to change after MTEF patching. After the later preview-refresh stage, however, OLE hashes must remain identical to the already-patched source because that stage is allowed to change only preview/geometry parts.

## Experience from the validated full-document run

The verified M19 run contained 5,758 Equation.DSMT4 objects. Preference replacement sometimes preserved stream size and sometimes grew it by one byte. Boundary-focused tests covered equations whose replacement required mini-stream-chain extension. The successful workflow used:

- deterministic EQP parsing;
- parse-and-round-trip `EQN_PREFS` detection;
- conservative mini-sector allocation;
- full post-patch EQP audit;
- unique-preview splitting before refresh;
- x86 OLE rendering in batches of 25;
- final `5758/5758` WMF/Word geometry consistency.

The key lesson is that binary correctness and visual correctness are separate gates. Passing only one is insufficient.
