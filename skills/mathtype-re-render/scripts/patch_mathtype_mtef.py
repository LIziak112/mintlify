"""Patch MathType preference/font records while preserving the DOCX container.

The script only touches the ``Equation Native`` stream in Equation.DSMT4 OLE
parts whose preferences differ from the supplied EPP/EQP file.  It does not
touch document.xml, relationships, WMF previews, or non-MathType OLEs.

MathType stores preferences as a nibble stream and font names as length-coded
records. A replacement may be a few bytes longer or shorter than the old
records; the embedded stream is rewritten in place, extending the existing
mini-stream chain when spare sectors are already present in the OLE container.
"""
from __future__ import annotations

import argparse
import json
import os
import shutil
import tempfile
import zipfile
from pathlib import Path

from cfb_ole import CompoundFile, ENDOFCHAIN, FREESECT, NOSTREAM
from equation_parts import enumerate_equations
from find_deviating import parse_eqp, parse_eqp_fonts, read_value, unpack_values


def expected_nibbles(values: list[tuple[str, str]]) -> list[int]:
    out: list[int] = []
    for index, (text, unit) in enumerate(values):
        out.append(2 if unit == "pt" else 4)
        for char in text:
            out.append(10 if char == "." else int(char))
        out.append(15)
        if index == 7:
            out.extend((1, 14))
    return out


def pack_nibbles(nibbles: list[int]) -> bytes:
    if len(nibbles) % 2:
        raise ValueError("preference nibble stream is not byte aligned")
    return bytes((nibbles[i] << 4) | nibbles[i + 1] for i in range(0, len(nibbles), 2))


def preference_span(payload: bytes, count: int) -> int:
    """Return the number of payload nibbles occupied by the preference block."""
    nibbles: list[int] = []
    for value in payload:
        nibbles.extend((value >> 4, value & 15))
    pos = 0
    for _ in range(8):
        item, pos = read_value(nibbles, pos)
        if item is None:
            raise ValueError("invalid size preference")
    if nibbles[pos : pos + 2] == [1, 14]:
        pos += 2
    elif pos % 2 == 1 and nibbles[pos : pos + 3] == [0, 1, 14]:
        # Preserve the writer's single zero pad nibble in the old span. The
        # replacement payload is canonical and does not need this pad.
        pos += 3
    else:
        raise ValueError("missing MathType spacing marker")
    for _ in range(8, count):
        item, pos = read_value(nibbles, pos)
        if item is None:
            raise ValueError("invalid spacing preference")
    return pos


def locate_preferences(native: bytes, count: int) -> tuple[int, int, list[tuple[str, str]]]:
    """Find the parseable EQN_PREFS record in Equation Native."""
    marker = bytes((0x12, 0x00))
    start = 0
    while True:
        record = native.find(marker, start)
        if record < 0:
            raise ValueError("EQN_PREFS record not found")
        try:
            span = preference_span(native[record + 3 :], count)
            values = unpack_values(native[record + 3 :], count)
            if values is not None:
                return record, span, values
        except (KeyError, ValueError, IndexError):
            pass
        start = record + 2


def replace_preferences(native: bytes, expected: list[tuple[str, str]]) -> tuple[bytes, list[tuple[str, str]]]:
    record, old_span, old_values = locate_preferences(native, len(expected))
    if old_values == expected:
        return native, old_values
    if old_span % 2:
        raise ValueError("unhandled half-byte preference boundary")
    new_payload = pack_nibbles(expected_nibbles(expected))
    old_payload_bytes = old_span // 2
    begin = record + 3
    replaced = native[:begin] + new_payload + native[begin + old_payload_bytes :]
    _, _, check = locate_preferences(replaced, len(expected))
    if check != expected:
        raise ValueError(f"preference replacement verification failed: {check!r}")
    return replaced, old_values


def normalize_fonts(native: bytes, expected_fonts: set[bytes]) -> tuple[bytes, str | None]:
    """Repair known TextFE font-definition drifts to 宋体.

    MathType font definitions are length-prefixed records.  The target EQP
    names are stored in legacy code pages; M13's lone drift contains Arial in
    the TextFE slot and has no Song font definition.  Replace only that exact
    record, leaving all other font definitions untouched.
    """
    song = "宋体".encode("cp936")
    right = b"\x00\x11\x06" + song + b"\x00"
    record, _, _ = locate_preferences(native, 38)
    region = native[:record]
    if song in region or song not in expected_fonts:
        return native, None
    known_drifts = (
        (b"\x00\x11\x05Arial\x00", "Arial->宋体"),
        (b"\x00\x11\x06" + "华文楷体".encode("cp936") + b"\x00", "华文楷体->宋体"),
    )
    for wrong, label in known_drifts:
        if wrong in region:
            return native.replace(wrong, right, 1), label
    return native, None


def _write_chain_at(cfb: CompoundFile, out: bytearray, start: int, table: list[int], offset: int, payload: bytes) -> None:
    """Write bytes into a regular-sector chain already allocated in the CFB."""
    remaining = memoryview(payload)
    position = offset
    chain = cfb.chain(start, table)
    while remaining:
        sector_index, within = divmod(position, cfb.sector_size)
        if sector_index >= len(chain):
            raise ValueError("directory stream capacity exceeded")
        sid = chain[sector_index]
        physical = (sid + 1) * cfb.sector_size + within
        take = min(len(remaining), cfb.sector_size - within)
        out[physical : physical + take] = remaining[:take]
        remaining = remaining[take:]
        position += take


def _write_regular_at(cfb: CompoundFile, out: bytearray, offset: int, payload: bytes) -> None:
    _write_chain_at(cfb, out, cfb.first_dir, cfb.fat, offset, payload)


def patch_cfb_stream(cfb: CompoundFile, original: bytes, entry, new_stream: bytes) -> bytes:
    """Return a CFB byte array with one stream replaced in-place."""
    if len(new_stream) > 4096:
        raise ValueError("this conservative patcher only handles mini-streams")
    if len(new_stream) < 28:
        raise ValueError("Equation Native stream is shorter than the MTEF header")
    # The MTEF header stores the payload length (native stream length minus its
    # fixed 28-byte header). Keep that internal length consistent when a
    # preference replacement changes the stream size.
    header_length = int.from_bytes(new_stream[8:12], "little")
    expected_length = len(new_stream) - 28
    if header_length != expected_length:
        candidate = bytearray(new_stream)
        candidate[8:12] = expected_length.to_bytes(4, "little")
        new_stream = bytes(candidate)
    out = bytearray(original)
    mini_table = list(cfb.minifat)
    mini_chain = cfb.chain(entry.start, mini_table)
    needed = (len(new_stream) + cfb.mini_sector_size - 1) // cfb.mini_sector_size
    if needed > len(mini_chain):
        # Prefer free mini-sectors that are physically covered by the Root
        # Entry's already allocated regular-sector chain. The logical root
        # size may be smaller than that physical capacity, so it is extended
        # below when a selected sector lies just beyond the old size.
        root_chain = cfb.chain(cfb.root.start, cfb.fat)
        root_capacity = len(root_chain) * cfb.sector_size
        mini_slots = root_capacity // cfb.mini_sector_size
        used = set()
        for stream in cfb.stream_entries():
            if stream.size < 4096 and stream.start not in (FREESECT, ENDOFCHAIN, NOSTREAM):
                used.update(cfb.chain(stream.start, mini_table))
        free = [
            sid
            for sid in range(min(len(mini_table), mini_slots))
            if mini_table[sid] == FREESECT and sid not in used
        ]
        extra = needed - len(mini_chain)
        if len(free) < extra:
            raise ValueError("not enough free mini-sectors for stream growth")
        appended = free[:extra]
        old_last = mini_chain[-1]
        for left, right in zip([old_last] + appended[:-1], appended):
            mini_table[left] = right
        mini_table[appended[-1]] = ENDOFCHAIN
        mini_chain = cfb.chain(entry.start, mini_table)
        required_root_size = max(cfb.root.size, (max(mini_chain) + 1) * cfb.mini_sector_size)
        if required_root_size > root_capacity:
            raise ValueError("mini-stream root needs additional regular sectors")
        # Persist MiniFAT links and the Root Entry logical size.
        for sid in [old_last] + appended:
            _write_chain_at(cfb, out, cfb.first_minifat, cfb.fat, sid * 4, mini_table[sid].to_bytes(4, "little"))
        root_size_offset = cfb.root.index * 128 + 0x78
        _write_regular_at(cfb, out, root_size_offset, required_root_size.to_bytes(8, "little"))
        cfb.minifat = mini_table
        cfb.root.size = required_root_size
    capacity = len(mini_chain) * cfb.mini_sector_size
    if len(new_stream) > capacity:
        raise ValueError(f"stream needs {len(new_stream)} bytes, capacity is {capacity}")
    root_chain = cfb.chain(cfb.root.start, cfb.fat)
    # Update the mini-stream bytes occupied by this entry.  The mini-stream is
    # itself a regular stream rooted at the Root Entry.
    for pos in range(len(new_stream)):
        mini_pos = mini_chain[pos // cfb.mini_sector_size] * cfb.mini_sector_size + (pos % cfb.mini_sector_size)
        regular_index, within = divmod(mini_pos, cfb.sector_size)
        if regular_index >= len(root_chain):
            raise ValueError("mini-stream capacity exceeded")
        physical = (root_chain[regular_index] + 1) * cfb.sector_size + within
        out[physical] = new_stream[pos]
    # Clear stale bytes if a future replacement is shorter.  Current target
    # replacements grow, but clearing makes this helper deterministic.
    for pos in range(len(new_stream), min(capacity, len(new_stream) + 16)):
        mini_pos = mini_chain[pos // cfb.mini_sector_size] * cfb.mini_sector_size + (pos % cfb.mini_sector_size)
        regular_index, within = divmod(mini_pos, cfb.sector_size)
        physical = (root_chain[regular_index] + 1) * cfb.sector_size + within
        out[physical] = 0
    # The directory stream records the logical stream length at +0x78.
    directory_offset = entry.index * 128 + 0x78
    _write_regular_at(cfb, out, directory_offset, len(new_stream).to_bytes(8, "little"))
    return bytes(out)


def patch_docx(source: Path, output: Path, eqp: Path) -> dict:
    source = source.resolve()
    output = output.resolve()
    if source == output:
        raise ValueError("source and output must be different files")
    if not source.is_file() or source.suffix.lower() != ".docx":
        raise ValueError(f"source must be an existing DOCX: {source}")
    if not eqp.is_file():
        raise ValueError(f"EQP does not exist: {eqp}")
    expected = [(value, unit) for _, value, unit in parse_eqp(eqp)]
    expected_fonts = parse_eqp_fonts(eqp)
    if len(expected) != 38:
        raise ValueError(f"expected 38 MathType values, got {len(expected)}")
    output.parent.mkdir(parents=True, exist_ok=True)
    changed: list[dict] = []
    unchanged_math: list[str] = []
    with zipfile.ZipFile(source, "r") as zin:
        equations = enumerate_equations(zin)
        target_by_part: dict[str, dict[str, object]] = {}
        for equation in equations:
            part = str(equation["ole_part"])
            if part in target_by_part:
                raise ValueError(f"OLE part is referenced by multiple Equation.DSMT4 objects: {part}")
            target_by_part[part] = equation
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx", dir=str(output.parent)) as tmp:
            temp_path = Path(tmp.name)
        try:
            with zipfile.ZipFile(temp_path, "w") as zout:
                for info in zin.infolist():
                    data = zin.read(info.filename)
                    equation = target_by_part.get(info.filename)
                    if equation:
                        part = info.filename
                        doc_index = int(equation["index"])
                        cfb = CompoundFile(data)
                        entry = cfb.find_stream("Equation Native")
                        if entry is None:
                            raise ValueError(f"{part}: Equation Native stream missing")
                        native = cfb.read_entry(entry)
                        replaced_native, old_values = replace_preferences(native, expected)
                        replaced_native, font_change = normalize_fonts(replaced_native, expected_fonts)
                        if replaced_native != native:
                            data = patch_cfb_stream(cfb, data, entry, replaced_native)
                            changed.append({
                                "index": doc_index,
                                "owner_part": equation["owner_part"],
                                "part": part,
                                "old_full_size": old_values[0] if old_values else None,
                                "old_stream_size": len(native),
                                "new_stream_size": len(replaced_native),
                                "font_change": font_change,
                            })
                        else:
                            unchanged_math.append(part)
                    # Preserve every non-target package part byte-for-byte at
                    # the logical level, including compression type and names.
                    zout.writestr(info, data)
            with zipfile.ZipFile(temp_path) as check:
                bad = check.testzip()
                if bad:
                    raise ValueError(f"ZIP CRC failure after patch: {bad}")
            os.replace(temp_path, output)
        finally:
            if temp_path.exists():
                temp_path.unlink()
    return {
        "schema_version": 2,
        "source": str(source),
        "output": str(output),
        "eqp": str(eqp.resolve()),
        "math_total": len(equations),
        "changed": changed,
        "unchanged": unchanged_math,
    }


def main() -> None:
    ap = argparse.ArgumentParser(description="Patch Equation.DSMT4 MTEF preferences in a DOCX copy.")
    ap.add_argument("eqp", type=Path)
    ap.add_argument("source", type=Path)
    ap.add_argument("output", type=Path)
    ap.add_argument("--report", type=Path)
    ap.add_argument("--overwrite", action="store_true")
    args = ap.parse_args()
    if args.output.exists() and not args.overwrite:
        raise ValueError(f"output exists; use --overwrite: {args.output}")
    result = patch_docx(args.source, args.output, args.eqp)
    print(f"SOURCE={result['source']}")
    print(f"OUTPUT={result['output']}")
    print(f"MATHTYPE_TOTAL={result['math_total']}")
    print(f"CHANGED={len(result['changed'])}")
    for row in result["changed"][:50]:
        print(f"{row['index']}\t{row['part']}\t{row['old_full_size']}\t{row['old_stream_size']}->{row['new_stream_size']}")
    if len(result["changed"]) > 50:
        print(f"... {len(result['changed']) - 50} more; use --report for the full list")
    if args.report:
        args.report.parent.mkdir(parents=True, exist_ok=True)
        args.report.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
