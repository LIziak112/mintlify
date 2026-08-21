"""Audit MathType MTEF preferences without opening Word or MathType."""
from __future__ import annotations

import argparse
import json
import re
import zipfile
from pathlib import Path

from cfb_ole import CompoundFile
from equation_parts import enumerate_equations


def decode_eqp(path: Path) -> str:
    raw = path.read_bytes()
    try:
        return raw.decode("utf-8-sig")
    except UnicodeDecodeError:
        return raw.decode("cp936")


def parse_eqp(path: Path) -> list[tuple[str, str, str]]:
    values: list[tuple[str, str, str]] = []
    section = ""
    for raw in decode_eqp(path).splitlines():
        line = raw.strip()
        if line.startswith("[") and line.endswith("]"):
            section = line[1:-1]
        elif "=" in line and section in {"Sizes", "Spacing"}:
            key, value = (item.strip() for item in line.split("=", 1))
            match = re.fullmatch(r"([0-9]+(?:\.[0-9]+)?)\s*(pt|%)", value)
            if match:
                values.append((f"{section}:{key}", match.group(1), match.group(2)))
    return values


def parse_eqp_fonts(path: Path) -> set[bytes]:
    fonts: set[bytes] = set()
    section = ""
    for raw in decode_eqp(path).splitlines():
        line = raw.strip()
        if line.startswith("[") and line.endswith("]"):
            section = line[1:-1]
        elif section == "Styles" and "=" in line and not line.startswith(";"):
            base = line.split("=", 1)[1].strip().split(",", 1)[0]
            if base:
                fonts.add(base.encode("cp936", "replace"))
    return fonts


def read_value(nibbles: list[int], position: int) -> tuple[tuple[str, str] | None, int]:
    if position >= len(nibbles) or nibbles[position] not in (2, 4):
        return None, position
    unit = {2: "pt", 4: "%"}[nibbles[position]]
    position += 1
    text = ""
    while position < len(nibbles) and nibbles[position] != 15:
        nibble = nibbles[position]
        if nibble <= 9:
            text += str(nibble)
        elif nibble == 10:
            text += "."
        else:
            return None, position
        position += 1
    if position >= len(nibbles) or not text:
        return None, position
    return (text, unit), position + 1


def unpack_values(payload: bytes, count: int) -> list[tuple[str, str]] | None:
    nibbles = [nibble for byte in payload for nibble in (byte >> 4, byte & 15)]
    values: list[tuple[str, str]] = []
    position = 0
    while len(values) < 8:
        item, position = read_value(nibbles, position)
        if item is None:
            return None
        values.append(item)
    if nibbles[position : position + 2] == [1, 14]:
        position += 2
    elif position % 2 == 1 and nibbles[position : position + 3] == [0, 1, 14]:
        position += 3
    else:
        return None
    while len(values) < count:
        item, position = read_value(nibbles, position)
        if item is None:
            return None
        values.append(item)
    return values


def locate_prefs(native: bytes, count: int) -> tuple[int, list[tuple[str, str]]] | None:
    start = 0
    while True:
        record = native.find(b"\x12\x00", start)
        if record < 0:
            return None
        values = unpack_values(native[record + 3 :], count)
        if values is not None:
            return record, values
        start = record + 2


def ole_prefs(data: bytes, expected_count: int, expected_fonts: set[bytes]) -> tuple[list[tuple[str, str]] | None, list[str]]:
    try:
        compound = CompoundFile(data)
        entry = compound.find_stream("Equation Native")
        if entry is None:
            return None, ["missing_equation_native"]
        native = compound.read_entry(entry)
    except (ValueError, IndexError, KeyError, StopIteration):
        return None, ["invalid_ole_container"]
    located = locate_prefs(native, expected_count)
    if located is None:
        return None, ["missing_or_unparseable_eqn_prefs"]
    record, preferences = located
    region = native[:record]
    issues = [
        "missing_font:" + font.decode("cp936", "replace")
        for font in expected_fonts
        if font not in region
    ]
    return preferences, issues


def audit(docx: Path, eqp: Path) -> tuple[list[dict[str, object]], list[tuple[str, str]]]:
    expected = [(value, unit) for _, value, unit in parse_eqp(eqp)]
    if len(expected) != 38:
        raise ValueError(f"expected 38 MathType values in EQP, got {len(expected)}")
    expected_fonts = parse_eqp_fonts(eqp)
    rows: list[dict[str, object]] = []
    with zipfile.ZipFile(docx) as archive:
        equations = enumerate_equations(archive)
        for equation in equations:
            part = str(equation["ole_part"])
            preferences, issues = ole_prefs(archive.read(part), len(expected), expected_fonts)
            if preferences is None:
                issues.append("unparseable_prefs")
            elif preferences != expected:
                issues.append("prefs_mismatch")
            rows.append({**equation, "prefs": preferences, "issues": issues})
    return rows, expected


def main() -> int:
    parser = argparse.ArgumentParser(description="Audit Equation.DSMT4 MTEF values against an EQP profile.")
    parser.add_argument("eqp", type=Path)
    parser.add_argument("docx", type=Path)
    parser.add_argument("--out", type=Path, help="Write a JSON audit report")
    args = parser.parse_args()
    rows, expected = audit(args.docx.resolve(), args.eqp.resolve())
    deviations = [row for row in rows if row["issues"]]
    report = {
        "schema_version": 2,
        "docx": str(args.docx.resolve()),
        "eqp": str(args.eqp.resolve()),
        "equation_count": len(rows),
        "expected_value_count": len(expected),
        "deviation_count": len(deviations),
        "deviations": deviations,
    }
    print(f"DOCX={report['docx']}")
    print(f"TOTAL={len(rows)}")
    print(f"EXPECTED_VALUES={len(expected)}")
    print(f"DEVIATING={len(deviations)}")
    for row in deviations[:50]:
        print(f"{row['index']}\t{row['owner_part']}\t{row['ole_part']}\t{','.join(row['issues'])}")
    if len(deviations) > 50:
        print(f"... {len(deviations) - 50} more; use --out for the full report")
    if args.out:
        args.out.parent.mkdir(parents=True, exist_ok=True)
        args.out.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return 0 if not deviations else 2


if __name__ == "__main__":
    raise SystemExit(main())
