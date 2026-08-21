"""Prove that a DOCX MTEF patch changed only Equation Native streams."""
from __future__ import annotations

import argparse
import hashlib
import json
import zipfile
from pathlib import Path

from cfb_ole import CompoundFile
from equation_parts import enumerate_equations
from find_deviating import audit


def sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest().upper()


def stream_map(data: bytes) -> dict[str, bytes]:
    compound = CompoundFile(data)
    return {entry.name: compound.read_entry(entry) for entry in compound.stream_entries()}


def verify(source: Path, patched: Path, eqp: Path) -> dict[str, object]:
    rows, _ = audit(patched, eqp)
    deviations = [row for row in rows if row["issues"]]
    if deviations:
        raise ValueError(f"patched DOCX still has {len(deviations)} MTEF deviations")

    changed_native = 0
    unchanged_native = 0
    with zipfile.ZipFile(source) as left, zipfile.ZipFile(patched) as right:
        left_names = left.namelist()
        right_names = right.namelist()
        if left_names != right_names:
            raise ValueError("DOCX member names/order changed during MTEF patch")
        left_equations = enumerate_equations(left)
        right_equations = enumerate_equations(right)
        left_signature = [
            (item["owner_part"], item["object_number_in_part"], item["relationship_id"], item["ole_part"])
            for item in left_equations
        ]
        right_signature = [
            (item["owner_part"], item["object_number_in_part"], item["relationship_id"], item["ole_part"])
            for item in right_equations
        ]
        if left_signature != right_signature:
            raise ValueError("Equation.DSMT4 object/relationship mapping changed")
        target_parts = {str(item["ole_part"]) for item in left_equations}
        for name in left_names:
            before = left.read(name)
            after = right.read(name)
            if name not in target_parts:
                if before != after:
                    raise ValueError(f"non-target DOCX part changed: {name}")
                continue
            before_streams = stream_map(before)
            after_streams = stream_map(after)
            if before_streams.keys() != after_streams.keys():
                raise ValueError(f"OLE stream set changed: {name}")
            if "Equation Native" not in before_streams:
                raise ValueError(f"Equation Native missing: {name}")
            for stream_name in before_streams:
                if stream_name != "Equation Native" and before_streams[stream_name] != after_streams[stream_name]:
                    raise ValueError(f"non-MTEF OLE stream changed: {name}/{stream_name}")
            native = after_streams["Equation Native"]
            if len(native) < 28 or int.from_bytes(native[8:12], "little") != len(native) - 28:
                raise ValueError(f"MTEF internal length is invalid: {name}")
            if before_streams["Equation Native"] == native:
                unchanged_native += 1
            else:
                changed_native += 1
        bad = right.testzip()
        if bad:
            raise ValueError(f"ZIP CRC failure: {bad}")
    return {
        "status": "ok",
        "source": str(source),
        "patched": str(patched),
        "eqp": str(eqp),
        "equation_count": len(rows),
        "changed_equation_native": changed_native,
        "unchanged_equation_native": unchanged_native,
        "mtef_deviations": 0,
        "non_target_parts_unchanged": True,
        "non_equation_native_ole_streams_unchanged": True,
        "zip_crc_ok": True,
        "source_sha256": sha256(source.read_bytes()),
        "patched_sha256": sha256(patched.read_bytes()),
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Verify a conservative DOCX MathType MTEF patch.")
    parser.add_argument("eqp", type=Path)
    parser.add_argument("source", type=Path)
    parser.add_argument("patched", type=Path)
    parser.add_argument("--report", type=Path)
    args = parser.parse_args()
    result = verify(args.source.resolve(), args.patched.resolve(), args.eqp.resolve())
    for key, value in result.items():
        print(f"{key}={value}")
    if args.report:
        args.report.parent.mkdir(parents=True, exist_ok=True)
        args.report.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
