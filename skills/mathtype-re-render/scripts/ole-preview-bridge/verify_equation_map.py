from __future__ import annotations

import argparse
import hashlib
import json
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from cfb_ole import CompoundFile


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "o": "urn:schemas-microsoft-com:office:office",
}


class VerificationError(RuntimeError):
    pass


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest().upper()


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest().upper()


def equation_native_stream(ole_data: bytes, label: str) -> bytes:
    try:
        compound = CompoundFile(ole_data)
        entry = compound.find_stream("Equation Native")
        if entry is None:
            raise VerificationError(f"{label} has no Equation Native stream")
        data = compound.read_entry(entry)
    except (ValueError, StopIteration, IndexError) as error:
        raise VerificationError(f"Cannot read {label} Equation Native stream: {error}") from error
    if not data:
        raise VerificationError(f"{label} has an empty Equation Native stream")
    return data


def count_equations(archive: zipfile.ZipFile) -> int:
    count = 0
    for part in archive.namelist():
        if not part.startswith("word/") or not part.endswith(".xml"):
            continue
        if part.startswith("word/_rels/"):
            continue
        root = ET.fromstring(archive.read(part))
        for obj in root.findall(".//w:object", NS):
            if any(
                ole.attrib.get("ProgID") == "Equation.DSMT4"
                for ole in obj.findall(".//o:OLEObject", NS)
            ):
                count += 1
    return count


def verify(source: Path, manifest_path: Path) -> dict[str, int]:
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    equations = manifest.get("equations")
    if not isinstance(equations, list):
        raise VerificationError("Manifest equations must be a list")
    if manifest.get("source_sha256") != sha256_file(source):
        raise VerificationError("Manifest source hash does not match the DOCX")
    if manifest.get("source_size") != source.stat().st_size:
        raise VerificationError("Manifest source size does not match the DOCX")

    ole_parts: set[str] = set()
    preview_parts: set[str] = set()
    ole_hashes: set[str] = set()
    native_hashes: set[str] = set()
    with zipfile.ZipFile(source) as archive:
        package_parts = set(archive.namelist())
        xml_equation_count = count_equations(archive)
        if xml_equation_count != len(equations):
            raise VerificationError(
                f"XML contains {xml_equation_count} equations, manifest contains {len(equations)}"
            )

        for expected_index, item in enumerate(equations, 1):
            if item.get("index") != expected_index:
                raise VerificationError(f"Manifest index is not sequential at {expected_index}")
            ole_part = item.get("ole_part")
            preview_part = item.get("preview_part")
            if not isinstance(ole_part, str) or ole_part not in package_parts:
                raise VerificationError(f"Equation {expected_index} has a missing OLE part")
            if not isinstance(preview_part, str) or preview_part not in package_parts:
                raise VerificationError(f"Equation {expected_index} has a missing preview part")
            if ole_part in ole_parts:
                raise VerificationError(f"OLE part is mapped more than once: {ole_part}")
            if preview_part in preview_parts:
                raise VerificationError(f"Preview part is mapped more than once: {preview_part}")

            ole_data = archive.read(ole_part)
            preview_data = archive.read(preview_part)
            ole_hash = sha256_bytes(ole_data)
            native_data = equation_native_stream(ole_data, f"Equation {expected_index}")
            native_hash = sha256_bytes(native_data)
            if item.get("ole_sha256") != ole_hash or item.get("ole_size") != len(ole_data):
                raise VerificationError(f"Equation {expected_index} OLE metadata is stale")
            if (
                item.get("equation_native_sha256") != native_hash
                or item.get("equation_native_size") != len(native_data)
            ):
                raise VerificationError(f"Equation {expected_index} Equation Native metadata is stale")
            if (
                item.get("preview_sha256") != sha256_bytes(preview_data)
                or item.get("preview_size") != len(preview_data)
            ):
                raise VerificationError(f"Equation {expected_index} preview metadata is stale")
            ole_parts.add(ole_part)
            preview_parts.add(preview_part)
            ole_hashes.add(ole_hash)
            native_hashes.add(native_hash)

    summary = manifest.get("summary", {})
    actual = {
        "equation_count": len(equations),
        "unique_ole_parts": len(ole_parts),
        "unique_ole_containers": len(ole_hashes),
        "unique_equation_native_payloads": len(native_hashes),
        "duplicate_equation_native_payloads": len(equations) - len(native_hashes),
        "unique_preview_parts": len(preview_parts),
    }
    for key, value in actual.items():
        if summary.get(key) != value:
            raise VerificationError(
                f"Manifest summary {key} is {summary.get(key)!r}, expected {value}"
            )
    return actual


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Verify an Equation.DSMT4 mapping manifest.")
    parser.add_argument("source", type=Path, help="Mapped DOCX")
    parser.add_argument("manifest", type=Path, help="JSON manifest")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    source = args.source.resolve()
    manifest = args.manifest.resolve()
    if not source.is_file():
        raise VerificationError(f"Input DOCX does not exist: {source}")
    if not manifest.is_file():
        raise VerificationError(f"Manifest does not exist: {manifest}")
    result = verify(source, manifest)
    print("status=ok")
    for key, value in result.items():
        print(f"{key}={value}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (VerificationError, zipfile.BadZipFile, ET.ParseError, OSError, ValueError) as error:
        print(f"status=error error={error}", file=sys.stderr)
        raise SystemExit(1)
