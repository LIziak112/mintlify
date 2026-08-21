from __future__ import annotations

import argparse
import hashlib
import json
import posixpath
import sys
import zipfile
from pathlib import Path, PurePosixPath
from xml.etree import ElementTree as ET

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from cfb_ole import CompoundFile


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "o": "urn:schemas-microsoft-com:office:office",
    "v": "urn:schemas-microsoft-com:vml",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
}

REL_ID = f"{{{NS['r']}}}id"
OLE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject"
IMAGE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"


class MappingError(RuntimeError):
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
            raise MappingError(f"{label} has no Equation Native stream")
        data = compound.read_entry(entry)
    except (ValueError, StopIteration, IndexError) as error:
        raise MappingError(f"Cannot read {label} Equation Native stream: {error}") from error
    if not data:
        raise MappingError(f"{label} has an empty Equation Native stream")
    return data


def relationships_part(owner_part: str) -> str:
    owner = PurePosixPath(owner_part)
    return str(owner.parent / "_rels" / f"{owner.name}.rels")


def resolve_target(owner_part: str, target: str) -> str:
    if target.startswith("/"):
        resolved = posixpath.normpath(target.lstrip("/"))
    else:
        resolved = posixpath.normpath(posixpath.join(posixpath.dirname(owner_part), target))
    if resolved == ".." or resolved.startswith("../"):
        raise MappingError(f"Relationship target escapes the package: {owner_part} -> {target}")
    return resolved


def load_relationships(
    archive: zipfile.ZipFile, owner_part: str, package_parts: set[str]
) -> dict[str, dict[str, str]]:
    rels_part = relationships_part(owner_part)
    if rels_part not in package_parts:
        return {}
    root = ET.fromstring(archive.read(rels_part))
    relationships: dict[str, dict[str, str]] = {}
    for rel in root.findall("pr:Relationship", NS):
        rel_id = rel.attrib.get("Id")
        target = rel.attrib.get("Target")
        if not rel_id or not target:
            continue
        target_mode = rel.attrib.get("TargetMode", "Internal")
        relationships[rel_id] = {
            "type": rel.attrib.get("Type", ""),
            "target_mode": target_mode,
            "target": target,
            "part": resolve_target(owner_part, target) if target_mode != "External" else "",
        }
    return relationships


def require_relationship(
    relationships: dict[str, dict[str, str]],
    rel_id: str | None,
    expected_type: str,
    owner_part: str,
    object_number: int,
    role: str,
    package_parts: set[str],
) -> dict[str, str]:
    label = f"{owner_part} object {object_number} {role}"
    if not rel_id:
        raise MappingError(f"{label} has no relationship id")
    relationship = relationships.get(rel_id)
    if relationship is None:
        raise MappingError(f"{label} references missing relationship {rel_id}")
    if relationship["target_mode"] == "External":
        raise MappingError(f"{label} uses an external relationship")
    if relationship["type"] != expected_type:
        raise MappingError(
            f"{label} relationship {rel_id} has type {relationship['type']!r}, "
            f"expected {expected_type!r}"
        )
    if relationship["part"] not in package_parts:
        raise MappingError(f"{label} target is missing: {relationship['part']}")
    return relationship


def map_equations(source: Path) -> dict[str, object]:
    equations: list[dict[str, object]] = []
    with zipfile.ZipFile(source) as archive:
        package_parts = set(archive.namelist())
        candidate_parts = sorted(
            part
            for part in package_parts
            if part.startswith("word/")
            and part.endswith(".xml")
            and not part.startswith("word/_rels/")
        )

        for owner_part in candidate_parts:
            try:
                root = ET.fromstring(archive.read(owner_part))
            except ET.ParseError as error:
                raise MappingError(f"Cannot parse {owner_part}: {error}") from error

            math_objects: list[ET.Element] = []
            for obj in root.findall(".//w:object", NS):
                ole_nodes = [
                    node
                    for node in obj.findall(".//o:OLEObject", NS)
                    if node.attrib.get("ProgID") == "Equation.DSMT4"
                ]
                if ole_nodes:
                    if len(ole_nodes) != 1:
                        raise MappingError(
                            f"{owner_part} contains a w:object with {len(ole_nodes)} Equation.DSMT4 nodes"
                        )
                    math_objects.append(obj)

            if not math_objects:
                continue

            relationships = load_relationships(archive, owner_part, package_parts)
            for object_number, obj in enumerate(math_objects, 1):
                ole = obj.find(".//o:OLEObject", NS)
                image_nodes = obj.findall(".//v:imagedata", NS)
                if ole is None:
                    raise MappingError(f"{owner_part} object {object_number} has no OLEObject")
                if len(image_nodes) != 1:
                    raise MappingError(
                        f"{owner_part} object {object_number} has {len(image_nodes)} preview images"
                    )

                image = image_nodes[0]
                ole_rid = ole.attrib.get(REL_ID)
                image_rid = image.attrib.get(REL_ID)
                ole_rel = require_relationship(
                    relationships,
                    ole_rid,
                    OLE_REL_TYPE,
                    owner_part,
                    object_number,
                    "OLE",
                    package_parts,
                )
                image_rel = require_relationship(
                    relationships,
                    image_rid,
                    IMAGE_REL_TYPE,
                    owner_part,
                    object_number,
                    "preview",
                    package_parts,
                )

                ole_data = archive.read(ole_rel["part"])
                image_data = archive.read(image_rel["part"])
                native_data = equation_native_stream(
                    ole_data, f"{owner_part} object {object_number}"
                )
                shape = obj.find(".//v:shape", NS)
                equations.append(
                    {
                        "index": len(equations) + 1,
                        "owner_part": owner_part,
                        "object_number_in_part": object_number,
                        "prog_id": "Equation.DSMT4",
                        "ole_relationship_id": ole_rid,
                        "ole_part": ole_rel["part"],
                        "ole_size": len(ole_data),
                        "ole_sha256": sha256_bytes(ole_data),
                        "equation_native_size": len(native_data),
                        "equation_native_sha256": sha256_bytes(native_data),
                        "preview_relationship_id": image_rid,
                        "preview_part": image_rel["part"],
                        "preview_size": len(image_data),
                        "preview_sha256": sha256_bytes(image_data),
                        "dxa_orig": obj.attrib.get(f"{{{NS['w']}}}dxaOrig"),
                        "dya_orig": obj.attrib.get(f"{{{NS['w']}}}dyaOrig"),
                        "shape_id": shape.attrib.get("id") if shape is not None else None,
                        "shape_style": shape.attrib.get("style") if shape is not None else None,
                    }
                )

    unique_ole_hashes = {item["ole_sha256"] for item in equations}
    unique_native_hashes = {item["equation_native_sha256"] for item in equations}
    unique_ole_parts = {item["ole_part"] for item in equations}
    unique_preview_parts = {item["preview_part"] for item in equations}
    return {
        "schema_version": 2,
        "source": str(source.resolve()),
        "source_size": source.stat().st_size,
        "source_sha256": sha256_file(source),
        "summary": {
            "equation_count": len(equations),
            "unique_ole_parts": len(unique_ole_parts),
            "unique_ole_containers": len(unique_ole_hashes),
            "unique_equation_native_payloads": len(unique_native_hashes),
            "duplicate_equation_native_payloads": len(equations) - len(unique_native_hashes),
            "unique_preview_parts": len(unique_preview_parts),
        },
        "equations": equations,
    }


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Map every Equation.DSMT4 OLE object to its paired DOCX preview image."
    )
    parser.add_argument("source", type=Path, help="Input DOCX; opened read-only")
    parser.add_argument("output", type=Path, help="Output JSON manifest")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    source = args.source.resolve()
    output = args.output.resolve()
    if not source.is_file():
        raise MappingError(f"Input DOCX does not exist: {source}")
    if source.suffix.lower() != ".docx":
        raise MappingError(f"Input must be a DOCX: {source}")
    if source == output:
        raise MappingError("Manifest output cannot overwrite the input DOCX")

    manifest = map_equations(source)
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    summary = manifest["summary"]
    print(f"source={source}")
    print(f"manifest={output}")
    print(f"equation_count={summary['equation_count']}")
    print(f"unique_ole_parts={summary['unique_ole_parts']}")
    print(f"unique_ole_containers={summary['unique_ole_containers']}")
    print(f"unique_equation_native_payloads={summary['unique_equation_native_payloads']}")
    print(f"duplicate_equation_native_payloads={summary['duplicate_equation_native_payloads']}")
    print(f"unique_preview_parts={summary['unique_preview_parts']}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (MappingError, zipfile.BadZipFile, OSError) as error:
        print(f"error={error}", file=sys.stderr)
        raise SystemExit(1)
