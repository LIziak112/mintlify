from __future__ import annotations

import argparse
import posixpath
import re
import struct
import zipfile
from pathlib import Path, PurePosixPath
from xml.etree import ElementTree as ET


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "o": "urn:schemas-microsoft-com:office:office",
    "v": "urn:schemas-microsoft-com:vml",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def relationships_part(owner_part: str) -> str:
    owner = PurePosixPath(owner_part)
    return str(owner.parent / "_rels" / f"{owner.name}.rels")


def resolve_target(owner_part: str, target: str) -> str:
    if target.startswith("/"):
        return posixpath.normpath(target.lstrip("/"))
    return posixpath.normpath(posixpath.join(posixpath.dirname(owner_part), target))


def parse_style(style: str) -> dict[str, str]:
    result: dict[str, str] = {}
    for declaration in style.split(";"):
        if ":" not in declaration:
            continue
        key, value = declaration.split(":", 1)
        result[key.strip().lower()] = value.strip()
    return result


def points(value: str) -> float:
    match = re.fullmatch(r"([-+]?[0-9]*\.?[0-9]+)\s*(pt|in|cm|mm|pc)?", value, re.IGNORECASE)
    if match is None:
        raise ValueError(f"unsupported VML dimension: {value!r}")
    number = float(match.group(1))
    unit = (match.group(2) or "pt").lower()
    factors = {"pt": 1.0, "in": 72.0, "cm": 72.0 / 2.54, "mm": 72.0 / 25.4, "pc": 12.0}
    return number * factors[unit]


def wmf_geometry(data: bytes) -> tuple[int, int, float, float]:
    if len(data) < 22:
        raise ValueError("WMF is shorter than its placeable header")
    key, _, left, top, right, bottom, units, _, _ = struct.unpack_from(
        "<I H h h h h H I H", data, 0
    )
    if key != 0x9AC6CDD7 or units <= 0:
        raise ValueError("WMF has no valid placeable header")
    width_pt = (right - left) * 72.0 / units
    height_pt = (bottom - top) * 72.0 / units
    return (
        round(width_pt * 20),
        round(height_pt * 20),
        round(width_pt / 0.75) * 0.75,
        round(height_pt / 0.75) * 0.75,
    )


def audit(path: Path) -> list[dict[str, object]]:
    rows: list[dict[str, object]] = []
    with zipfile.ZipFile(path) as archive:
        package_parts = set(archive.namelist())
        owner_parts = sorted(
            part
            for part in package_parts
            if part.startswith("word/")
            and part.endswith(".xml")
            and not part.startswith("word/_rels/")
        )
        for owner_part in owner_parts:
            root = ET.fromstring(archive.read(owner_part))
            rels_name = relationships_part(owner_part)
            relationships: dict[str, str] = {}
            if rels_name in package_parts:
                rels = ET.fromstring(archive.read(rels_name))
                for rel in rels.findall("pr:Relationship", NS):
                    rel_id = rel.attrib.get("Id")
                    target = rel.attrib.get("Target")
                    if rel_id and target and rel.attrib.get("TargetMode", "Internal") != "External":
                        relationships[rel_id] = resolve_target(owner_part, target)

            for obj in root.findall(".//w:object", NS):
                ole = obj.find(".//o:OLEObject", NS)
                if ole is None or ole.attrib.get("ProgID") != "Equation.DSMT4":
                    continue
                image = obj.find(".//v:imagedata", NS)
                shape = obj.find(".//v:shape", NS)
                if image is None or shape is None:
                    raise ValueError(f"{owner_part}: MathType object is missing VML preview data")
                rel_id = image.attrib.get(f"{{{NS['r']}}}id")
                preview_part = relationships.get(rel_id or "")
                if not preview_part or preview_part not in package_parts:
                    raise ValueError(f"{owner_part}: preview relationship {rel_id!r} is invalid")

                expected_dxa, expected_dya, expected_width, expected_height = wmf_geometry(
                    archive.read(preview_part)
                )
                style = parse_style(shape.attrib.get("style", ""))
                actual_width = points(style["width"]) if "width" in style else None
                actual_height = points(style["height"]) if "height" in style else None
                actual_dxa = int(obj.attrib[f"{{{NS['w']}}}dxaOrig"])
                actual_dya = int(obj.attrib[f"{{{NS['w']}}}dyaOrig"])
                rows.append(
                    {
                        "index": len(rows) + 1,
                        "preview_part": preview_part,
                        "dxa": actual_dxa,
                        "dya": actual_dya,
                        "width_pt": actual_width,
                        "height_pt": actual_height,
                        "ok": actual_dxa == expected_dxa
                        and actual_dya == expected_dya
                        and (actual_width is None or abs(actual_width - expected_width) < 1e-9)
                        and (actual_height is None or abs(actual_height - expected_height) < 1e-9),
                        "expected": (expected_dxa, expected_dya, expected_width, expected_height),
                    }
                )
    return rows


def main() -> int:
    parser = argparse.ArgumentParser(description="Audit MathType WMF and Word object geometry.")
    parser.add_argument("docx", type=Path)
    parser.add_argument("--show", type=int, nargs="*", default=[])
    args = parser.parse_args()
    rows = audit(args.docx.resolve())
    failures = [row for row in rows if not row["ok"]]
    missing_width = sum(row["width_pt"] is None for row in rows)
    missing_height = sum(row["height_pt"] is None for row in rows)
    print(f"document={args.docx.resolve()}")
    print(f"equations={len(rows)}")
    print(f"consistent={len(rows) - len(failures)}")
    print(f"failures={len(failures)}")
    print(f"missing_vml_width={missing_width}")
    print(f"missing_vml_height={missing_height}")
    for index in args.show:
        if 1 <= index <= len(rows):
            row = rows[index - 1]
            print(
                f"formula_{index}=dxa:{row['dxa']},dya:{row['dya']},"
                f"width:{row['width_pt']}pt,height:{row['height_pt']}pt,"
                f"preview:{row['preview_part']},ok:{row['ok']}"
            )
    for row in failures[:20]:
        print(f"failure={row}")
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())
