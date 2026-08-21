"""Enumerate Equation.DSMT4 objects across every internal Word XML part."""
from __future__ import annotations

import posixpath
import zipfile
from pathlib import PurePosixPath
from xml.etree import ElementTree as ET

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
O_NS = "urn:schemas-microsoft-com:office:office"
PR_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"w": W_NS, "o": O_NS, "r": R_NS, "pr": PR_NS}
REL_ID = f"{{{R_NS}}}id"
OLE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject"


def rels_part(owner_part: str) -> str:
    owner = PurePosixPath(owner_part)
    return str(owner.parent / "_rels" / f"{owner.name}.rels")


def resolve_target(owner_part: str, target: str) -> str:
    if target.startswith("/"):
        result = posixpath.normpath(target.lstrip("/"))
    else:
        result = posixpath.normpath(posixpath.join(posixpath.dirname(owner_part), target))
    if result == ".." or result.startswith("../"):
        raise ValueError(f"relationship escapes package: {owner_part} -> {target}")
    return result


def load_relationships(archive: zipfile.ZipFile, owner_part: str) -> dict[str, dict[str, str]]:
    name = rels_part(owner_part)
    if name not in archive.namelist():
        return {}
    root = ET.fromstring(archive.read(name))
    result: dict[str, dict[str, str]] = {}
    for rel in root.findall("pr:Relationship", NS):
        rid = rel.attrib.get("Id")
        target = rel.attrib.get("Target")
        if not rid or not target:
            continue
        external = rel.attrib.get("TargetMode", "Internal") == "External"
        result[rid] = {
            "type": rel.attrib.get("Type", ""),
            "external": str(external),
            "part": "" if external else resolve_target(owner_part, target),
        }
    return result


def enumerate_equations(archive: zipfile.ZipFile) -> list[dict[str, object]]:
    names = set(archive.namelist())
    owners = sorted(
        name
        for name in names
        if name.startswith("word/")
        and name.endswith(".xml")
        and not name.startswith("word/_rels/")
    )
    equations: list[dict[str, object]] = []
    for owner in owners:
        root = ET.fromstring(archive.read(owner))
        relationships = load_relationships(archive, owner)
        number = 0
        for obj in root.findall(".//w:object", NS):
            ole_nodes = [
                node
                for node in obj.findall(".//o:OLEObject", NS)
                if node.attrib.get("ProgID") == "Equation.DSMT4"
            ]
            if not ole_nodes:
                continue
            if len(ole_nodes) != 1:
                raise ValueError(f"{owner}: object contains {len(ole_nodes)} Equation.DSMT4 nodes")
            number += 1
            rid = ole_nodes[0].attrib.get(REL_ID)
            relationship = relationships.get(rid or "")
            if relationship is None:
                raise ValueError(f"{owner}: equation {number} has missing relationship {rid!r}")
            if relationship["external"] == "True":
                raise ValueError(f"{owner}: equation {number} uses an external OLE relationship")
            if relationship["type"] != OLE_REL_TYPE:
                raise ValueError(f"{owner}: equation {number} relationship is not oleObject")
            part = relationship["part"]
            if part not in names:
                raise ValueError(f"{owner}: equation {number} target is missing: {part}")
            equations.append(
                {
                    "index": len(equations) + 1,
                    "owner_part": owner,
                    "object_number_in_part": number,
                    "relationship_id": rid,
                    "ole_part": part,
                }
            )
    return equations
