from __future__ import annotations

import argparse
import json
import os
import posixpath
import re
import zipfile
from collections import defaultdict
from pathlib import Path, PurePosixPath


class SplitError(RuntimeError):
    pass


def rels_part(owner_part: str) -> str:
    owner = PurePosixPath(owner_part)
    return str(owner.parent / "_rels" / f"{owner.name}.rels")


def relative_target(owner_part: str, target_part: str) -> str:
    return posixpath.relpath(target_part, posixpath.dirname(owner_part))


def replace_preview_rid(block: bytes, old_rid: str, new_rid: str) -> bytes:
    pattern = re.compile(br"<v:imagedata\b[^>]*>")
    match = pattern.search(block)
    if match is None:
        raise SplitError("MathType object has no v:imagedata")
    opening = match.group(0)
    old = f'r:id="{old_rid}"'.encode("ascii")
    new = f'r:id="{new_rid}"'.encode("ascii")
    if opening.count(old) != 1:
        raise SplitError(f"Expected preview relationship {old_rid} exactly once")
    updated = opening.replace(old, new, 1)
    return block[: match.start()] + updated + block[match.end() :]


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Give every MathType object its own WMF preview relationship and part."
    )
    parser.add_argument("source", type=Path)
    parser.add_argument("manifest", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--overwrite", action="store_true")
    args = parser.parse_args()
    source = args.source.resolve()
    output = args.output.resolve()
    if source == output:
        raise SplitError("Output must differ from source")
    if output.exists() and not args.overwrite:
        raise SplitError(f"Output exists: {output}")

    manifest = json.loads(args.manifest.read_text(encoding="utf-8"))
    equations = manifest["equations"]
    groups: dict[str, list[dict[str, object]]] = defaultdict(list)
    for item in equations:
        groups[str(item["preview_part"])].append(item)
    to_split = [item for group in groups.values() if len(group) > 1 for item in group[1:]]

    with zipfile.ZipFile(source) as archive:
        names = archive.namelist()
        max_image = max(
            (int(m.group(1)) for name in names if (m := re.fullmatch(r"word/media/image(\d+)\.wmf", name))),
            default=0,
        )
        owner_data = {
            owner: archive.read(owner)
            for owner in {str(item["owner_part"]) for item in to_split}
        }
        rel_data = {rels_part(owner): archive.read(rels_part(owner)) for owner in owner_data}
        additions: dict[str, bytes] = {}
        assignments: list[tuple[dict[str, object], str, str]] = []

        next_image = max_image + 1
        next_rid: dict[str, int] = {}
        for owner in owner_data:
            rp = rels_part(owner)
            ids = [int(x) for x in re.findall(br'\bId="rId(\d+)"', rel_data[rp])]
            next_rid[owner] = max(ids, default=0) + 1

        for item in to_split:
            owner = str(item["owner_part"])
            new_part = f"word/media/image{next_image}.wmf"
            next_image += 1
            new_rid = f"rId{next_rid[owner]}"
            next_rid[owner] += 1
            additions[new_part] = archive.read(str(item["preview_part"]))
            assignments.append((item, new_part, new_rid))

        by_owner: dict[str, list[tuple[dict[str, object], str, str]]] = defaultdict(list)
        for assignment in assignments:
            by_owner[str(assignment[0]["owner_part"])].append(assignment)

        for owner, owner_assignments in by_owner.items():
            data = owner_data[owner]
            pattern = re.compile(br"<w:object\b[\s\S]*?</w:object>")
            objects = [m for m in pattern.finditer(data) if b'ProgID="Equation.DSMT4"' in m.group(0)]
            replacements: list[tuple[int, int, bytes]] = []
            rp = rels_part(owner)
            appended_relationships: list[bytes] = []
            for item, new_part, new_rid in owner_assignments:
                number = int(item["object_number_in_part"])
                match = objects[number - 1]
                block = replace_preview_rid(
                    match.group(0), str(item["preview_relationship_id"]), new_rid
                )
                replacements.append((match.start(), match.end(), block))
                target = relative_target(owner, new_part)
                appended_relationships.append(
                    (
                        '<Relationship Id="{}" '
                        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
                        'Target="{}"/>'
                    ).format(new_rid, target).encode("ascii")
                )
            for start, end, block in sorted(replacements, key=lambda value: value[0], reverse=True):
                data = data[:start] + block + data[end:]
            owner_data[owner] = data
            closing = rel_data[rp].rfind(b"</Relationships>")
            if closing < 0:
                raise SplitError(f"Invalid relationships part: {rp}")
            rel_data[rp] = (
                rel_data[rp][:closing]
                + b"".join(appended_relationships)
                + rel_data[rp][closing:]
            )

        partial = output.with_name(output.name + ".partial")
        if partial.exists():
            partial.unlink()
        try:
            with zipfile.ZipFile(partial, "w", zipfile.ZIP_DEFLATED, allowZip64=True) as out:
                for info in archive.infolist():
                    data = owner_data.get(info.filename, rel_data.get(info.filename))
                    if data is None:
                        data = archive.read(info.filename)
                    out.writestr(info, data)
                for name, data in additions.items():
                    out.writestr(name, data)
            with zipfile.ZipFile(partial) as check:
                bad = check.testzip()
                if bad:
                    raise SplitError(f"CRC failure: {bad}")
            os.replace(partial, output)
        finally:
            if partial.exists():
                partial.unlink()

    print(f"status=ok")
    print(f"split_objects={len(to_split)}")
    print(f"added_preview_parts={len(additions)}")
    print(f"output={output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
