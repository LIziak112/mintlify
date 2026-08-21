from __future__ import annotations

import argparse
import base64
import hashlib
import json
import os
import queue
import re
import shutil
import struct
import subprocess
import sys
import tempfile
import threading
import time
import zipfile
from pathlib import Path

from verify_equation_map import verify as verify_equation_map


class BridgeError(RuntimeError):
    pass


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest().upper()


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest().upper()


def encode_path(path: Path) -> str:
    return base64.b64encode(str(path.resolve()).encode("utf-8")).decode("ascii")


def decode_error(value: str) -> str:
    try:
        return base64.b64decode(value).decode("utf-8", "replace")
    except (ValueError, UnicodeError):
        return value


def validate_wmf(path: Path) -> dict[str, int | str]:
    data = path.read_bytes()
    if len(data) < 40:
        raise BridgeError(f"WMF is too short: {path}")
    key, handle, left, top, right, bottom, units, reserved, checksum = struct.unpack_from(
        "<I H h h h h H I H", data, 0
    )
    if key != 0x9AC6CDD7:
        raise BridgeError(f"WMF has no placeable header: {path}")
    words = struct.unpack_from("<10H", data, 0)
    calculated_checksum = 0
    for word in words:
        calculated_checksum ^= word
    if calculated_checksum != checksum:
        raise BridgeError(f"WMF placeable-header checksum is invalid: {path}")
    if units <= 0 or right <= left or bottom <= top:
        raise BridgeError(f"WMF has invalid bounds: {path}")

    file_type, header_words, version, file_words, objects, max_record, parameters = struct.unpack_from(
        "<H H H I H I H", data, 22
    )
    if file_type not in (1, 2) or header_words != 9 or version not in (0x0100, 0x0300):
        raise BridgeError(f"WMF standard header is invalid: {path}")
    if file_words * 2 > len(data) - 22 or objects < 0 or max_record < 3 or parameters != 0:
        raise BridgeError(f"WMF record metadata is invalid: {path}")
    return {
        "sha256": sha256_bytes(data),
        "size": len(data),
        "left": left,
        "top": top,
        "right": right,
        "bottom": bottom,
        "units_per_inch": units,
    }


def format_points(value: float) -> str:
    return f"{value:.4f}".rstrip("0").rstrip(".")


def wmf_geometry(path: Path) -> dict[str, int | float | str]:
    metadata = validate_wmf(path)
    width_pt = (int(metadata["right"]) - int(metadata["left"])) * 72.0 / int(
        metadata["units_per_inch"]
    )
    height_pt = (int(metadata["bottom"]) - int(metadata["top"])) * 72.0 / int(
        metadata["units_per_inch"]
    )
    # Word/MathType quantizes the VML display box to quarter-picas (0.75 pt).
    vml_width_pt = round(width_pt / 0.75) * 0.75
    vml_height_pt = round(height_pt / 0.75) * 0.75
    return {
        **metadata,
        "width_pt": width_pt,
        "height_pt": height_pt,
        "dxa_orig": round(width_pt * 20),
        "dya_orig": round(height_pt * 20),
        "vml_width_pt": vml_width_pt,
        "vml_height_pt": vml_height_pt,
        "vml_width": format_points(vml_width_pt) + "pt",
        "vml_height": format_points(vml_height_pt) + "pt",
    }


def replace_attribute(block: bytes, element: bytes, attribute: bytes, value: str) -> bytes:
    element_pattern = re.compile(br"<" + re.escape(element) + br"\b[^>]*>")
    match = element_pattern.search(block)
    if match is None:
        raise BridgeError(f"Cannot find {element.decode()} while updating equation geometry")
    opening = match.group(0)
    attribute_pattern = re.compile(
        br"(?<![-\w:])" + re.escape(attribute) + br'=\"[^\"]*\"'
    )
    replacement = attribute + b'=\"' + value.encode("ascii") + b'\"'
    if attribute_pattern.search(opening):
        updated_opening = attribute_pattern.sub(replacement, opening, count=1)
    else:
        insert_at = opening.find(b" ")
        if insert_at < 0:
            insert_at = opening.find(b">")
        updated_opening = opening[:insert_at] + b" " + replacement + opening[insert_at:]
    return block[: match.start()] + updated_opening + block[match.end() :]


def replace_style_dimension(style: bytes, key: bytes, value: str) -> bytes:
    pattern = re.compile(
        br"(?<![-\w])"
        + re.escape(key)
        + br"\s*:\s*[-+]?(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[a-z%]+)?",
        re.IGNORECASE,
    )
    replacement = key + b":" + value.encode("ascii")
    if pattern.search(style):
        return pattern.sub(replacement, style, count=1)
    separator = b"" if not style or style.endswith(b";") else b";"
    return style + separator + replacement


def patch_object_geometry(block: bytes, geometry: dict[str, int | float | str]) -> bytes:
    block = replace_attribute(block, b"w:object", b"w:dxaOrig", str(geometry["dxa_orig"]))
    block = replace_attribute(block, b"w:object", b"w:dyaOrig", str(geometry["dya_orig"]))
    shape_pattern = re.compile(br'<v:shape\b[^>]*\bstyle="([^"]*)"[^>]*>')
    shape = shape_pattern.search(block)
    if shape is None:
        raise BridgeError("MathType object has no VML shape style")
    style = shape.group(1)
    style = replace_style_dimension(style, b"width", str(geometry["vml_width"]))
    style = replace_style_dimension(style, b"height", str(geometry["vml_height"]))
    return block[: shape.start(1)] + style + block[shape.end(1) :]


def patch_owner_geometry(
    owner_data: bytes,
    owner_part: str,
    equations: list[dict[str, object]],
    geometries: dict[str, dict[str, int | float | str]],
) -> bytes:
    object_pattern = re.compile(br"<w:object\b[\s\S]*?</w:object>")
    matches = [
        match
        for match in object_pattern.finditer(owner_data)
        if b'ProgID="Equation.DSMT4"' in match.group(0)
    ]
    replacements: list[tuple[int, int, bytes]] = []
    for item in equations:
        object_number = int(item["object_number_in_part"])
        if object_number < 1 or object_number > len(matches):
            raise BridgeError(
                f"{owner_part} has no MathType object number {object_number}"
            )
        match = matches[object_number - 1]
        block = match.group(0)
        ole_rid = str(item["ole_relationship_id"]).encode("ascii")
        preview_rid = str(item["preview_relationship_id"]).encode("ascii")
        if ole_rid not in block or preview_rid not in block:
            raise BridgeError(
                f"{owner_part} object {object_number} relationships changed after mapping"
            )
        native_hash = str(item["equation_native_sha256"])
        replacements.append(
            (match.start(), match.end(), patch_object_geometry(block, geometries[native_hash]))
        )
    for start, end, block in reversed(replacements):
        owner_data = owner_data[:start] + block + owner_data[end:]
    return owner_data


class Worker:
    def __init__(self, executable: Path, log_path: Path, timeout_seconds: float):
        self.executable = executable
        self.log_path = log_path
        self.timeout_seconds = timeout_seconds
        self.process: subprocess.Popen[str] | None = None
        self.lines: queue.Queue[str | None] = queue.Queue()
        self.stderr_lines: list[str] = []
        self.stdout_thread: threading.Thread | None = None
        self.stderr_thread: threading.Thread | None = None

    @staticmethod
    def _read_lines(stream, target: queue.Queue[str | None] | list[str]) -> None:
        try:
            for line in iter(stream.readline, ""):
                value = line.rstrip("\r\n")
                if isinstance(target, queue.Queue):
                    target.put(value)
                else:
                    target.append(value)
        finally:
            if isinstance(target, queue.Queue):
                target.put(None)

    def start(self) -> None:
        if self.process is not None:
            raise BridgeError("Worker is already running")
        self.lines = queue.Queue()
        self.stderr_lines = []
        self.process = subprocess.Popen(
            [str(self.executable), "--worker"],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            errors="replace",
            bufsize=1,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        assert self.process.stdout is not None
        assert self.process.stderr is not None
        self.stdout_thread = threading.Thread(
            target=self._read_lines, args=(self.process.stdout, self.lines), daemon=True
        )
        self.stderr_thread = threading.Thread(
            target=self._read_lines, args=(self.process.stderr, self.stderr_lines), daemon=True
        )
        self.stdout_thread.start()
        self.stderr_thread.start()
        ready = self._next_line()
        if ready is None or not ready.lstrip("\ufeff").startswith("READY\t"):
            self.stop(force=True)
            raise BridgeError(f"Worker did not become ready: {ready!r}")

    def _next_line(self) -> str | None:
        try:
            return self.lines.get(timeout=self.timeout_seconds)
        except queue.Empty as error:
            self.stop(force=True)
            raise BridgeError(f"Worker timed out after {self.timeout_seconds:g} seconds") from error

    def render(self, ole_path: Path, wmf_path: Path, dimensions_path: Path) -> dict[str, int]:
        if self.process is None or self.process.stdin is None:
            raise BridgeError("Worker is not running")
        request = "\t".join(
            ["RENDER", encode_path(ole_path), encode_path(wmf_path), encode_path(dimensions_path)]
        )
        self.process.stdin.write(request + "\n")
        self.process.stdin.flush()
        response = self._next_line()
        if response is None:
            raise BridgeError("Worker exited before returning a result")
        fields = response.lstrip("\ufeff").split("\t")
        if fields[0] == "ERR":
            raise BridgeError(decode_error(fields[1] if len(fields) > 1 else "Worker error"))
        if len(fields) != 4 or fields[0] != "OK":
            raise BridgeError(f"Invalid worker response: {response!r}")
        return {
            "mapping_mode": int(fields[1]),
            "x_ext": int(fields[2]),
            "y_ext": int(fields[3]),
        }

    def stop(self, force: bool = False) -> None:
        process = self.process
        if process is None:
            return
        try:
            if not force and process.poll() is None and process.stdin is not None:
                process.stdin.write("QUIT\n")
                process.stdin.flush()
                try:
                    process.wait(timeout=5)
                except subprocess.TimeoutExpired:
                    force = True
            if force and process.poll() is None:
                process.kill()
                process.wait(timeout=5)
        finally:
            self.log_path.parent.mkdir(parents=True, exist_ok=True)
            if self.stderr_lines:
                self.log_path.write_text("\n".join(self.stderr_lines) + "\n", encoding="utf-8")
            self.process = None


def load_manifest(source: Path, manifest_path: Path) -> dict[str, object]:
    verify_equation_map(source, manifest_path)
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    if manifest.get("schema_version") != 2:
        raise BridgeError("Mapping manifest schema version 2 is required")
    equations = manifest.get("equations")
    if not isinstance(equations, list) or not equations:
        raise BridgeError("Mapping manifest contains no equations")
    for item in equations:
        preview = item.get("preview_part")
        if not isinstance(preview, str) or not preview.lower().endswith(".wmf"):
            raise BridgeError(f"Only WMF preview parts are supported: {preview!r}")
    return manifest


def extract_unique_oles(
    source: Path, equations: list[dict[str, object]], extraction_dir: Path
) -> dict[str, Path]:
    extraction_dir.mkdir(parents=True, exist_ok=True)
    unique: dict[str, dict[str, object]] = {}
    for item in equations:
        native_hash = str(item["equation_native_sha256"])
        unique.setdefault(native_hash, item)

    result: dict[str, Path] = {}
    with zipfile.ZipFile(source) as archive:
        for native_hash, item in unique.items():
            data = archive.read(str(item["ole_part"]))
            if sha256_bytes(data) != item["ole_sha256"]:
                raise BridgeError(f"OLE changed after mapping: {item['ole_part']}")
            path = extraction_dir / f"{native_hash}.bin"
            path.write_bytes(data)
            result[native_hash] = path
    return result


def render_unique_previews(
    executable: Path,
    renderer_hash: str,
    ole_paths: dict[str, Path],
    cache_dir: Path,
    work_dir: Path,
    batch_size: int,
    timeout_seconds: float,
) -> tuple[dict[str, Path], dict[str, object]]:
    cache_namespace = cache_dir / f"renderer-{renderer_hash}"
    cache_namespace.mkdir(parents=True, exist_ok=True)
    results: dict[str, Path] = {}
    rendered = 0
    cache_hits = 0
    worker_starts = 0
    worker: Worker | None = None
    worker_items = 0
    started = time.perf_counter()

    try:
        for index, (native_hash, ole_path) in enumerate(ole_paths.items(), 1):
            cached_wmf = cache_namespace / f"{native_hash}.wmf"
            cached_dimensions = cache_namespace / f"{native_hash}.json"
            if cached_wmf.is_file() and cached_dimensions.is_file():
                validate_wmf(cached_wmf)
                results[native_hash] = cached_wmf
                cache_hits += 1
                print(f"[{index}/{len(ole_paths)}] cache {native_hash[:12]}", flush=True)
                continue

            if worker is None or worker_items >= batch_size:
                if worker is not None:
                    worker.stop()
                worker_starts += 1
                worker = Worker(
                    executable,
                    work_dir / "logs" / f"worker-{worker_starts:03d}.stderr.log",
                    timeout_seconds,
                )
                worker.start()
                worker_items = 0

            temporary_wmf = work_dir / "rendered" / f"{native_hash}.wmf"
            temporary_dimensions = work_dir / "rendered" / f"{native_hash}.json"
            temporary_wmf.parent.mkdir(parents=True, exist_ok=True)
            worker.render(ole_path, temporary_wmf, temporary_dimensions)
            validate_wmf(temporary_wmf)
            if not temporary_dimensions.is_file():
                raise BridgeError(f"Worker did not write dimensions: {temporary_dimensions}")
            os.replace(temporary_wmf, cached_wmf)
            os.replace(temporary_dimensions, cached_dimensions)
            results[native_hash] = cached_wmf
            rendered += 1
            worker_items += 1
            print(f"[{index}/{len(ole_paths)}] render {native_hash[:12]}", flush=True)
    finally:
        if worker is not None:
            worker.stop()

    return results, {
        "unique_equations": len(ole_paths),
        "rendered": rendered,
        "cache_hits": cache_hits,
        "worker_starts": worker_starts,
        "elapsed_seconds": round(time.perf_counter() - started, 3),
    }


def patch_docx(
    source: Path,
    output: Path,
    equations: list[dict[str, object]],
    rendered: dict[str, Path],
) -> tuple[dict[str, str], dict[str, object]]:
    replacements: dict[str, bytes] = {}
    geometries: dict[str, dict[str, int | float | str]] = {}
    for item in equations:
        preview_part = str(item["preview_part"])
        native_hash = str(item["equation_native_sha256"])
        data = rendered[native_hash].read_bytes()
        geometries.setdefault(native_hash, wmf_geometry(rendered[native_hash]))
        existing = replacements.get(preview_part)
        if existing is not None and existing != data:
            raise BridgeError(f"Conflicting replacement for {preview_part}")
        replacements[preview_part] = data

    equations_by_owner: dict[str, list[dict[str, object]]] = {}
    for item in equations:
        equations_by_owner.setdefault(str(item["owner_part"]), []).append(item)

    temporary = output.with_name(output.name + ".partial")
    if temporary.exists():
        temporary.unlink()
    try:
        with zipfile.ZipFile(source) as original, zipfile.ZipFile(
            temporary, "w", compression=zipfile.ZIP_DEFLATED, allowZip64=True
        ) as patched:
            for info in original.infolist():
                data = replacements.get(info.filename)
                if data is None:
                    data = original.read(info.filename)
                if info.filename in equations_by_owner:
                    data = patch_owner_geometry(
                        data,
                        info.filename,
                        equations_by_owner[info.filename],
                        geometries,
                    )
                patched.writestr(info, data)
        with zipfile.ZipFile(temporary) as check:
            bad = check.testzip()
            if bad is not None:
                raise BridgeError(f"Patched DOCX CRC failed at {bad}")
        replacement_hashes = {part: sha256_bytes(data) for part, data in replacements.items()}
        verification = verify_output(
            source, temporary, equations, replacement_hashes, geometries
        )
        os.replace(temporary, output)
    finally:
        if temporary.exists():
            temporary.unlink()
    return replacement_hashes, verification


def verify_output(
    source: Path,
    output: Path,
    equations: list[dict[str, object]],
    replacement_hashes: dict[str, str],
    geometries: dict[str, dict[str, int | float | str]],
) -> dict[str, object]:
    preview_parts = set(replacement_hashes)
    owner_parts = {str(item["owner_part"]) for item in equations}
    changed_parts: list[str] = []
    unchanged_parts = 0
    with zipfile.ZipFile(source) as before, zipfile.ZipFile(output) as after:
        before_names = before.namelist()
        after_names = after.namelist()
        if before_names != after_names:
            raise BridgeError("DOCX package entry order or membership changed")
        for part in before_names:
            before_data = before.read(part)
            after_data = after.read(part)
            if part in preview_parts:
                if sha256_bytes(after_data) != replacement_hashes[part]:
                    raise BridgeError(f"Replacement hash mismatch: {part}")
                if before_data != after_data:
                    changed_parts.append(part)
            elif part in owner_parts:
                expected = patch_owner_geometry(
                    before_data,
                    part,
                    [item for item in equations if str(item["owner_part"]) == part],
                    geometries,
                )
                if after_data != expected:
                    raise BridgeError(f"Unexpected XML change outside equation geometry: {part}")
                if before_data != after_data:
                    changed_parts.append(part)
            else:
                if before_data != after_data:
                    raise BridgeError(f"Unexpected package change: {part}")
                unchanged_parts += 1

    expected_ole_hashes = {str(item["ole_part"]): str(item["ole_sha256"]) for item in equations}
    with zipfile.ZipFile(output) as archive:
        for part, expected_hash in expected_ole_hashes.items():
            if sha256_bytes(archive.read(part)) != expected_hash:
                raise BridgeError(f"OLE part changed: {part}")

    changed_preview_parts = preview_parts.intersection(changed_parts)
    if len(changed_preview_parts) != len(preview_parts):
        unchanged_previews = sorted(preview_parts.difference(changed_preview_parts))
        raise BridgeError(
            f"Only {len(changed_preview_parts)} of {len(preview_parts)} previews changed; "
            f"unchanged examples: {unchanged_previews[:3]}"
        )
    return {
        "package_entries": unchanged_parts + len(preview_parts) + len(owner_parts),
        "replaced_preview_parts": len(preview_parts),
        "updated_geometry_objects": len(equations),
        "updated_geometry_parts": sorted(owner_parts),
        "changed_parts": changed_parts,
        "unchanged_non_preview_parts": unchanged_parts,
        "ole_parts_verified_unchanged": len(expected_ole_hashes),
        "baseline_position_preserved": True,
    }


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Render Equation.DSMT4 previews through an x86 OLE worker and replace mapped WMFs."
    )
    parser.add_argument("source", type=Path, help="Mapped source DOCX; never modified")
    parser.add_argument("output", type=Path, help="New DOCX with refreshed WMF previews")
    parser.add_argument("--manifest", type=Path, required=True, help="Schema-v2 equation map")
    parser.add_argument(
        "--renderer",
        type=Path,
        default=Path(__file__).parent / "renderer32" / "OlePreviewRenderer.exe",
    )
    parser.add_argument("--cache", type=Path, default=Path(__file__).parent / "cache")
    parser.add_argument("--report", type=Path, default=None)
    parser.add_argument("--batch-size", type=int, default=25)
    parser.add_argument("--timeout", type=float, default=30.0)
    parser.add_argument("--overwrite", action="store_true")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    source = args.source.resolve()
    output = args.output.resolve()
    manifest_path = args.manifest.resolve()
    renderer = args.renderer.resolve()
    cache_dir = args.cache.resolve()
    report_path = (
        args.report.resolve()
        if args.report is not None
        else output.with_suffix(output.suffix + ".preview-report.json")
    )
    if not source.is_file():
        raise BridgeError(f"Source DOCX does not exist: {source}")
    if not renderer.is_file():
        raise BridgeError(f"x86 renderer does not exist: {renderer}")
    if source == output:
        raise BridgeError("Output must differ from source; in-place updates are not allowed")
    if output.exists() and not args.overwrite:
        raise BridgeError(f"Output already exists; use --overwrite: {output}")
    if args.batch_size < 1:
        raise BridgeError("--batch-size must be at least 1")
    if args.timeout <= 0:
        raise BridgeError("--timeout must be positive")

    manifest = load_manifest(source, manifest_path)
    equations = manifest["equations"]
    assert isinstance(equations, list)
    renderer_hash = sha256_file(renderer)
    output.parent.mkdir(parents=True, exist_ok=True)
    cache_dir.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory(prefix="ole-preview-bridge-", dir=str(output.parent)) as temp_name:
        work_dir = Path(temp_name)
        ole_paths = extract_unique_oles(source, equations, work_dir / "ole")
        rendered, render_stats = render_unique_previews(
            renderer,
            renderer_hash,
            ole_paths,
            cache_dir,
            work_dir,
            args.batch_size,
            args.timeout,
        )
        replacement_hashes, verification = patch_docx(source, output, equations, rendered)
    report = {
        "schema_version": 1,
        "source": str(source),
        "source_sha256": sha256_file(source),
        "output": str(output),
        "output_sha256": sha256_file(output),
        "manifest": str(manifest_path),
        "renderer": str(renderer),
        "renderer_sha256": renderer_hash,
        "batch_size": args.batch_size,
        "render": render_stats,
        "verification": verification,
    }
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"status=ok")
    print(f"output={output}")
    print(f"output_sha256={report['output_sha256']}")
    print(f"equations={len(equations)}")
    print(f"unique_equations={render_stats['unique_equations']}")
    print(f"rendered={render_stats['rendered']}")
    print(f"cache_hits={render_stats['cache_hits']}")
    print(f"worker_starts={render_stats['worker_starts']}")
    print(f"elapsed_seconds={render_stats['elapsed_seconds']}")
    print(f"report={report_path}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (BridgeError, OSError, ValueError, zipfile.BadZipFile) as error:
        print(f"status=error error={error}", file=sys.stderr)
        raise SystemExit(1)
