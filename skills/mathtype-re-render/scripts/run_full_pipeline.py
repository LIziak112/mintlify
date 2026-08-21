"""Run the complete EQP -> MTEF patch -> OLE/WMF refresh pipeline."""
from __future__ import annotations

import argparse
import hashlib
import json
import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parent
BRIDGE = ROOT / "ole-preview-bridge"


def run(arguments: list[str], allowed: tuple[int, ...] = (0,)) -> subprocess.CompletedProcess[str]:
    print("RUN=" + subprocess.list2cmdline(arguments), flush=True)
    result = subprocess.run(arguments, text=True, encoding="utf-8", errors="replace")
    if result.returncode not in allowed:
        raise RuntimeError(f"command failed with exit code {result.returncode}: {arguments[0]}")
    return result


def file_hash(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest().upper()


def main() -> int:
    parser = argparse.ArgumentParser(description="Patch all MathType MTEF settings and regenerate DOCX WMF previews.")
    parser.add_argument("eqp", type=Path)
    parser.add_argument("source", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--work-dir", type=Path)
    parser.add_argument("--cache", type=Path)
    parser.add_argument("--batch-size", type=int, default=25)
    parser.add_argument("--timeout", type=float, default=30.0)
    parser.add_argument("--overwrite", action="store_true")
    args = parser.parse_args()

    source = args.source.resolve()
    output = args.output.resolve()
    eqp = args.eqp.resolve()
    if not source.is_file() or source.suffix.lower() != ".docx":
        raise ValueError(f"source must be an existing DOCX: {source}")
    if not eqp.is_file():
        raise ValueError(f"EQP does not exist: {eqp}")
    if source == output:
        raise ValueError("output must differ from source")
    if output.exists() and not args.overwrite:
        raise ValueError(f"output exists; use --overwrite: {output}")
    if args.batch_size < 1 or args.timeout <= 0:
        raise ValueError("batch size and timeout must be positive")

    work = (args.work_dir or output.parent / f"{output.stem}.mathtype-work").resolve()
    cache = (args.cache or work / "cache").resolve()
    work.mkdir(parents=True, exist_ok=True)
    output.parent.mkdir(parents=True, exist_ok=True)
    python = sys.executable

    pre_audit = work / "01-source-mtef-audit.json"
    patched = work / "02-mtef-patched.docx"
    patch_report = work / "02-mtef-patch-report.json"
    patch_verify = work / "03-mtef-patch-verification.json"
    first_map = work / "04-patched-equation-map.json"
    unique_source = work / "05-unique-preview-source.docx"
    final_map = work / "05-equation-map.json"
    refresh_report = work / "06-preview-refresh-report.json"
    owned_outputs = [
        pre_audit,
        patched,
        patch_report,
        patch_verify,
        first_map,
        unique_source,
        final_map,
        refresh_report,
        work / "03-patched-mtef-audit.json",
        work / "pipeline-report.json",
    ]
    existing = [path for path in owned_outputs if path.exists()]
    if existing and not args.overwrite:
        raise ValueError(f"work directory contains existing pipeline outputs; use --overwrite: {existing[0]}")

    # A non-zero code of 2 means deviations were found, which is normal before patching.
    run([python, str(ROOT / "find_deviating.py"), str(eqp), str(source), "--out", str(pre_audit)], (0, 2))
    patch_args = [
        python,
        str(ROOT / "patch_mathtype_mtef.py"),
        str(eqp),
        str(source),
        str(patched),
        "--report",
        str(patch_report),
    ]
    if args.overwrite:
        patch_args.append("--overwrite")
    run(patch_args)
    run([
        python,
        str(ROOT / "verify_mtef_patch.py"),
        str(eqp),
        str(source),
        str(patched),
        "--report",
        str(patch_verify),
    ])
    run([python, str(ROOT / "find_deviating.py"), str(eqp), str(patched), "--out", str(work / "03-patched-mtef-audit.json")])
    run([python, str(BRIDGE / "map_equations.py"), str(patched), str(first_map)])
    manifest = json.loads(first_map.read_text(encoding="utf-8"))
    summary = manifest["summary"]
    if summary["unique_preview_parts"] < summary["equation_count"]:
        split_args = [
            python,
            str(BRIDGE / "split_shared_previews.py"),
            str(patched),
            str(first_map),
            str(unique_source),
        ]
        if args.overwrite:
            split_args.append("--overwrite")
        run(split_args)
        mapped_source = unique_source
        run([python, str(BRIDGE / "map_equations.py"), str(mapped_source), str(final_map)])
    else:
        mapped_source = patched
        final_map.write_text(first_map.read_text(encoding="utf-8"), encoding="utf-8")
    run([python, str(BRIDGE / "verify_equation_map.py"), str(mapped_source), str(final_map)])

    renderer = BRIDGE / "renderer32" / "OlePreviewRenderer.exe"
    if not renderer.is_file():
        run(["powershell", "-ExecutionPolicy", "Bypass", "-File", str(BRIDGE / "build_renderer32.ps1")])
    refresh_args = [
        python,
        str(BRIDGE / "refresh_previews.py"),
        str(mapped_source),
        str(output),
        "--manifest",
        str(final_map),
        "--renderer",
        str(renderer),
        "--cache",
        str(cache),
        "--report",
        str(refresh_report),
        "--batch-size",
        str(args.batch_size),
        "--timeout",
        str(args.timeout),
    ]
    if args.overwrite:
        refresh_args.append("--overwrite")
    run(refresh_args)
    run([python, str(BRIDGE / "audit_geometry.py"), str(output)])

    refresh = json.loads(refresh_report.read_text(encoding="utf-8"))
    result = {
        "status": "ok",
        "source": str(source),
        "eqp": str(eqp),
        "output": str(output),
        "output_sha256": file_hash(output),
        "equations": summary["equation_count"],
        "shared_previews_split": summary["equation_count"] - summary["unique_preview_parts"],
        "render": refresh["render"],
        "verification": refresh["verification"],
        "work_dir": str(work),
    }
    report = work / "pipeline-report.json"
    report.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    for key, value in result.items():
        print(f"{key}={value}")
    print(f"report={report}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (OSError, ValueError, RuntimeError, json.JSONDecodeError) as error:
        print(f"status=error error={error}", file=sys.stderr)
        raise SystemExit(1)
