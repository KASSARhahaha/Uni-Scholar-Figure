#!/usr/bin/env python3
"""Validate evidence, FigEdit gates, and native objects in a Figure 1 bundle."""

from __future__ import annotations

import argparse
import csv
import json
import re
import sys
import zipfile
from pathlib import Path
from typing import Any


FULL_TEXT_FILES = {
    "research/research_brief.md": [
        "# Research brief",
        "## Research question",
        "## Scientific gap",
        "## Supported claims",
        "## Conflicts and missing evidence",
        "## Defensible Figure 1 conclusion",
    ],
    "design/figure_contract.md": [
        "# Figure contract",
        "Core conclusion:",
        "Figure role:",
        "Figure archetype:",
        "Target journal and article type:",
        "## Panel map",
        "## Evidence hierarchy",
        "Main reviewer risk:",
    ],
    "design/style_reference.md": [
        "# Style reference",
        "Reference citation:",
        "Reference URL or DOI:",
        "Target journal guide:",
        "## Abstract design tokens",
        "## Originality statement",
    ],
    "delivery/release_notes.md": [
        "# Figure 1 release notes",
        "Core conclusion:",
        "Target journal:",
        "Research cutoff date:",
        "Reference figure:",
        "Editability summary:",
        "Native PowerPoint rendering:",
        "Unresolved risks:",
    ],
}

RELEASE_STATES = {
    "PPTX_STRUCTURAL_PASS",
    "LIBREOFFICE_RENDER_PASS_ADVISORY",
    "POWERPOINT_NATIVE_PASS",
}

PLACEHOLDER_RE = re.compile(
    r"(?:\bUNSET\b|\bPLACEHOLDER\b|\[TODO[^\]]*\]|:\s*(?:SET|TBD|TODO|N/?A)\s*$)",
    re.IGNORECASE | re.MULTILINE,
)


def inspect_pptx(path: Path) -> dict[str, int | bool]:
    counts: dict[str, int | bool] = {
        "slides": 0,
        "text_runs": 0,
        "shapes": 0,
        "connectors": 0,
        "pictures": 0,
    }
    with zipfile.ZipFile(path) as archive:
        slide_names = sorted(
            name
            for name in archive.namelist()
            if re.fullmatch(r"ppt/slides/slide\d+\.xml", name)
        )
        counts["slides"] = len(slide_names)
        for name in slide_names:
            xml = archive.read(name).decode("utf-8", errors="replace")
            counts["text_runs"] = int(counts["text_runs"]) + len(re.findall(r"<a:t(?:\s|>)", xml))
            counts["shapes"] = int(counts["shapes"]) + len(re.findall(r"<p:sp(?:\s|>)", xml))
            counts["connectors"] = int(counts["connectors"]) + len(re.findall(r"<p:cxnSp(?:\s|>)", xml))
            counts["pictures"] = int(counts["pictures"]) + len(re.findall(r"<p:pic(?:\s|>)", xml))
    counts["flattened_risk"] = bool(
        int(counts["slides"]) > 0
        and int(counts["pictures"]) >= int(counts["slides"])
        and int(counts["text_runs"]) == 0
        and int(counts["shapes"]) == 0
        and int(counts["connectors"]) == 0
    )
    return counts


def parse_status(text: str) -> str | None:
    match = re.search(r"^-?\s*Status:\s*`?([^`\n]+)`?\s*$", text, re.IGNORECASE | re.MULTILINE)
    return match.group(1).strip() if match else None


def check_structured_text(root: Path, mode: str, errors: list[str], warnings: list[str]) -> dict[str, str]:
    required = FULL_TEXT_FILES if mode == "full" else {
        "delivery/release_notes.md": FULL_TEXT_FILES["delivery/release_notes.md"]
    }
    contents: dict[str, str] = {}
    for relative, markers in required.items():
        path = root / relative
        if not path.is_file():
            errors.append(f"missing required file: {relative}")
            continue
        text = path.read_text(encoding="utf-8", errors="replace")
        contents[relative] = text
        if len(text.strip()) < 120:
            errors.append(f"required file is too short to be substantive: {relative}")
        missing_markers = [marker for marker in markers if marker not in text]
        if missing_markers:
            errors.append(f"required structure missing in {relative}: {', '.join(missing_markers)}")
        if PLACEHOLDER_RE.search(text):
            errors.append(f"unresolved placeholder fields: {relative}")
        status = parse_status(text)
        if status is None:
            errors.append(f"missing Status field: {relative}")
        elif status.lower() in {"draft", "pending", "incomplete"}:
            warnings.append(f"non-release status in {relative}: {status}")
    return contents


def check_evidence_ledger(root: Path, errors: list[str]) -> dict[str, dict[str, str]]:
    path = root / "research" / "evidence_ledger.csv"
    expected = {
        "claim_id",
        "claim_text",
        "evidence_type",
        "source_title",
        "doi_or_url",
        "publication_year",
        "accessed_date",
        "support_level",
        "figure_element",
        "notes",
    }
    if not path.is_file():
        errors.append("missing required file: research/evidence_ledger.csv")
        return {}
    try:
        with path.open(newline="", encoding="utf-8-sig") as handle:
            reader = csv.DictReader(handle)
            fields = set(reader.fieldnames or [])
            if fields != expected:
                errors.append("evidence ledger columns do not match the required schema")
                return {}
            rows = list(reader)
    except (OSError, csv.Error) as exc:
        errors.append(f"cannot parse evidence ledger: {exc}")
        return {}
    if not rows:
        errors.append("evidence ledger has no claim rows")
        return {}

    supported = 0
    claims: dict[str, dict[str, str]] = {}
    allowed_evidence_types = {
        "author-supplied",
        "primary-literature",
        "official-guidance",
        "hypothesis",
        "visual-reference",
    }
    allowed_support_levels = {
        "supports",
        "partly-supports",
        "contradicts",
        "context-only",
        "pending",
    }
    source_required = {"primary-literature", "official-guidance", "visual-reference"}
    for number, row in enumerate(rows, start=2):
        for field in ("claim_id", "claim_text", "evidence_type", "support_level", "figure_element"):
            if not (row.get(field) or "").strip():
                errors.append(f"evidence ledger row {number} has an empty {field}")
        evidence_type = (row.get("evidence_type") or "").strip()
        support_level = (row.get("support_level") or "").strip()
        claim_id = (row.get("claim_id") or "").strip()
        if evidence_type not in allowed_evidence_types:
            errors.append(f"evidence ledger row {number} has invalid evidence_type: {evidence_type or 'empty'}")
        if support_level not in allowed_support_levels:
            errors.append(f"evidence ledger row {number} has invalid support_level: {support_level or 'empty'}")
        if evidence_type in source_required:
            for field in ("source_title", "doi_or_url", "accessed_date"):
                if not (row.get(field) or "").strip():
                    errors.append(f"evidence ledger row {number} requires {field}")
        if evidence_type == "primary-literature" and not re.fullmatch(r"\d{4}", (row.get("publication_year") or "").strip()):
            errors.append(f"evidence ledger row {number} requires a four-digit publication_year")
        if claim_id in claims:
            errors.append(f"duplicate evidence claim_id: {claim_id}")
        elif claim_id:
            claims[claim_id] = row
        if row.get("support_level") in {"supports", "partly-supports"}:
            supported += 1
    if supported == 0:
        errors.append("evidence ledger has no supporting or partly supporting claim")
    return claims


def cross_check_evidence(
    manifest: dict[str, Any] | None,
    claims: dict[str, dict[str, str]],
    errors: list[str],
) -> None:
    if not manifest or not claims:
        return
    elements = [item for item in manifest.get("elements", []) if isinstance(item, dict)]
    panels = [item for item in manifest.get("panels", []) if isinstance(item, dict)]
    manifest_ids = {str(item.get("id")) for item in [*elements, *panels] if item.get("id")}
    for claim_id, row in claims.items():
        targets = [item.strip() for item in re.split(r"[;,]", row.get("figure_element") or "") if item.strip()]
        missing = [target for target in targets if target not in manifest_ids]
        if missing:
            errors.append(f"evidence claim {claim_id} maps to unknown figure element(s): {', '.join(missing)}")

    support_by_claim = {
        claim_id: row.get("support_level") in {"supports", "partly-supports"}
        for claim_id, row in claims.items()
    }
    semantic_types = {"text", "line", "path", "polyline"}
    for element in elements:
        if element.get("type") not in semantic_types:
            continue
        element_id = str(element.get("id") or "<missing-id>")
        evidence_ids = element.get("evidence_ids")
        exemption = str(element.get("evidence_exempt_reason") or "").strip()
        if not evidence_ids and not exemption:
            errors.append(f"semantic figure element lacks evidence_ids or an exemption: {element_id}")
            continue
        if evidence_ids:
            if not isinstance(evidence_ids, list) or not all(isinstance(item, str) and item for item in evidence_ids):
                errors.append(f"semantic figure element has invalid evidence_ids: {element_id}")
                continue
            unknown = [claim_id for claim_id in evidence_ids if claim_id not in claims]
            if unknown:
                errors.append(f"semantic figure element cites unknown evidence: {element_id} -> {', '.join(unknown)}")
                continue
            if not any(support_by_claim.get(claim_id, False) for claim_id in evidence_ids):
                evidence_state = element.get("evidence_state")
                if evidence_state not in {"hypothesis", "planned", "context"}:
                    errors.append(
                        f"semantic figure element lacks supporting evidence and is not marked hypothesis/planned/context: {element_id}"
                    )


def load_manifest(path: Path, errors: list[str]) -> dict[str, Any] | None:
    try:
        manifest = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        errors.append(f"cannot parse FigEdit manifest: {exc}")
        return None
    if not isinstance(manifest, dict):
        errors.append("FigEdit manifest root must be an object")
        return None
    plan = manifest.get("reconstruction_plan")
    if not isinstance(plan, dict):
        errors.append("FigEdit manifest lacks reconstruction_plan")
    elif plan.get("open_questions"):
        errors.append("FigEdit manifest still has open_questions")

    elements = manifest.get("elements")
    if not isinstance(elements, list) or not elements:
        errors.append("FigEdit manifest has no editable elements")
        return manifest
    for element in elements:
        if not isinstance(element, dict):
            errors.append("FigEdit manifest contains a non-object element")
            continue
        element_id = str(element.get("id") or "<missing-id>")
        if element.get("review_status") != "verified":
            errors.append(f"FigEdit element is not verified: {element_id}")
        if not element.get("decision"):
            errors.append(f"FigEdit element lacks a reconstruction decision: {element_id}")
    return manifest


def parse_quality_gates(text: str) -> dict[str, str]:
    return {
        name: status.lower()
        for name, status in re.findall(r"^- ([a-z0-9_]+): `([^`]+)`", text, re.MULTILINE)
    }


def metric(text: str, label: str) -> int | None:
    match = re.search(rf"^- {re.escape(label)}:\s*(\d+)\s*$", text, re.MULTILINE)
    return int(match.group(1)) if match else None


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("directory", type=Path)
    parser.add_argument("--mode", choices=["full", "convert-only"], default="full")
    parser.add_argument("--strict", action="store_true")
    parser.add_argument("--json", action="store_true", dest="as_json")
    args = parser.parse_args()

    root = args.directory.expanduser().resolve()
    errors: list[str] = []
    warnings: list[str] = []
    text_contents = check_structured_text(root, args.mode, errors, warnings)
    claims = check_evidence_ledger(root, errors) if args.mode == "full" else {}

    release_text = text_contents.get("delivery/release_notes.md", "")
    release_match = re.search(r"^Native PowerPoint rendering:\s*([^\n]+)$", release_text, re.MULTILINE)
    release_state = release_match.group(1).strip() if release_match else None
    if release_state not in RELEASE_STATES:
        warnings.append("release notes do not declare an accepted PPTX verification state")

    source_candidates = list((root / "source").glob("*.svg")) if (root / "source").exists() else []
    preview_candidates = list((root / "source").glob("*.png")) if (root / "source").exists() else []
    if not source_candidates:
        errors.append("missing source SVG in source/")
    if not preview_candidates:
        errors.append("missing source preview PNG in source/")

    figedit_out = root / "figedit" / "out"
    required_outputs = [
        "editable.svg",
        "editable.pptx",
        "manifest.json",
        "quality_report.md",
        "editability_report.md",
    ]
    for filename in required_outputs:
        if not (figedit_out / filename).is_file():
            errors.append(f"missing FigEdit output: figedit/out/{filename}")

    manifest_path = figedit_out / "manifest.json"
    manifest = load_manifest(manifest_path, errors) if manifest_path.is_file() else None
    if args.mode == "full":
        cross_check_evidence(manifest, claims, errors)
    elements = manifest.get("elements", []) if manifest else []
    assets = manifest.get("assets", []) if manifest else []
    text_count = sum(isinstance(item, dict) and item.get("type") == "text" for item in elements)
    structural_count = sum(
        isinstance(item, dict)
        and item.get("type") in {"rect", "line", "path", "circle", "ellipse", "polygon", "polyline"}
        for item in elements
    )
    if manifest and text_count == 0:
        errors.append("FigEdit manifest has no editable text elements")
    if manifest and structural_count == 0:
        errors.append("FigEdit manifest has no editable structural elements")

    gates: dict[str, str] = {}
    quality_path = figedit_out / "quality_report.md"
    if quality_path.is_file():
        quality_text = quality_path.read_text(encoding="utf-8", errors="replace")
        gates = parse_quality_gates(quality_text)
        for name in ("xml_editable", "xml_embedded", "pptx_export", "pptx_text_fit"):
            if gates.get(name) != "ok":
                errors.append(f"required FigEdit quality gate is not ok: {name}={gates.get(name, 'missing')}")
        for name, gate_status in gates.items():
            if gate_status in {"failed", "fail", "error", "review"}:
                errors.append(f"FigEdit quality gate requires attention: {name}={gate_status}")

    editability_status = None
    image_count = None
    editability_path = figedit_out / "editability_report.md"
    if editability_path.is_file():
        editability_text = editability_path.read_text(encoding="utf-8", errors="replace")
        editability_status = parse_status(editability_text)
        image_count = metric(editability_text, "Image elements")
        report_text_count = metric(editability_text, "SVG text elements")
        report_structural_count = metric(editability_text, "Structural SVG elements")
        if report_text_count is None or report_text_count < 1:
            errors.append("editability report does not confirm editable SVG text")
        if report_structural_count is None or report_structural_count < 1:
            errors.append("editability report does not confirm structural SVG elements")
        if editability_status not in {"ok", "unavailable"}:
            errors.append(f"editability report status requires attention: {editability_status or 'missing'}")

    pptx_info = None
    pptx_path = figedit_out / "editable.pptx"
    if pptx_path.is_file():
        try:
            pptx_info = inspect_pptx(pptx_path)
            if pptx_info["flattened_risk"]:
                errors.append("editable.pptx appears to contain only flattened slide images")
            if int(pptx_info["text_runs"]) < max(1, text_count):
                errors.append("editable.pptx has fewer text runs than the verified manifest")
            if int(pptx_info["shapes"]) + int(pptx_info["connectors"]) < max(1, structural_count):
                errors.append("editable.pptx has fewer native shapes than the verified manifest")
        except (OSError, zipfile.BadZipFile) as exc:
            errors.append(f"cannot inspect editable.pptx: {exc}")

    if editability_status == "unavailable":
        deterministic_exception = bool(
            manifest
            and not assets
            and image_count == 0
            and release_state in RELEASE_STATES
            and pptx_info
            and int(pptx_info["text_runs"]) >= max(1, text_count)
            and int(pptx_info["shapes"]) + int(pptx_info["connectors"]) >= max(1, structural_count)
        )
        if not deterministic_exception:
            errors.append("editability is unavailable without a qualifying deterministic, image-free structural verification")

    status = "fail" if errors or (args.strict and warnings) else "pass"
    result = {
        "root": str(root),
        "mode": args.mode,
        "status": status,
        "errors": errors,
        "warnings": warnings,
        "release_state": release_state,
        "quality_gates": gates,
        "pptx": pptx_info,
    }
    if args.as_json:
        print(json.dumps(result, indent=2, ensure_ascii=False))
    else:
        print(f"Figure 1 bundle: {root}")
        print(f"Status: {status}")
        for item in errors:
            print(f"ERROR: {item}")
        for item in warnings:
            print(f"WARNING: {item}")
        if pptx_info:
            print("PPTX objects: " + ", ".join(f"{key}={value}" for key, value in pptx_info.items()))
    return 1 if status == "fail" else 0


if __name__ == "__main__":
    sys.exit(main())
