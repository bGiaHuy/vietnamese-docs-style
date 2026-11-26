"""Independent curation validator, zero-leakage linter, solvability runner, and commitment sealer."""

from __future__ import annotations

import argparse
import csv
import hashlib
import html as html_lib
import json
import random
import re
from datetime import datetime
from pathlib import Path
from typing import Any

from eval_harness.grader import grade_case
from eval_harness.ooxml_reader import inspect_docx
from eval_harness.render_inspector import render_pages_png, render_to_pdf

_FORBIDDEN_PROMPT_PATTERNS: list[tuple[str, re.Pattern[str]]] = [
    ("vndocs", re.compile(r"\bvndocs\b", re.IGNORECASE)),
    ("layout-edit", re.compile(r"\blayout[-_]edit\b", re.IGNORECASE)),
    ("map_docx", re.compile(r"\bmap_docx\b", re.IGNORECASE)),
    ("diff-layout", re.compile(r"\bdiff[-_]layout\b", re.IGNORECASE)),
    ("auto-paginate", re.compile(r"\bauto[-_]paginate\b", re.IGNORECASE)),
    ("validate_docx", re.compile(r"\bvalidate_docx\b", re.IGNORECASE)),
    ("clear_cell_overrides", re.compile(r"\bclear_cell_overrides\b", re.IGNORECASE)),
    ("preserve_header_size", re.compile(r"\bpreserve_header_size\b", re.IGNORECASE)),
    ("w:xml_tag", re.compile(r"\bw:(?:tblpr|gridspan|tcw|tblw|ppr|rpr)\b", re.IGNORECASE)),
    ("mc:xml_tag", re.compile(r"\bmc:(?:ignorable|alternatecontent)\b", re.IGNORECASE)),
    ("styles.xml", re.compile(r"\bstyles\.xml\b", re.IGNORECASE)),
    ("document.xml", re.compile(r"\bdocument\.xml\b", re.IGNORECASE)),
    ("sz=", re.compile(r"\bsz\s*=", re.IGNORECASE)),
    ("dxa", re.compile(r"\bdxa\b", re.IGNORECASE)),
    ("emu", re.compile(r"\bemu\b", re.IGNORECASE)),
    ("anchor_id", re.compile(r"\b(?:T|P|IMG)#[A-Za-z0-9]+")),
]

_CASE_ID_RE = re.compile(r"^(VAL|SEAL)-[A-E][0-9]{2}_[a-z0-9_]+$")
_VALID_GROUPS = ("A", "B", "C", "D", "E")
_VALID_LENGTHS = ("short", "medium", "long", "very_long")
_VALID_ITYPE = ("descriptive", "inferential")
_VALID_OUTCOME = ("edit", "refuse")
_VALID_SOURCE_CATS = {
    "gov_portal_nd30",
    "legal_normative",
    "party_document",
    "university_academic",
    "enterprise_office",
    "public_open_template",
    "author_manual_office",
}
_VALID_APPS = {
    "microsoft_word",
    "google_docs",
    "libreoffice_writer",
    "wps_office",
    "mixed_or_unknown",
}

_SPLIT_QUOTAS = {
    "val-v1": {
        "total": 20,
        "per_group": 4,
        "min_near_boundary": 2,
        "refuse_range": (3, 5),
        "min_itype_ratio": 0.40,
    },
    "test-sealed-v1": {
        "total": 30,
        "per_group": 6,
        "min_near_boundary": 3,
        "refuse_range": (4, 7),
        "min_itype_ratio": 0.40,
    },
}


def _sha256_file(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def lint_prompt_zero_leakage(prompt_vi: str) -> list[str]:
    """Return forbidden internal token hits found in prompt_vi."""
    hits: list[str] = []
    for label, pat in _FORBIDDEN_PROMPT_PATTERNS:
        if pat.search(prompt_vi or ""):
            hits.append(label)
    return hits


def _load_cases(manifest_path: str | Path) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    raw = json.loads(Path(manifest_path).read_text(encoding="utf-8"))
    if isinstance(raw, dict) and "cases" in raw:
        return raw, raw["cases"]
    if isinstance(raw, list):
        return {"cases": raw}, raw
    raise ValueError(f"Unsupported manifest structure in {manifest_path}")


def validate_manifest(
    manifest_path: str | Path,
    *,
    corpus_dir: str | Path | None = None,
    enforce_quotas: bool = True,
) -> dict[str, Any]:
    """Validate an independent evaluation manifest against schema rules, quotas, and local files."""
    mpath = Path(manifest_path)
    doc, cases = _load_cases(mpath)
    errors: list[str] = []
    seen_ids: set[str] = set()
    splits_seen: set[str] = set()
    source_hashes_seen: dict[str, list[str]] = {}

    group_counts = {g: 0 for g in _VALID_GROUPS}
    length_counts = {lt: 0 for lt in _VALID_LENGTHS}
    itype_counts = {it: 0 for it in _VALID_ITYPE}
    outcome_counts = {oc: 0 for oc in _VALID_OUTCOME}
    near_boundary_count = 0

    cdir = Path(corpus_dir) if corpus_dir is not None else None
    check_files = cdir is not None and cdir.is_dir()

    for idx, c in enumerate(cases):
        cid = c.get("case_id")
        if not cid or not _CASE_ID_RE.match(str(cid)):
            errors.append(f"case[{idx}]: invalid or missing case_id {cid!r}")
            cid = f"case_{idx}"
        elif cid in seen_ids:
            errors.append(f"{cid}: duplicate case_id")
        seen_ids.add(cid)

        if c.get("schema_version") != "1.0":
            errors.append(f"{cid}: schema_version must be '1.0'")

        split = c.get("split")
        if split not in _SPLIT_QUOTAS:
            errors.append(f"{cid}: invalid split {split!r}")
        else:
            splits_seen.add(split)

        grp = c.get("taxonomy_group")
        if grp not in group_counts:
            errors.append(f"{cid}: invalid taxonomy_group {grp!r}")
        else:
            group_counts[grp] += 1

        lt = c.get("length_tier")
        if lt not in length_counts:
            errors.append(f"{cid}: invalid length_tier {lt!r}")
        else:
            length_counts[lt] += 1

        it = c.get("instruction_type")
        if it not in itype_counts:
            errors.append(f"{cid}: invalid instruction_type {it!r}")
        else:
            itype_counts[it] += 1

        oc = c.get("expected_outcome")
        if oc not in outcome_counts:
            errors.append(f"{cid}: invalid expected_outcome {oc!r}")
        else:
            outcome_counts[oc] += 1

        if not isinstance(c.get("near_boundary_valid"), bool):
            errors.append(f"{cid}: near_boundary_valid must be bool")
        elif c["near_boundary_valid"]:
            if oc != "edit":
                errors.append(f"{cid}: near_boundary_valid=true requires expected_outcome='edit'")
            near_boundary_count += 1

        prov = c.get("provenance")
        if not isinstance(prov, dict):
            errors.append(f"{cid}: missing provenance object")
        else:
            if prov.get("created_without_vndocs_builders") is not True:
                errors.append(f"{cid}: provenance.created_without_vndocs_builders must be true")
            if prov.get("source_category") not in _VALID_SOURCE_CATS:
                errors.append(f"{cid}: invalid source_category {prov.get('source_category')!r}")
            if prov.get("authoring_software") not in _VALID_APPS:
                errors.append(f"{cid}: invalid authoring_software {prov.get('authoring_software')!r}")
            if not str(prov.get("source_cluster_id") or "").strip():
                errors.append(f"{cid}: empty source_cluster_id")

        prompt = c.get("prompt_vi") or ""
        if len(prompt.strip()) < 15:
            errors.append(f"{cid}: prompt_vi too short")
        leaks = lint_prompt_zero_leakage(prompt)
        if leaks:
            errors.append(f"{cid}: prompt_vi leakage tokens detected: {leaks}")

        grading = c.get("grading")
        if not isinstance(grading, dict) or not grading:
            errors.append(f"{cid}: missing grading object")

        rel = c.get("source_docx_relpath")
        if not rel:
            errors.append(f"{cid}: missing source_docx_relpath")
        elif check_files:
            fpath = cdir / rel
            if not fpath.is_file():
                errors.append(f"{cid}: missing source file {fpath}")
            else:
                actual_sha = _sha256_file(fpath)
                decl_sha = c.get("source_docx_sha256")
                if decl_sha and decl_sha != actual_sha:
                    errors.append(f"{cid}: source_docx_sha256 mismatch ({decl_sha} != {actual_sha})")
                source_hashes_seen.setdefault(actual_sha, []).append(cid)
                try:
                    inspect_docx(fpath)
                except Exception as exc:
                    errors.append(f"{cid}: invalid DOCX ({exc})")

            for att in c.get("attachments") or []:
                arel = att.get("relpath")
                ap = cdir / arel if arel else None
                if ap is None or not ap.is_file():
                    errors.append(f"{cid}: missing attachment {ap}")
                elif att.get("sha256") and _sha256_file(ap) != att["sha256"]:
                    errors.append(f"{cid}: attachment sha256 mismatch for {arel}")

    if len(splits_seen) > 1:
        errors.append(f"Mixed splits in single manifest: {sorted(splits_seen)}")

    if enforce_quotas and len(splits_seen) == 1:
        split_name = next(iter(splits_seen))
        q = _SPLIT_QUOTAS[split_name]
        if len(cases) != q["total"]:
            errors.append(f"Expected {q['total']} cases for {split_name}, got {len(cases)}")
        for g, cnt in group_counts.items():
            if cnt != q["per_group"]:
                errors.append(f"Group {g} count={cnt}, expected {q['per_group']}")
        for lt, cnt in length_counts.items():
            if cnt < 1:
                errors.append(f"Length tier {lt!r} has 0 cases")
        min_it = int(q["total"] * q["min_itype_ratio"])
        for it, cnt in itype_counts.items():
            if cnt < min_it:
                errors.append(f"Instruction type {it!r} count={cnt} < {min_it}")
        r_lo, r_hi = q["refuse_range"]
        if not (r_lo <= outcome_counts["refuse"] <= r_hi):
            errors.append(f"Refuse count={outcome_counts['refuse']} outside [{r_lo}, {r_hi}]")
        if near_boundary_count < q["min_near_boundary"]:
            errors.append(f"near_boundary_valid count={near_boundary_count} < {q['min_near_boundary']}")

    dup_sources = {h: cids for h, cids in source_hashes_seen.items() if len(cids) > 1}
    if dup_sources:
        errors.append(f"Duplicate source DOCX files detected: {dup_sources}")

    return {
        "valid": len(errors) == 0,
        "manifest_path": str(mpath),
        "files_checked": check_files,
        "case_count": len(cases),
        "group_counts": group_counts,
        "length_counts": length_counts,
        "instruction_type_counts": itype_counts,
        "outcome_counts": outcome_counts,
        "near_boundary_valid_count": near_boundary_count,
        "unique_source_clusters": len({
            c.get("provenance", {}).get("source_cluster_id")
            for c in cases
            if isinstance(c.get("provenance"), dict)
        }),
        "errors": errors,
    }


def check_disjoint(
    val_manifest_path: str | Path,
    sealed_manifest_path: str | Path,
) -> dict[str, Any]:
    """Verify zero source_cluster_id and zero source_docx_sha256 overlap between val and sealed."""
    _, val_cases = _load_cases(val_manifest_path)
    _, seal_cases = _load_cases(sealed_manifest_path)

    val_clusters = {
        c.get("provenance", {}).get("source_cluster_id")
        for c in val_cases
        if isinstance(c.get("provenance"), dict) and c["provenance"].get("source_cluster_id")
    }
    seal_clusters = {
        c.get("provenance", {}).get("source_cluster_id")
        for c in seal_cases
        if isinstance(c.get("provenance"), dict) and c["provenance"].get("source_cluster_id")
    }

    val_hashes = {c.get("source_docx_sha256") for c in val_cases if c.get("source_docx_sha256")}
    seal_hashes = {c.get("source_docx_sha256") for c in seal_cases if c.get("source_docx_sha256")}

    overlap_clusters = sorted(val_clusters & seal_clusters)
    overlap_hashes = sorted(val_hashes & seal_hashes)
    return {
        "disjoint": len(overlap_clusters) == 0 and len(overlap_hashes) == 0,
        "val_clusters": len(val_clusters),
        "sealed_clusters": len(seal_clusters),
        "overlapping_clusters": overlap_clusters,
        "overlapping_hashes": overlap_hashes,
    }


def verify_solvability(
    manifest_path: str | Path,
    corpus_dir: str | Path,
    positive_dir: str | Path,
    negative_dir: str | Path,
    *,
    run_render: bool = True,
) -> dict[str, Any]:
    """Run Independent Solvability Gate: positive non-vndocs must PASS, negative mutation must FAIL."""
    _, cases = _load_cases(manifest_path)
    cdir = Path(corpus_dir)
    pdir = Path(positive_dir)
    ndir = Path(negative_dir)

    rows: list[dict[str, Any]] = []
    pos_passes = 0
    neg_fails = 0
    for case in cases:
        cid = case["case_id"]
        src_docx = cdir / case["source_docx_relpath"]
        pos_res = grade_case(case, src_docx, pdir / cid, run_render=run_render)
        neg_res = grade_case(case, src_docx, ndir / cid, run_render=run_render)
        p_ok = pos_res["hit"] == 1
        n_ok = neg_res["hit"] == 0
        if p_ok:
            pos_passes += 1
        if n_ok:
            neg_fails += 1
        rows.append({
            "case_id": cid,
            "positive_hit": pos_res["hit"],
            "positive_failures": pos_res["failure_reasons"],
            "negative_hit": neg_res["hit"],
            "negative_failures": neg_res["failure_reasons"],
            "gate_passed": p_ok and n_ok,
        })

    return {
        "all_passed": pos_passes == len(cases) and neg_fails == len(cases),
        "total_cases": len(cases),
        "positive_passes": pos_passes,
        "negative_failures_caught": neg_fails,
        "cases": rows,
    }


def _compute_sealed_hashes(
    sroot: Path,
    sealed_manifest_relpath: str,
) -> tuple[str, str, list[dict[str, Any]]]:
    """Compute canonical bundle_sha256 and per_case_manifest_sha256 for a sealed directory."""
    mpath = sroot / sealed_manifest_relpath
    _, cases = _load_cases(mpath)
    bundle_hasher = hashlib.sha256()
    manifest_sha = _sha256_file(mpath)

    rel_files: list[Path] = [Path(sealed_manifest_relpath)]
    for c in cases:
        rel_files.append(Path(c["source_docx_relpath"]))
        for att in c.get("attachments") or []:
            rel_files.append(Path(att["relpath"]))

    seen_rel: set[str] = set()
    for rel in sorted(rel_files, key=lambda p: p.as_posix()):
        rel_posix = rel.as_posix()
        if rel_posix in seen_rel:
            continue
        seen_rel.add(rel_posix)
        data = (sroot / rel).read_bytes()
        bundle_hasher.update(rel_posix.encode("utf-8") + b"\x00")
        bundle_hasher.update(hashlib.sha256(data).hexdigest().encode("ascii") + b"\n")
    return bundle_hasher.hexdigest(), manifest_sha, cases


def seal_commitment(
    sealed_dir: str | Path,
    sealed_manifest_relpath: str,
    val_manifest_path: str | Path,
    output_commitment_path: str | Path,
    *,
    frozen_system_tag: str | None = None,
    frozen_system_commit_sha: str | None = None,
) -> dict[str, Any]:
    """Validate an external test-sealed-v1 bundle and write its SHA-256 commitment JSON."""
    sroot = Path(sealed_dir).resolve()
    mpath = sroot / sealed_manifest_relpath
    val_report = validate_manifest(mpath, corpus_dir=sroot, enforce_quotas=True)
    if not val_report["valid"]:
        raise ValueError(f"Sealed manifest failed validation: {val_report['errors']}")

    disjoint_res = check_disjoint(val_manifest_path, mpath)
    if not disjoint_res["disjoint"]:
        raise ValueError(f"Sealed split overlaps with val-v1: {disjoint_res}")

    bundle_sha, manifest_sha, cases = _compute_sealed_hashes(sroot, sealed_manifest_relpath)

    commitment = {
        "schema_version": "1.0",
        "status": "committed_awaiting_eval",
        "sealed_split_id": "test-sealed-v1",
        "protocol_doc": "evals/EVALUATION-PROTOCOL.md",
        "case_schema": "evals/schemas/independent-case.schema.json",
        "target_case_count": len(cases),
        "taxonomy_quota": val_report["group_counts"],
        "storage_policy": "Stored strictly outside development workspace; unsealed exactly once after frozen_system_tag is set.",
        "bundle_sha256": bundle_sha,
        "per_case_manifest_sha256": manifest_sha,
        "frozen_system_tag": frozen_system_tag,
        "frozen_system_commit_sha": frozen_system_commit_sha,
        "committed_at": datetime.now().astimezone().isoformat(),
        "evaluated_at": None,
    }
    out_path = Path(output_commitment_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(json.dumps(commitment, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return commitment


def verify_commitment(
    commitment_path: str | Path,
    sealed_dir: str | Path,
    sealed_manifest_relpath: str,
    *,
    allow_already_evaluated: bool = False,
) -> dict[str, Any]:
    """Verify sealed bundle hashes against commitment JSON and enforce single-use policy."""
    cpath = Path(commitment_path)
    commitment = json.loads(cpath.read_text(encoding="utf-8"))
    errors: list[str] = []

    if commitment.get("evaluated_at") is not None and not allow_already_evaluated:
        errors.append(f"Sealed commitment already evaluated at {commitment['evaluated_at']}")
    if not commitment.get("bundle_sha256") or not commitment.get("per_case_manifest_sha256"):
        errors.append("Commitment JSON has null SHA-256 hashes (not yet sealed).")

    sroot = Path(sealed_dir).resolve()
    actual_bundle_sha, actual_manifest_sha, cases = _compute_sealed_hashes(sroot, sealed_manifest_relpath)
    if commitment.get("bundle_sha256") and actual_bundle_sha != commitment["bundle_sha256"]:
        errors.append(
            f"bundle_sha256 mismatch: expected={commitment['bundle_sha256']} actual={actual_bundle_sha}"
        )
    if commitment.get("per_case_manifest_sha256") and actual_manifest_sha != commitment["per_case_manifest_sha256"]:
        errors.append(
            f"per_case_manifest_sha256 mismatch: expected={commitment['per_case_manifest_sha256']} actual={actual_manifest_sha}"
        )

    return {
        "valid": len(errors) == 0,
        "commitment_path": str(cpath),
        "case_count": len(cases),
        "bundle_sha256": actual_bundle_sha,
        "per_case_manifest_sha256": actual_manifest_sha,
        "already_evaluated": commitment.get("evaluated_at") is not None,
        "errors": errors,
    }


def mark_commitment_evaluated(
    commitment_path: str | Path,
    *,
    evaluated_at: str | None = None,
    notes: str | None = None,
) -> dict[str, Any]:
    """Stamp evaluated_at on a sealed commitment JSON so it cannot be re-run cleanly."""
    cpath = Path(commitment_path)
    commitment = json.loads(cpath.read_text(encoding="utf-8"))
    if not commitment.get("bundle_sha256") or not commitment.get("per_case_manifest_sha256"):
        raise ValueError("Cannot mark unsealed commitment (bundle_sha256 is null) as evaluated.")
    if commitment.get("evaluated_at") is not None:
        raise ValueError(f"Sealed commitment was already evaluated at {commitment['evaluated_at']}.")

    commitment["evaluated_at"] = evaluated_at or datetime.now().astimezone().isoformat()
    commitment["status"] = "evaluated_sealed_once"
    if notes:
        commitment["evaluation_notes"] = notes
    cpath.write_text(json.dumps(commitment, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return commitment


def prepare_blind_audit(
    run_items: list[dict[str, Any]],
    out_dir: str | Path,
    *,
    seed: int = 42,
) -> dict[str, Any]:
    """Export anonymized before/after PNGs and a blind scoring CSV without arm labels."""
    out_root = Path(out_dir)
    items_dir = out_root / "items"
    items_dir.mkdir(parents=True, exist_ok=True)

    shuffled = list(run_items)
    rng = random.Random(seed)
    rng.shuffle(shuffled)

    key_rows: list[dict[str, Any]] = []
    csv_rows: list[list[str]] = [
        [
            "blind_id",
            "case_id",
            "prompt_vi",
            "score_request_accuracy_1_to_5",
            "score_visual_aesthetics_1_to_5",
            "score_non_target_integrity_1_to_5",
            "notes",
        ]
    ]

    for idx, item in enumerate(shuffled, start=1):
        blind_id = f"BLIND-{idx:03d}"
        item_dir = items_dir / blind_id
        before_dir = item_dir / "before"
        after_dir = item_dir / "after"
        b_ren = render_to_pdf(item["src_docx"], before_dir)
        a_ren = render_to_pdf(item["dst_docx"], after_dir)
        b_pngs = render_pages_png(b_ren["pdf_path"], before_dir) if b_ren.get("status") == "rendered" else []
        a_pngs = render_pages_png(a_ren["pdf_path"], after_dir) if a_ren.get("status") == "rendered" else []

        key_rows.append({
            "blind_id": blind_id,
            "case_id": item["case_id"],
            "arm": item["arm"],
            "model": item.get("model"),
            "before_pngs": [str(Path(p).relative_to(out_root)) for p in b_pngs],
            "after_pngs": [str(Path(p).relative_to(out_root)) for p in a_pngs],
        })
        csv_rows.append([blind_id, item["case_id"], item.get("prompt_vi", ""), "", "", "", ""])

    (out_root / "key.json").write_text(json.dumps(key_rows, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    with (out_root / "scoring_sheet.csv").open("w", encoding="utf-8-sig", newline="") as fh:
        csv.writer(fh).writerows(csv_rows)
    viewer_path = render_blind_viewer_html(out_root)
    return {"item_count": len(key_rows), "out_dir": str(out_root), "viewer_html": str(viewer_path)}


def render_blind_viewer_html(out_dir: str | Path) -> Path:
    """Render a zero-leakage static HTML viewer from scoring_sheet.csv and items/ PNGs only."""
    out_root = Path(out_dir)
    csv_path = out_root / "scoring_sheet.csv"
    if not csv_path.is_file():
        raise FileNotFoundError(f"scoring_sheet.csv not found in {out_root}")

    rows: list[dict[str, Any]] = []
    with csv_path.open("r", encoding="utf-8-sig", newline="") as fh:
        reader = csv.DictReader(fh)
        for rec in reader:
            bid = rec["blind_id"]
            b_dir = out_root / "items" / bid / "before"
            a_dir = out_root / "items" / bid / "after"
            b_pngs = [p.relative_to(out_root).as_posix() for p in sorted(b_dir.glob("*.png"))] if b_dir.is_dir() else []
            a_pngs = [p.relative_to(out_root).as_posix() for p in sorted(a_dir.glob("*.png"))] if a_dir.is_dir() else []
            rows.append({
                "blind_id": bid,
                "case_id": rec["case_id"],
                "prompt_vi": rec["prompt_vi"],
                "before_pngs": b_pngs,
                "after_pngs": a_pngs,
                "score_request_accuracy_1_to_5": rec.get("score_request_accuracy_1_to_5", ""),
                "score_visual_aesthetics_1_to_5": rec.get("score_visual_aesthetics_1_to_5", ""),
                "score_non_target_integrity_1_to_5": rec.get("score_non_target_integrity_1_to_5", ""),
                "notes": rec.get("notes", ""),
            })

    cards: list[str] = []
    for item in rows:
        bid = html_lib.escape(item["blind_id"])
        cid = html_lib.escape(item["case_id"])
        prompt = html_lib.escape(item["prompt_vi"])
        b_imgs = "\n".join(
            f'<figure><figcaption>{html_lib.escape(Path(p).name)}</figcaption>'
            f'<img src="{html_lib.escape(p)}" loading="lazy" alt="Before"></figure>'
            for p in item["before_pngs"]
        ) or "<p class='empty'>No before PNG</p>"
        a_imgs = "\n".join(
            f'<figure><figcaption>{html_lib.escape(Path(p).name)}</figcaption>'
            f'<img src="{html_lib.escape(p)}" loading="lazy" alt="After"></figure>'
            for p in item["after_pngs"]
        ) or "<p class='empty'>No after PNG</p>"
        cards.append(
            f"""<section class="card" data-blind-id="{bid}">
  <header><h2>{bid} — <code>{cid}</code></h2></header>
  <p class="prompt">{prompt}</p>
  <div class="grid">
    <div class="col"><h3>Before (Original)</h3>{b_imgs}</div>
    <div class="col"><h3>After (Submitted)</h3>{a_imgs}</div>
  </div>
  <div class="scores">
    <label>Request Accuracy (1–5): <input type="number" min="1" max="5" data-field="score_request_accuracy_1_to_5" value="{html_lib.escape(item['score_request_accuracy_1_to_5'])}"></label>
    <label>Visual Aesthetics (1–5): <input type="number" min="1" max="5" data-field="score_visual_aesthetics_1_to_5" value="{html_lib.escape(item['score_visual_aesthetics_1_to_5'])}"></label>
    <label>Non-Target Integrity (1–5): <input type="number" min="1" max="5" data-field="score_non_target_integrity_1_to_5" value="{html_lib.escape(item['score_non_target_integrity_1_to_5'])}"></label>
    <label class="notes">Notes: <input type="text" data-field="notes" value="{html_lib.escape(item['notes'])}"></label>
  </div>
</section>"""
        )

    payload_json = json.dumps(
        [
            {"blind_id": r["blind_id"], "case_id": r["case_id"], "prompt_vi": r["prompt_vi"]}
            for r in rows
        ],
        ensure_ascii=False,
    ).replace("</", "<\\/")

    html_doc = f"""<!doctype html>
<html lang="vi">
<head>
<meta charset="utf-8">
<title>Blind Human Visual Audit Viewer ({len(rows)} items)</title>
<style>
  body {{ font-family: system-ui, -apple-system, sans-serif; margin: 24px; background: #f7f8fa; color: #1f2937; }}
  .topbar {{ position: sticky; top: 0; background: #fff; padding: 12px 16px; border: 1px solid #d1d5db; border-radius: 8px; display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px; z-index: 10; }}
  button {{ background: #2563eb; color: #fff; border: 0; padding: 8px 14px; border-radius: 6px; font-weight: 600; cursor: pointer; }}
  .card {{ background: #fff; border: 1px solid #e5e7eb; border-radius: 10px; padding: 16px; margin-bottom: 24px; }}
  .prompt {{ background: #f3f4f6; padding: 10px 12px; border-radius: 6px; }}
  .grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 16px; margin: 12px 0; }}
  .col {{ border: 1px solid #e5e7eb; border-radius: 8px; padding: 8px; max-height: 760px; overflow-y: auto; background: #fafafa; }}
  figure {{ margin: 0 0 12px 0; }}
  figcaption {{ font-size: 12px; color: #6b7280; margin-bottom: 4px; }}
  img {{ width: 100%; border: 1px solid #d1d5db; border-radius: 4px; }}
  .scores {{ display: flex; flex-wrap: wrap; gap: 12px; padding-top: 8px; border-top: 1px solid #e5e7eb; }}
  .scores label {{ font-size: 13px; font-weight: 500; }}
  .scores input[type="number"] {{ width: 56px; padding: 4px; margin-left: 4px; }}
  .scores label.notes {{ flex: 1 1 280px; display: flex; gap: 6px; align-items: center; }}
  .scores label.notes input {{ flex: 1; padding: 4px 6px; }}
</style>
</head>
<body>
<div class="topbar">
  <div><strong>Blind Human Visual Audit</strong> — {len(rows)} anonymized items (arm/model hidden)</div>
  <button onclick="downloadCsv()">Download scoring_sheet.csv</button>
</div>
{''.join(cards)}
<script>
const ITEMS = {payload_json};
function csvEscape(v) {{
  const s = String(v ?? '');
  return /[",\\r\\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
}}
function downloadCsv() {{
  const hdr = ['blind_id','case_id','prompt_vi','score_request_accuracy_1_to_5','score_visual_aesthetics_1_to_5','score_non_target_integrity_1_to_5','notes'];
  const lines = [hdr.join(',')];
  for (const item of ITEMS) {{
    const sec = document.querySelector('section[data-blind-id="' + item.blind_id + '"]');
    const val = (f) => sec.querySelector('input[data-field="' + f + '"]').value;
    lines.push([
      item.blind_id, item.case_id, item.prompt_vi,
      val('score_request_accuracy_1_to_5'),
      val('score_visual_aesthetics_1_to_5'),
      val('score_non_target_integrity_1_to_5'),
      val('notes'),
    ].map(csvEscape).join(','));
  }}
  const blob = new Blob(['\\uFEFF' + lines.join('\\r\\n') + '\\r\\n'], {{ type: 'text/csv;charset=utf-8;' }});
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'scoring_sheet.csv';
  a.click();
}}
</script>
</body>
</html>
"""
    viewer_path = out_root / "viewer.html"
    viewer_path.write_text(html_doc, encoding="utf-8")
    return viewer_path


def prepare_blind_audit_from_report(
    report_path: str | Path,
    manifest_path: str | Path,
    corpus_dir: str | Path,
    out_dir: str | Path,
    *,
    runs_manifest_path: str | Path | None = None,
    max_cases: int = 12,
    seed: int = 42,
    arms: list[str] | None = None,
) -> dict[str, Any]:
    """Select stratified edited cases from a benchmark report and export a blind visual audit pack."""
    rpath = Path(report_path)
    rep = json.loads(rpath.read_text(encoding="utf-8"))
    _, cases = _load_cases(manifest_path)
    case_map = {c["case_id"]: c for c in cases}

    rm_path = Path(runs_manifest_path) if runs_manifest_path else (rpath.parent / "runs_manifest.json")
    if not rm_path.is_file():
        raise FileNotFoundError(f"runs_manifest.json not found near {rpath}: {rm_path}")
    rm_data = json.loads(rm_path.read_text(encoding="utf-8"))
    run_entries = rm_data["runs"] if isinstance(rm_data, dict) and "runs" in rm_data else rm_data
    wdir_map = {(e["case_id"], e["arm"]): Path(e["work_dir"]) for e in run_entries}

    wanted_arms = set(arms or ["arm0_no_skill", "arm1_standards_only", "arm2_full_skill"])
    # Stratify edited cases round-robin across visual taxonomy groups (preferring A, B, C, E).
    by_group: dict[str, list[str]] = {}
    for cid, case in sorted(case_map.items()):
        if case.get("expected_outcome") != "edit":
            continue
        by_group.setdefault(case.get("taxonomy_group", "A"), []).append(cid)

    group_order = [g for g in ("A", "B", "C", "E", "D") if g in by_group]
    selected_cases: list[str] = []
    idx_by_group = {g: 0 for g in group_order}
    while len(selected_cases) < max_cases:
        progressed = False
        for g in group_order:
            pos = idx_by_group[g]
            if pos < len(by_group[g]):
                selected_cases.append(by_group[g][pos])
                idx_by_group[g] = pos + 1
                progressed = True
                if len(selected_cases) >= max_cases:
                    break
        if not progressed:
            break

    selected_set = set(selected_cases)
    cdir = Path(corpus_dir)
    run_items: list[dict[str, Any]] = []
    for row in rep.get("rows", []):
        cid = row["case_id"]
        arm = row["arm"]
        if cid not in selected_set or arm not in wanted_arms:
            continue
        wdir = wdir_map.get((cid, arm))
        if wdir is None:
            continue
        dst_docx = wdir / f"{cid}.docx"
        if not dst_docx.is_file():
            continue
        case = case_map[cid]
        run_items.append(
            {
                "case_id": cid,
                "arm": arm,
                "model": row.get("model_id"),
                "prompt_vi": case.get("prompt_vi", ""),
                "src_docx": str(cdir / case["source_docx_relpath"]),
                "dst_docx": str(dst_docx),
            }
        )

    res = prepare_blind_audit(run_items, out_dir, seed=seed)
    res["selected_cases"] = sorted(selected_set)
    return res


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_lint = sub.add_parser("lint-prompts", help="Check prompts for zero internal leakage")
    p_lint.add_argument("--manifest", required=True)

    p_val = sub.add_parser("validate-manifest", help="Validate manifest schema, quotas, and local files")
    p_val.add_argument("--manifest", required=True)
    p_val.add_argument("--corpus-dir")
    p_val.add_argument("--no-quotas", action="store_true")

    p_disj = sub.add_parser("check-disjoint", help="Check source disjointness between val and sealed")
    p_disj.add_argument("--val-manifest", required=True)
    p_disj.add_argument("--sealed-manifest", required=True)

    p_solv = sub.add_parser("verify-solvability", help="Run Independent Solvability Gate")
    p_solv.add_argument("--manifest", required=True)
    p_solv.add_argument("--corpus-dir", required=True)
    p_solv.add_argument("--positive-dir", required=True)
    p_solv.add_argument("--negative-dir", required=True)
    p_solv.add_argument("--no-render", action="store_true")
    p_solv.add_argument("--output")

    p_seal = sub.add_parser("seal-commitment", help="Compute and write SEALED-TEST-COMMITMENT.json")
    p_seal.add_argument("--sealed-dir", required=True)
    p_seal.add_argument("--sealed-manifest-relpath", required=True)
    p_seal.add_argument("--val-manifest", default="evals/val_v1/val_v1_cases.json")
    p_seal.add_argument("--output", default="evals/SEALED-TEST-COMMITMENT.json")
    p_seal.add_argument("--frozen-tag")
    p_seal.add_argument("--frozen-sha")

    p_vcom = sub.add_parser("verify-commitment", help="Verify sealed SHA-256 hashes and single-use status")
    p_vcom.add_argument("--commitment", default="evals/SEALED-TEST-COMMITMENT.json")
    p_vcom.add_argument("--sealed-dir", required=True)
    p_vcom.add_argument("--sealed-manifest-relpath", required=True)
    p_vcom.add_argument("--allow-already-evaluated", action="store_true")

    p_mark = sub.add_parser("mark-evaluated", help="Stamp evaluated_at on SEALED-TEST-COMMITMENT.json")
    p_mark.add_argument("--commitment", default="evals/SEALED-TEST-COMMITMENT.json")
    p_mark.add_argument("--evaluated-at", default=None)
    p_mark.add_argument("--notes", default=None)

    p_blind = sub.add_parser("prepare-blind-audit", help="Export blind before/after PNG visual audit pack")
    p_blind.add_argument("--report", required=True)
    p_blind.add_argument("--manifest", required=True)
    p_blind.add_argument("--corpus-dir", required=True)
    p_blind.add_argument("--out-dir", required=True)
    p_blind.add_argument("--runs-manifest", default=None)
    p_blind.add_argument("--max-cases", type=int, default=12)
    p_blind.add_argument("--seed", type=int, default=42)
    p_blind.add_argument("--arms", nargs="+", default=None)

    args = parser.parse_args(argv)
    if args.cmd == "lint-prompts":
        _, cases = _load_cases(args.manifest)
        leaks = {c["case_id"]: lint_prompt_zero_leakage(c.get("prompt_vi", "")) for c in cases}
        leaks = {k: v for k, v in leaks.items() if v}
        out: Any = {"clean": len(leaks) == 0, "leaks": leaks}
    elif args.cmd == "validate-manifest":
        out = validate_manifest(
            args.manifest,
            corpus_dir=args.corpus_dir,
            enforce_quotas=not args.no_quotas,
        )
    elif args.cmd == "check-disjoint":
        out = check_disjoint(args.val_manifest, args.sealed_manifest)
    elif args.cmd == "verify-solvability":
        out = verify_solvability(
            args.manifest,
            args.corpus_dir,
            args.positive_dir,
            args.negative_dir,
            run_render=not args.no_render,
        )
        if args.output:
            op = Path(args.output)
            op.parent.mkdir(parents=True, exist_ok=True)
            op.write_text(json.dumps(out, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    elif args.cmd == "verify-commitment":
        out = verify_commitment(
            args.commitment,
            args.sealed_dir,
            args.sealed_manifest_relpath,
            allow_already_evaluated=args.allow_already_evaluated,
        )
    elif args.cmd == "mark-evaluated":
        out = mark_commitment_evaluated(
            args.commitment,
            evaluated_at=args.evaluated_at,
            notes=args.notes,
        )
    elif args.cmd == "prepare-blind-audit":
        out = prepare_blind_audit_from_report(
            args.report,
            args.manifest,
            args.corpus_dir,
            args.out_dir,
            runs_manifest_path=args.runs_manifest,
            max_cases=args.max_cases,
            seed=args.seed,
            arms=args.arms,
        )
    else:
        out = seal_commitment(
            args.sealed_dir,
            args.sealed_manifest_relpath,
            args.val_manifest,
            args.output,
            frozen_system_tag=args.frozen_tag,
            frozen_system_commit_sha=args.frozen_sha,
        )

    print(json.dumps(out, ensure_ascii=False, indent=2))
    if args.cmd == "verify-commitment" and not out.get("valid"):
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
