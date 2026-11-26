"""Independent curation validator, zero-leakage linter, solvability runner, and commitment sealer."""

from __future__ import annotations

import argparse
import csv
import hashlib
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
        fpath = sroot / rel
        data = fpath.read_bytes()
        bundle_hasher.update(rel_posix.encode("utf-8") + b"\x00")
        bundle_hasher.update(hashlib.sha256(data).hexdigest().encode("ascii") + b"\n")

    commitment = {
        "schema_version": "1.0",
        "status": "committed_awaiting_eval",
        "sealed_split_id": "test-sealed-v1",
        "protocol_doc": "evals/EVALUATION-PROTOCOL.md",
        "case_schema": "evals/schemas/independent-case.schema.json",
        "target_case_count": len(cases),
        "taxonomy_quota": val_report["group_counts"],
        "storage_policy": "Stored strictly outside development workspace; unsealed exactly once after frozen_system_tag is set.",
        "bundle_sha256": bundle_hasher.hexdigest(),
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
    return {"item_count": len(key_rows), "out_dir": str(out_root)}


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
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
