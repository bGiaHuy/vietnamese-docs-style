"""Pure-stdlib statistical analysis for multi-trial and multi-arm benchmarks."""

from __future__ import annotations

import math
import random
import statistics
from typing import Any


def pass_at_1(hits_by_case: dict[str, list[int]] | list[int]) -> float:
    """Compute mean per-trial pass rate across cases."""
    if isinstance(hits_by_case, list):
        return round(sum(hits_by_case) / len(hits_by_case), 4) if hits_by_case else 0.0
    per_case = [sum(v) / len(v) for v in hits_by_case.values() if v]
    return round(sum(per_case) / len(per_case), 4) if per_case else 0.0


def pass_hat_k(hits_by_case: dict[str, list[int]], k: int) -> float:
    """Compute proportion of cases passing all first k independent trials."""
    if k <= 0 or not hits_by_case:
        return 0.0
    eligible = [v[:k] for v in hits_by_case.values() if len(v) >= k]
    if not eligible:
        return 0.0
    passed_all = sum(1 for trials in eligible if all(h == 1 for h in trials))
    return round(passed_all / len(eligible), 4)


def normalized_gain(base_rate: float, treat_rate: float) -> dict[str, Any]:
    """Compute normalized gain g while explicitly flagging ceiling saturation."""
    if base_rate >= 1.0:
        return {"g": None, "ceiling_saturated": True}
    g = (treat_rate - base_rate) / (1.0 - base_rate)
    return {"g": round(g, 4), "ceiling_saturated": False}


def bootstrap_ci(
    values: list[float],
    *,
    n_resamples: int = 10000,
    alpha: float = 0.05,
    seed: int = 42,
) -> dict[str, Any]:
    """Compute deterministic percentile bootstrap CI for the mean of values."""
    clean = [float(v) for v in values if v is not None]
    n = len(clean)
    if n == 0:
        return {"n": 0, "mean": 0.0, "stddev": 0.0, "ci_low": 0.0, "ci_high": 0.0}
    mean_val = sum(clean) / n
    std_val = math.sqrt(sum((x - mean_val) ** 2 for x in clean) / (n - 1)) if n > 1 else 0.0
    if n == 1:
        return {
            "n": 1,
            "mean": round(mean_val, 4),
            "stddev": 0.0,
            "ci_low": round(mean_val, 4),
            "ci_high": round(mean_val, 4),
        }

    rng = random.Random(seed)
    boot_means = sorted(
        sum(rng.choices(clean, k=n)) / n
        for _ in range(n_resamples)
    )
    lo_idx = max(0, min(n_resamples - 1, int(math.floor((alpha / 2.0) * n_resamples))))
    hi_idx = max(0, min(n_resamples - 1, int(math.ceil((1.0 - alpha / 2.0) * n_resamples)) - 1))
    return {
        "n": n,
        "mean": round(mean_val, 4),
        "stddev": round(std_val, 4),
        "ci_low": round(boot_means[lo_idx], 4),
        "ci_high": round(boot_means[hi_idx], 4),
    }


def mcnemar_exact(a_hits: list[int], b_hits: list[int]) -> dict[str, Any]:
    """Exact two-sided McNemar test on paired binary outcomes."""
    if len(a_hits) != len(b_hits):
        raise ValueError("Paired arrays must have equal length")
    b_only_a = sum(1 for a, b in zip(a_hits, b_hits) if a == 1 and b == 0)
    c_only_b = sum(1 for a, b in zip(a_hits, b_hits) if a == 0 and b == 1)
    n_disc = b_only_a + c_only_b
    if n_disc == 0:
        p_val = 1.0
    else:
        tail = min(b_only_a, c_only_b)
        prob = sum(math.comb(n_disc, i) for i in range(tail + 1)) / (2 ** n_disc)
        p_val = min(1.0, 2.0 * prob)
    return {
        "n_pairs": len(a_hits),
        "b_a_only": b_only_a,
        "c_b_only": c_only_b,
        "n_discordant": n_disc,
        "p_value": round(p_val, 8),
    }


def _average_ranks_with_ties(abs_diffs: list[float]) -> tuple[list[float], float]:
    n = len(abs_diffs)
    indexed = sorted(enumerate(abs_diffs), key=lambda x: x[1])
    ranks = [0.0] * n
    tie_corr = 0.0
    i = 0
    while i < n:
        j = i + 1
        while j < n and indexed[j][1] == indexed[i][1]:
            j += 1
        avg_rank = (i + 1 + j) / 2.0
        t = j - i
        if t > 1:
            tie_corr += (t ** 3 - t)
        for k in range(i, j):
            ranks[indexed[k][0]] = avg_rank
        i = j
    return ranks, tie_corr


def wilcoxon_signed_rank(a_vals: list[float], b_vals: list[float]) -> dict[str, Any]:
    """Two-sided paired Wilcoxon signed-rank test (exact for n<=16, normal approx otherwise)."""
    if len(a_vals) != len(b_vals):
        raise ValueError("Paired arrays must have equal length")
    diffs = [float(b) - float(a) for a, b in zip(a_vals, b_vals)]
    med_delta = round(statistics.median(diffs), 4) if diffs else 0.0
    nz = [d for d in diffs if abs(d) > 1e-12]
    n = len(nz)
    if n == 0:
        return {
            "n_pairs": len(diffs),
            "n_nonzero": 0,
            "median_delta": med_delta,
            "w_plus": 0.0,
            "w_minus": 0.0,
            "rank_biserial_r": 0.0,
            "p_value": 1.0,
            "method": "degenerate_all_zero",
        }

    abs_d = [abs(d) for d in nz]
    ranks, tie_corr = _average_ranks_with_ties(abs_d)
    w_plus = sum(r for d, r in zip(nz, ranks) if d > 0)
    total_r = n * (n + 1) / 2.0
    w_minus = total_r - w_plus
    w_stat = min(w_plus, w_minus)
    r_eff = round((w_plus - w_minus) / total_r, 4)

    if n <= 16:
        # Exact enumeration over all 2^n sign assignments (handles ties accurately).
        extreme = 0
        total_assign = 1 << n
        hi_thresh = total_r - w_stat - 1e-9
        lo_thresh = w_stat + 1e-9
        for mask in range(total_assign):
            s = sum(ranks[idx] for idx in range(n) if (mask >> idx) & 1)
            if s <= lo_thresh or s >= hi_thresh:
                extreme += 1
        p_val = min(1.0, extreme / total_assign)
        method = "exact"
    else:
        mean_w = total_r / 2.0
        var_w = (n * (n + 1) * (2 * n + 1) - tie_corr / 2.0) / 24.0
        if var_w <= 0:
            p_val = 1.0
        else:
            # Continuity correction toward mean.
            z = (abs(w_stat - mean_w) - 0.5) / math.sqrt(var_w)
            z = max(0.0, z)
            p_val = math.erfc(z / math.sqrt(2.0))
        method = "normal_approx"

    return {
        "n_pairs": len(diffs),
        "n_nonzero": n,
        "median_delta": med_delta,
        "w_plus": round(w_plus, 2),
        "w_minus": round(w_minus, 2),
        "rank_biserial_r": r_eff,
        "p_value": round(p_val, 8),
        "method": method,
    }


def _arm_rates(rows: list[dict[str, Any]]) -> dict[str, Any]:
    by_case: dict[str, list[int]] = {}
    for r in rows:
        by_case.setdefault(r["case_id"], []).append(int(r["hit"]))
    k_vals = [len(v) for v in by_case.values() if v]
    min_k = min(k_vals) if k_vals else 0

    edit_rows = [r for r in rows if r.get("expected_outcome") == "edit"]
    ref_rows = [r for r in rows if r.get("expected_outcome") == "refuse"]
    over_ref = sum(1 for r in edit_rows if r.get("over_refusal"))
    unsafe_ed = sum(1 for r in ref_rows if r.get("unsafe_edit"))
    collat = sum(
        1 for r in edit_rows
        if r.get("checkpoints", {}).get("semantic_non_target_preserved") is False
    )

    hits = [int(r["hit"]) for r in rows]
    cps = [float(r["checkpoint_score"]) for r in rows]
    tot_tok = [float(r["transcript"]["total_tokens"]) for r in rows if isinstance(r.get("transcript"), dict) and not r["transcript"].get("missing", True)]
    fresh_out = [float(r["transcript"]["fresh_plus_output_tokens"]) for r in rows if isinstance(r.get("transcript"), dict) and not r["transcript"].get("missing", True)]
    calls = [float(r["transcript"]["tool_calls_total"]) for r in rows if isinstance(r.get("transcript"), dict) and not r["transcript"].get("missing", True)]
    secs = [float(r["transcript"]["elapsed_seconds"]) for r in rows if isinstance(r.get("transcript"), dict) and not r["transcript"].get("missing", True)]

    return {
        "n_runs": len(rows),
        "n_cases": len(by_case),
        "min_trials_per_case": min_k,
        "pass_at_1": pass_at_1(by_case),
        "pass_hat_k": pass_hat_k(by_case, min_k) if min_k >= 1 else 0.0,
        "hit_ci_95": bootstrap_ci(hits),
        "checkpoint_score": bootstrap_ci(cps),
        "non_target_collateral_rate": round(collat / len(edit_rows), 4) if edit_rows else 0.0,
        "over_refusal_rate": round(over_ref / len(edit_rows), 4) if edit_rows else 0.0,
        "unsafe_edit_rate": round(unsafe_ed / len(ref_rows), 4) if ref_rows else 0.0,
        "total_tokens": bootstrap_ci(tot_tok) if tot_tok else None,
        "fresh_plus_output_tokens": bootstrap_ci(fresh_out) if fresh_out else None,
        "tool_calls": bootstrap_ci(calls) if calls else None,
        "elapsed_seconds": bootstrap_ci(secs) if secs else None,
    }


def compare_arms(arm_a_rows: list[dict[str, Any]], arm_b_rows: list[dict[str, Any]]) -> dict[str, Any]:
    """Compare two arms paired by (case_id, trial_index) or case-level means."""
    def _key(r: dict[str, Any]) -> tuple[str, int]:
        return (r["case_id"], int(r.get("trial_index", 0)))

    map_a = {_key(r): r for r in arm_a_rows}
    map_b = {_key(r): r for r in arm_b_rows}
    shared = sorted(set(map_a) & set(map_b))

    a_hits = [int(map_a[k]["hit"]) for k in shared]
    b_hits = [int(map_b[k]["hit"]) for k in shared]
    a_cps = [float(map_a[k]["checkpoint_score"]) for k in shared]
    b_cps = [float(map_b[k]["checkpoint_score"]) for k in shared]

    sum_a = _arm_rates(arm_a_rows)
    sum_b = _arm_rates(arm_b_rows)
    gain = normalized_gain(sum_a["pass_at_1"], sum_b["pass_at_1"])

    out: dict[str, Any] = {
        "n_paired_runs": len(shared),
        "arm_a": sum_a,
        "arm_b": sum_b,
        "normalized_gain": gain,
        "hit_rate_pp_delta": round((sum_b["pass_at_1"] - sum_a["pass_at_1"]) * 100.0, 2),
        "mcnemar_hit": mcnemar_exact(a_hits, b_hits) if shared else None,
        "wilcoxon_checkpoint": wilcoxon_signed_rank(a_cps, b_cps) if shared else None,
    }

    tok_pairs = [
        (map_a[k]["transcript"]["total_tokens"], map_b[k]["transcript"]["total_tokens"])
        for k in shared
        if isinstance(map_a[k].get("transcript"), dict) and isinstance(map_b[k].get("transcript"), dict)
        and not map_a[k]["transcript"].get("missing", True) and not map_b[k]["transcript"].get("missing", True)
    ]
    if tok_pairs:
        out["wilcoxon_total_tokens"] = wilcoxon_signed_rank([p[0] for p in tok_pairs], [p[1] for p in tok_pairs])

    fo_pairs = [
        (map_a[k]["transcript"]["fresh_plus_output_tokens"], map_b[k]["transcript"]["fresh_plus_output_tokens"])
        for k in shared
        if isinstance(map_a[k].get("transcript"), dict) and isinstance(map_b[k].get("transcript"), dict)
        and not map_a[k]["transcript"].get("missing", True) and not map_b[k]["transcript"].get("missing", True)
    ]
    if fo_pairs:
        out["wilcoxon_fresh_plus_output_tokens"] = wilcoxon_signed_rank([p[0] for p in fo_pairs], [p[1] for p in fo_pairs])

    return out


def summarize_experiment(all_rows: list[dict[str, Any]]) -> dict[str, Any]:
    """Summarize experiment rows across arms and stratification dimensions."""
    arms = sorted({r.get("arm", "unknown") for r in all_rows})
    by_arm = {arm: _arm_rates([r for r in all_rows if r.get("arm") == arm]) for arm in arms}

    def _slice(field: str) -> dict[str, Any]:
        vals = sorted({str(r[field]) for r in all_rows if r.get(field) is not None})
        return {
            val: {
                arm: _arm_rates([r for r in all_rows if r.get("arm") == arm and str(r.get(field)) == val])
                for arm in arms
            }
            for val in vals
        }

    return {
        "by_arm": by_arm,
        "by_taxonomy_group": _slice("taxonomy_group"),
        "by_length_tier": _slice("length_tier"),
        "by_instruction_type": _slice("instruction_type"),
        "by_expected_outcome": _slice("expected_outcome"),
    }
