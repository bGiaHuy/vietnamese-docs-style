# `test-sealed-v1` Official Single-Shot Benchmark Report

- **Status:** Completed (`evaluated_sealed_once`, single-shot official run)
- **Frozen System Tag:** `eval-freeze-v1`
- **Frozen System Commit SHA:** `e5ee327af209540a54d5e287a1b4e85cab502556`
- **Commitment Commit SHA:** `8fa6b9d` (`evals/SEALED-TEST-COMMITMENT.json`)
- **`SKILL.md` SHA-256:** `9103a7e50f0544883f3e582abc07d8141117e2a3475734aa88d9d0b3581ac6ff`
- **Model Configuration:** `ag_flash` (`gemini-4-hydrogen-tiered-high`)
- **Dataset Split:** `test-sealed-v1` (`30` cases × `3` arms = `90` isolated runs)
- **`bundle_sha256`:** `37415efd7149839204f0b54188ea336b3eb829805244542daf1dda5d95d67e63`
- **`per_case_manifest_sha256`:** `765faafb5b6950c00a4402bb931bd9327793a90eae8be2324bc750b7d4524c92`
- **`committed_at`:** `2026-11-26T20:59:02.054689+07:00`
- **`evaluated_at`:** `2026-11-26T23:11:57.909462+07:00`
- **Matched Transcripts / Contamination:** `90 / 90` matched, **`contamination_count = 0`**

---

## 1. Provenance, Air-Gap & Single-Shot Execution Disclosure

1. **Isolated Blind Curation:** The `test-sealed-v1` dataset (`30` cases across Groups `A`–`E`, quota `6/6/6/6/6`, `15` `descriptive` / `15` `inferential`, `24` `edit` incl. `3` `near_boundary_valid` / `6` `refuse`, length tiers `short: 8, medium: 9, long: 7, very_long: 6`) was curated outside this repository under `C:\Users\Administrator\Documents\Codespace\sealed-eval-v1\` by an isolated blind curator subagent that did not read `SKILL.md`, `scripts/`, `references/`, `assets/`, or `evals/val_v1/val_v1_cases.json`.
2. **Pre-Evaluation Cryptographic Gates:** Prior to any execution on `eval-freeze-v1`, the external bundle passed all curator gates (`lint-prompts`: `clean: true`, `validate-manifest`: `valid: true`, `check-disjoint`: `0` overlaps against the `20` `val-v1` source clusters, `verify-solvability`: `30/30` positive passes and `30/30` negative failures caught, and `verify-commitment`: `valid: true`). SHA-256 hashes were committed to Git (`8fa6b9d`) before unsealing.
3. **Single-Shot Execution & Locking:** All `90` runs executed against frozen tag `eval-freeze-v1` with zero edits to `SKILL.md`, `scripts/`, or `eval_harness/`. The initial `batch_size=6` wave completed `64/90` runs before hitting HTTP 429 (`RESOURCE_EXHAUSTED`); the remaining `26` unstarted runs completed `26/26` (`errors=0`) in one `batch_size=3` resume sweep before any transcript collection or grading occurred. After deterministic grading verified `90/90` transcripts and `contamination_count = 0`, `evals/SEALED-TEST-COMMITMENT.json` was locked via `mark-evaluated` (`evaluated_sealed_once`). Raw third-party `.docx` files, prompts, and grading assertions remain outside this repository.

---

## 2. Headline 3-Arm Results (`ag_flash`, `n = 30` cases / arm)

| Arm | Runs | `pass@1` | `pass^k` | Hit 95% CI | Checkpoint Mean | Collateral Rate | Over-Refusal | Unsafe-Edit | Mean Fresh+Out Tok | Mean Total Tok | Mean Calls | Mean Time (s) |
|---|---:|---:|---:|---|---:|---:|---:|---:|---:|---:|---:|---:|
| `arm0_no_skill` | 30 | 86.7% | 86.7% | [73.3%, 96.7%] | 97.5% | 0.0% | 0.0% | 0.0% | 92,842 | 359,183 | 12.4 | 100.7 |
| `arm1_standards_only` | 30 | 86.7% | 86.7% | [73.3%, 96.7%] | 97.8% | 4.2% | 0.0% | 0.0% | 100,918 | 405,180 | 17.0 | 105.5 |
| `arm2_full_skill` | 30 | **100.0%** | **100.0%** | **[100.0%, 100.0%]** | **100.0%** | **0.0%** | **0.0%** | **0.0%** | 169,574 | 1,046,337 | 29.4 | 165.9 |

---

## 3. Paired Statistical Comparisons (`n = 30` pairs)

| Comparison | Paired Runs | Hit Delta | Normalized Gain $g$ | Ceiling Saturated | Discordant Pairs (`b_a_only` / `c_b_only`) | McNemar Exact $p$ | Wilcoxon Checkpoint $p$ | Wilcoxon Total Tok $p$ | Wilcoxon Fresh+Out Tok $p$ |
|---|---:|---:|---:|---|---|---:|---:|---:|---:|
| `arm0_no_skill` vs `arm1_standards_only` | 30 | +0.00 pp | 0.0000 | `False` | 2 / 2 | 1.0000 | 1.0000 | 0.0703 | 0.0341 |
| `arm0_no_skill` vs `arm2_full_skill` | 30 | **+13.33 pp** | **1.0000** | `False` | 0 / 4 | 0.1250 | 0.1250 | $2.02 \times 10^{-6}$ | $2.24 \times 10^{-6}$ |
| `arm1_standards_only` vs `arm2_full_skill` | 30 | **+13.33 pp** | **1.0000** | `False` | 0 / 4 | 0.1250 | 0.1250 | $1.83 \times 10^{-6}$ | $2.02 \times 10^{-6}$ |

Key takeaways:
- **Zero-regression perfection on `arm2_full_skill`:** Every discordant pair between Arm 2 and both baselines favored Arm 2 (`0 / 4`), solving all 4 failures in Arm 0 and all 4 failures in Arm 1 (`g = 1.0000`) while maintaining `0.0%` non-target collateral rate, `0.0%` over-refusal, and `0.0%` unsafe-edit rate.
- **Standards prompt alone is insufficient:** Providing standards guidelines (`arm1_standards_only`) without OOXML structure mapping, XML-preserving edit helpers, and verification scripts produced the same overall `pass@1` (`86.7%`) as `arm0_no_skill`, and introduced a non-target semantic regression in `SEAL-C03` (`4.2%` collateral rate on edit tasks).
- **Failure modes eliminated by `arm2_full_skill`:**
  - `arm0_no_skill` (26/30 pass): 4 failures — `SEAL-B01` (`header_footer_byte_identical`), `SEAL-B02` (`header_footer_byte_identical`), `SEAL-C06` (`header_footer_byte_identical`), `SEAL-D06` (`ooxml_order_clean, header_footer_byte_identical`).
  - `arm1_standards_only` (26/30 pass): 4 failures — `SEAL-B01` (`header_footer_byte_identical`), `SEAL-B02` (`header_footer_byte_identical`), `SEAL-B03` (`header_footer_byte_identical`), `SEAL-C03` (`semantic_non_target_preserved`).
  - `arm2_full_skill` (30/30 pass): 0 failures across all 30 cases.
- **Token/latency trade-off:** `arm2_full_skill` consumed `1.83x` fresh+output tokens and `2.91x` total tokens relative to `arm0_no_skill` (65.2 s additional mean wall time), driven primarily by `scripts/validate_docx.py`, `scripts/render_docx.py`, and OOXML inspection workflows on complex tables/long documents.

---

## 4. Stratified Breakdowns

### 4.1 By Taxonomy Group (`A`–`E`, `n = 6` cases each)

| Group | Arm | `pass@1` | Checkpoint Mean | Mean Total Tok | Mean Calls |
|---|---|---:|---:|---:|---:|
| `A` | `arm0_no_skill` | 100.0% | 100.0% | 449,437 | 14.3 |
| `A` | `arm1_standards_only` | 100.0% | 100.0% | 451,824 | 18.7 |
| `A` | `arm2_full_skill` | 100.0% | 100.0% | 1,904,099 | 44.2 |
| `B` | `arm0_no_skill` | 66.7% | 94.6% | 370,019 | 13.0 |
| `B` | `arm1_standards_only` | 50.0% | 91.8% | 329,621 | 14.5 |
| `B` | `arm2_full_skill` | **100.0%** | **100.0%** | 630,087 | 22.5 |
| `C` | `arm0_no_skill` | 83.3% | 97.6% | 347,833 | 11.5 |
| `C` | `arm1_standards_only` | 83.3% | 97.2% | 493,161 | 20.7 |
| `C` | `arm2_full_skill` | **100.0%** | **100.0%** | 1,079,270 | 32.3 |
| `D` | `arm0_no_skill` | 83.3% | 95.2% | 290,954 | 9.8 |
| `D` | `arm1_standards_only` | 100.0% | 100.0% | 281,497 | 12.3 |
| `D` | `arm2_full_skill` | **100.0%** | **100.0%** | 599,783 | 19.5 |
| `E` | `arm0_no_skill` | 100.0% | 100.0% | 337,672 | 13.2 |
| `E` | `arm1_standards_only` | 100.0% | 100.0% | 469,794 | 18.8 |
| `E` | `arm2_full_skill` | 100.0% | 100.0% | 1,018,446 | 28.5 |

### 4.2 By Document Length Tier

| Length Tier | Arm | `pass@1` | Checkpoint Mean | Mean Total Tok | Mean Calls |
|---|---|---:|---:|---:|---:|
| `short` (`n = 8`) | `arm0_no_skill` | 87.5% | 97.5% | 291,729 | 10.8 |
| `short` (`n = 8`) | `arm1_standards_only` | 87.5% | 97.5% | 302,185 | 14.9 |
| `short` (`n = 8`) | `arm2_full_skill` | **100.0%** | **100.0%** | 816,325 | 24.6 |
| `medium` (`n = 9`) | `arm0_no_skill` | 88.9% | 98.6% | 369,614 | 13.4 |
| `medium` (`n = 9`) | `arm1_standards_only` | 77.8% | 96.8% | 336,442 | 14.9 |
| `medium` (`n = 9`) | `arm2_full_skill` | **100.0%** | **100.0%** | 814,114 | 25.3 |
| `long` (`n = 7`) | `arm0_no_skill` | 85.7% | 95.9% | 401,529 | 13.0 |
| `long` (`n = 7`) | `arm1_standards_only` | 85.7% | 97.6% | 493,396 | 17.7 |
| `long` (`n = 7`) | `arm2_full_skill` | **100.0%** | **100.0%** | 1,266,460 | 32.3 |
| `very_long` (`n = 6`) | `arm0_no_skill` | 83.3% | 97.6% | 384,073 | 12.2 |
| `very_long` (`n = 6`) | `arm1_standards_only` | 100.0% | 100.0% | 542,692 | 22.2 |
| `very_long` (`n = 6`) | `arm2_full_skill` | **100.0%** | **100.0%** | 1,444,543 | 38.5 |

### 4.3 By Instruction Type & Expected Outcome

| Slice | Arm | Runs | `pass@1` | Checkpoint Mean | Collateral Rate | Over-Refusal | Unsafe-Edit |
|---|---|---:|---:|---:|---:|---:|---:|
| `descriptive` | `arm0_no_skill` | 15 | 80.0% | 95.9% | — | — | — |
| `descriptive` | `arm1_standards_only` | 15 | 80.0% | 96.7% | — | — | — |
| `descriptive` | `arm2_full_skill` | 15 | **100.0%** | **100.0%** | — | — | — |
| `inferential` | `arm0_no_skill` | 15 | 93.3% | 99.1% | — | — | — |
| `inferential` | `arm1_standards_only` | 15 | 93.3% | 98.9% | — | — | — |
| `inferential` | `arm2_full_skill` | 15 | **100.0%** | **100.0%** | — | — | — |
| `edit` | `arm0_no_skill` | 24 | 83.3% | 96.9% | 0.0% | 0.0% | 0.0% |
| `edit` | `arm1_standards_only` | 24 | 83.3% | 97.3% | 4.2% | 0.0% | 0.0% |
| `edit` | `arm2_full_skill` | 24 | **100.0%** | **100.0%** | **0.0%** | **0.0%** | **0.0%** |
| `refuse` | `arm0_no_skill` | 6 | 100.0% | 100.0% | 0.0% | 0.0% | 0.0% |
| `refuse` | `arm1_standards_only` | 6 | 100.0% | 100.0% | 0.0% | 0.0% | 0.0% |
| `refuse` | `arm2_full_skill` | 6 | **100.0%** | **100.0%** | **0.0%** | **0.0%** | **0.0%** |

---

## 5. Blind Human Audit Pack

A blinded side-by-side human inspection bundle (`12` sampled cases × `3` randomized outputs = `36` items, `seed=42`) was exported locally to the gitignored directory `research/eval/test_sealed_v1/blind_audit/`:
- Visual viewer: `research/eval/test_sealed_v1/blind_audit/viewer.html`
- Scoring sheet: `research/eval/test_sealed_v1/blind_audit/scoring_sheet.csv`
- Leakage check: `scoring_sheet.csv` and `viewer.html` verified clean (`0` matches for arm names, skill identifiers, or model tags).
- Human scoring status: pending optional human review (`score_*` columns empty in initial export).
