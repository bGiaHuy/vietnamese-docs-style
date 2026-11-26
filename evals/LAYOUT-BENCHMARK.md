# Layout Editing Benchmark Report (`evals/layout_edit_cases.json` v2.0)

> [!IMPORTANT]
> **Frozen Historical Development Baseline (`dev-v2.0`):** Benchmark này được đồng thiết kế trong quá trình phát triển `vndocs` (`27` tài liệu sinh từ builder nội bộ + `3` starter templates; oracle/grader dùng chung code với `scripts/`; `50` Live A/B trajectories trên `T01`–`T18` đạt `100%` hit rate ở cả hai arm / bão hòa trần, $g = 0.0\%$). Toàn bộ dữ liệu và số liệu lịch sử bên dưới được **giữ nguyên không chỉnh sửa** để làm baseline lịch sử và hồi quy kỹ thuật, **không** dùng làm bằng chứng tổng quát hóa trên tài liệu thực tế. Xem manifest đóng băng tại [snapshots/dev-v2.0-manifest.json](snapshots/dev-v2.0-manifest.json) và giao thức đánh giá độc lập tại [EVALUATION-PROTOCOL.md](EVALUATION-PROTOCOL.md).

> **Scope & Methodology Note:** Self-authored benchmark designed for Vietnamese `.docx` workflows (Decree 30/2020/NĐ-CP layout protection, Party HD 05-HD/VPTW formatting, NFD Vietnamese, Word 365 namespaces, OOXML digital signatures). Cases `T01`–`T18` were evaluated in **50 isolated Live Agent A/B trajectories** (`without_skill` vs `with_skill`, `0` Arm-A contamination) alongside the deterministic oracle; extended cases `T19`–`T30` are verified by the **30-case deterministic oracle** (`100.0%` hit, `xml_ntpr = 1.0`) and have not yet been evaluated in Live Agent A/B runs. All local machine paths and session IDs are stripped from [evals/layout-benchmark-results.json](layout-benchmark-results.json).

- Generated: `2026-11-22T21:26:23.682590+07:00`
- Corpus cases: **30** (`short`: 14, `medium`: 10, `long`: 6)
- Deterministic Oracle Hit Rate: **30/30 (100.0%)**
- Mean Checkpoint Score: **100.0%**
- Mean `xml_ntpr`: **1.0000** | Mean ZIP non-target byte identity: **0.9960**

---

## 1. Live Agent A/B Comparisons (`without_skill` vs `with_skill`, `T01`–`T18`, 50 Trajectories)

### Batch 1 (`T01`–`T12`: Tables, Merged Grid/VMerge, Multi-Section Landscape, Ripple, Refusals — 32 Runs)

| Model | Cases (each arm) | Arm | Hit Rate | Sem. NTP | `xml_ntpr` | Mean Fresh+Out Tok | Mean Total Tok | Mean Calls | Mean Time (s) | Token & Time Delta |
|---|---:|---|---:|---:|---:|---:|---:|---:|---:|---|
| `flash` | 12 (`T01`–`T12`) | `without_skill` | 12/12 (100%) | 100.0% | 1.000 | 47,178 | 219,273 | 11.3 | 98.1 | — |
| `flash` | 12 (`T01`–`T12`) | `with_skill` | **12/12 (100%)** | **100.0%** | **1.000** | **27,158** | **72,536** | **6.3** | **41.7** | **-66.9% total tok · -42.4% fresh+out · -57.5% time** |
| `pro` | 4 (`T03, T04, T07, T08`) | `without_skill` | 4/4 (100%) | 100.0% | 1.000 | 96,552 | 619,055 | 19.0 | 159.0 | — |
| `pro` | 4 (`T03, T04, T07, T08`) | `with_skill` | **4/4 (100%)** | **100.0%** | **1.000** | **43,885** | **155,376** | **9.5** | **59.5** | **-74.9% total tok · -54.5% fresh+out · -62.6% time** |

### Batch 2 (`T13`–`T18`: Cropped Screenshots, Cell Margins/Fonts, Image Scale, Party HD05, Multi-Op — 18 Runs)

| Model | Cases (each arm) | Arm | Hit Rate | Sem. NTP | `xml_ntpr` | Mean Fresh+Out Tok | Mean Total Tok | Mean Calls | Mean Time (s) | Token & Time Delta |
|---|---:|---|---:|---:|---:|---:|---:|---:|---:|---|
| `flash` | 6 (`T13`–`T18`) | `without_skill` | 6/6 (100%) | 100.0% | 1.000 | 42,844 | 208,550 | 11.8 | 77.2 | — |
| `flash` | 6 (`T13`–`T18`) | `with_skill` | **6/6 (100%)** | **100.0%** | **1.000** | **29,739** | **97,790** | **7.0** | **41.3** | **-53.1% total tok · -30.6% fresh+out · -46.5% time** |
| `pro` | 3 (`T13, T16, T17`) | `without_skill` | 3/3 (100%) | 100.0% | 1.000 | 58,507 | 324,405 | 17.7 | 110.7 | — |
| `pro` | 3 (`T13, T16, T17`) | `with_skill` | **3/3 (100%)** | **100.0%** | **1.000** | **48,202** | **237,946** | **13.3** | **85.0** | **-26.7% total tok · -17.6% fresh+out · -23.2% time** |

> **Honest Observation on Screenshot Crop Grounding (`T13`, `T17`):** In Batch 2, running `render --annotate-marks` unconditionally on `T13`/`T17` added extra rendered image tokens on medium-length `flash` runs (`with_skill` `156,181` vs `without_skill` `135,831` total tokens on `T13`, and `102,964` vs `65,310` on `T17`), even while `pro` stayed lower across all 3 cases (`378,129` vs `393,209` on `T13`; `146,079` vs `185,763` on `T17`). Based on this measurement, [SKILL.md](../SKILL.md) was updated to match unique screenshot text/structure against `vndocs map` first and only run `--annotate-marks` when `>= 2` candidates remain ambiguous.

<details>
<summary>All 50 Live Agent A/B Trajectories (exact metrics from sanitized logs)</summary>

#### Batch 1 (`T01`–`T12`, 32 runs)

| Model | Case | Tier | Type | Arm | Hit | CP | Invoked | Total Tok | Fresh+Out | Calls | Time (s) |
|---|---|---|---|---|---:|---:|---:|---:|---:|---:|---:|
| `flash` | `T01_starter_scale_center` | short | descriptive | `without_skill` | 1 | 100% | N | 234563 | 51105 | 12 | 120.0 |
| `flash` | `T01_starter_scale_center` | short | descriptive | `with_skill` | 1 | 100% | Y | 74527 | 28173 | 7 | 56.0 |
| `flash` | `T02_starter_minutes_cols` | short | descriptive | `without_skill` | 1 | 100% | N | 180946 | 39625 | 8 | 109.0 |
| `flash` | `T02_starter_minutes_cols` | short | descriptive | `with_skill` | 1 | 100% | Y | 62549 | 25013 | 6 | 36.0 |
| `flash` | `T03_nd30_budget_not_header` | short | inferential | `without_skill` | 1 | 100% | N | 213561 | 49667 | 13 | 95.0 |
| `flash` | `T03_nd30_budget_not_header` | short | inferential | `with_skill` | 1 | 100% | Y | 100899 | 33548 | 9 | 56.0 |
| `flash` | `T04_nd30_layout_only_refusal` | short | inferential | `without_skill` | 1 | 100% | N | 57455 | 23425 | 5 | 29.0 |
| `flash` | `T04_nd30_layout_only_refusal` | short | inferential | `with_skill` | 1 | 100% | Y | 32777 | 16166 | 3 | 15.0 |
| `flash` | `T05_long_multi_similar_tables` | long | inferential | `without_skill` | 1 | 100% | N | 396176 | 70542 | 16 | 136.0 |
| `flash` | `T05_long_multi_similar_tables` | long | inferential | `with_skill` | 1 | 100% | Y | 69433 | 27402 | 6 | 33.0 |
| `flash` | `T06_overflow_wide_table` | short | inferential | `without_skill` | 1 | 100% | N | 135692 | 34088 | 8 | 63.0 |
| `flash` | `T06_overflow_wide_table` | short | inferential | `with_skill` | 1 | 100% | Y | 83545 | 30818 | 6 | 63.0 |
| `flash` | `T07_merged_gridspan_vmerge_med` | medium | descriptive | `without_skill` | 1 | 100% | N | 242225 | 51028 | 11 | 95.0 |
| `flash` | `T07_merged_gridspan_vmerge_med` | medium | descriptive | `with_skill` | 1 | 100% | Y | 66974 | 26213 | 6 | 32.0 |
| `flash` | `T08_long_landscape_multisection` | long | inferential | `without_skill` | 1 | 100% | N | 174075 | 39762 | 8 | 70.0 |
| `flash` | `T08_long_landscape_multisection` | long | inferential | `with_skill` | 1 | 100% | Y | 91749 | 31927 | 8 | 45.0 |
| `flash` | `T09_long_multipage_table` | long | descriptive | `without_skill` | 1 | 100% | N | 255991 | 53290 | 15 | 109.0 |
| `flash` | `T09_long_multipage_table` | long | descriptive | `with_skill` | 1 | 100% | Y | 76206 | 28041 | 7 | 41.0 |
| `flash` | `T10_nfd_vietnamese_med` | medium | descriptive | `without_skill` | 1 | 100% | N | 427484 | 67872 | 20 | 179.0 |
| `flash` | `T10_nfd_vietnamese_med` | medium | descriptive | `with_skill` | 1 | 100% | Y | 66110 | 26286 | 6 | 33.0 |
| `flash` | `T11_impossible_narrow_col_refusal` | medium | descriptive | `without_skill` | 1 | 100% | N | 74678 | 28590 | 7 | 40.0 |
| `flash` | `T11_impossible_narrow_col_refusal` | medium | descriptive | `with_skill` | 1 | 100% | Y | 32702 | 16159 | 3 | 17.0 |
| `flash` | `T12_pagination_ripple_awareness` | short | descriptive | `without_skill` | 1 | 100% | N | 238433 | 57142 | 13 | 132.0 |
| `flash` | `T12_pagination_ripple_awareness` | short | descriptive | `with_skill` | 1 | 100% | Y | 112957 | 36156 | 9 | 73.0 |
| `pro` | `T03_nd30_budget_not_header` | short | inferential | `without_skill` | 1 | 100% | N | 731200 | 95969 | 23 | 202.0 |
| `pro` | `T03_nd30_budget_not_header` | short | inferential | `with_skill` | 1 | 100% | Y | 121685 | 48457 | 10 | 65.0 |
| `pro` | `T04_nd30_layout_only_refusal` | short | inferential | `without_skill` | 1 | 100% | N | 124716 | 46760 | 9 | 63.0 |
| `pro` | `T04_nd30_layout_only_refusal` | short | inferential | `with_skill` | 1 | 100% | Y | 68015 | 25639 | 6 | 32.0 |
| `pro` | `T07_merged_gridspan_vmerge_med` | medium | descriptive | `without_skill` | 1 | 100% | N | 886851 | 99026 | 22 | 194.0 |
| `pro` | `T07_merged_gridspan_vmerge_med` | medium | descriptive | `with_skill` | 1 | 100% | Y | 182102 | 47170 | 10 | 66.0 |
| `pro` | `T08_long_landscape_multisection` | long | inferential | `without_skill` | 1 | 100% | N | 733453 | 144452 | 22 | 177.0 |
| `pro` | `T08_long_landscape_multisection` | long | inferential | `with_skill` | 1 | 100% | Y | 249700 | 54273 | 12 | 75.0 |

#### Batch 2 (`T13`–`T18`, 18 runs)

| Model | Case | Tier | Type | Arm | Hit | CP | Invoked | Total Tok | Fresh+Out | Calls | Time (s) |
|---|---|---|---|---|---:|---:|---:|---:|---:|---:|---:|
| `flash` | `T13_screenshot_cropped_similar_tables` | medium | inferential | `without_skill` | 1 | 100% | N | 135831 | 36300 | 12 | 60.0 |
| `flash` | `T13_screenshot_cropped_similar_tables` | medium | inferential | `with_skill` | 1 | 100% | Y | 156181 | 38881 | 10 | 78.0 |
| `flash` | `T14_table_cell_margins_and_font_nd30` | short | descriptive | `without_skill` | 1 | 100% | N | 183953 | 41126 | 9 | 70.0 |
| `flash` | `T14_table_cell_margins_and_font_nd30` | short | descriptive | `with_skill` | 1 | 100% | Y | 57726 | 22892 | 5 | 27.0 |
| `flash` | `T15_image_scale_aspect_ratio` | short | inferential | `without_skill` | 1 | 100% | N | 131973 | 32695 | 10 | 72.0 |
| `flash` | `T15_image_scale_aspect_ratio` | short | inferential | `with_skill` | 1 | 100% | Y | 56071 | 23016 | 5 | 24.0 |
| `flash` | `T16_party_hd05_paragraph_format` | short | inferential | `without_skill` | 1 | 100% | N | 240354 | 49780 | 17 | 94.0 |
| `flash` | `T16_party_hd05_paragraph_format` | short | inferential | `with_skill` | 1 | 100% | Y | 70502 | 26221 | 6 | 33.0 |
| `flash` | `T17_screenshot_ambiguous_identical_tables_refusal` | medium | inferential | `without_skill` | 1 | 100% | N | 65310 | 21587 | 6 | 31.0 |
| `flash` | `T17_screenshot_ambiguous_identical_tables_refusal` | medium | inferential | `with_skill` | 1 | 100% | Y | 102964 | 30518 | 8 | 43.0 |
| `flash` | `T18_combined_multi_op_long_doc` | long | descriptive | `without_skill` | 1 | 100% | N | 493876 | 75573 | 17 | 136.0 |
| `flash` | `T18_combined_multi_op_long_doc` | long | descriptive | `with_skill` | 1 | 100% | Y | 143299 | 36906 | 8 | 43.0 |
| `pro` | `T13_screenshot_cropped_similar_tables` | medium | inferential | `without_skill` | 1 | 100% | N | 393209 | 65611 | 20 | 120.0 |
| `pro` | `T13_screenshot_cropped_similar_tables` | medium | inferential | `with_skill` | 1 | 100% | Y | 378129 | 60295 | 18 | 113.0 |
| `pro` | `T16_party_hd05_paragraph_format` | short | inferential | `without_skill` | 1 | 100% | N | 394242 | 66905 | 20 | 147.0 |
| `pro` | `T16_party_hd05_paragraph_format` | short | inferential | `with_skill` | 1 | 100% | Y | 189630 | 47393 | 11 | 83.0 |
| `pro` | `T17_screenshot_ambiguous_identical_tables_refusal` | medium | inferential | `without_skill` | 1 | 100% | N | 185763 | 43005 | 13 | 65.0 |
| `pro` | `T17_screenshot_ambiguous_identical_tables_refusal` | medium | inferential | `with_skill` | 1 | 100% | Y | 146079 | 36919 | 11 | 59.0 |

</details>

---

## 2. Deterministic 30-Case Oracle & Static Context Payload (`T01`–`T30`)

### Stratification Summary

| Slice | Cases | Hits | Hit Rate | Mean Checkpoint | Mean `xml_ntpr` |
|---|---:|---:|---:|---:|---:|
| `length:short` | 14 | 14 | 100.0% | 100.0% | 1.0000 |
| `length:medium` | 10 | 10 | 100.0% | 100.0% | 1.0000 |
| `length:long` | 6 | 6 | 100.0% | 100.0% | 1.0000 |
| `type:descriptive` | 16 | 16 | 100.0% | 100.0% | 1.0000 |
| `type:inferential` | 14 | 14 | 100.0% | 100.0% | 1.0000 |
| `outcome:edit` | 26 | 26 | 100.0% | 100.0% | 1.0000 |
| `outcome:refuse` | 4 | 4 | 100.0% | 100.0% | 1.0000 |

### Static Context Payload (`word/document.xml` vs `vndocs map`)

- `SKILL.md`: **9484 bytes** (~2371 est tok)
- `SKILL.md` + `references/tooling.md`: **24382 bytes** (~6096 est tok)

| Tier | Cases | Mean `document.xml` | Mean `map --format markdown` | Reduction |
|---|---:|---:|---:|---:|
| `short` | 13 | 5167 B (~1292 tok) | 513 B (~128 tok) | **90.1%** |
| `medium` | 10 | 16025 B (~4006 tok) | 503 B (~126 tok) | **96.9%** |
| `long` | 6 | 38742 B (~9686 tok) | 763 B (~191 tok) | **98.0%** |

### Per-Case Oracle Results (`T01`–`T30`)

| Case ID | Tier | Type | Outcome | Hit | CP | `xml_ntpr` | Failures |
|---|---|---|---|---:|---:|---:|---|
| `T01_starter_scale_center` | short | descriptive | edit | 1 | 100% | 1.0000 | - |
| `T02_starter_minutes_cols` | short | descriptive | edit | 1 | 100% | 1.0000 | - |
| `T03_nd30_budget_not_header` | short | inferential | edit | 1 | 100% | 1.0000 | - |
| `T04_nd30_layout_only_refusal` | short | inferential | refuse | 1 | 100% | 1.0000 | - |
| `T05_long_multi_similar_tables` | long | inferential | edit | 1 | 100% | 1.0000 | - |
| `T06_overflow_wide_table` | short | inferential | edit | 1 | 100% | 1.0000 | - |
| `T07_merged_gridspan_vmerge_med` | medium | descriptive | edit | 1 | 100% | 1.0000 | - |
| `T08_long_landscape_multisection` | long | inferential | edit | 1 | 100% | 1.0000 | - |
| `T09_long_multipage_table` | long | descriptive | edit | 1 | 100% | 1.0000 | - |
| `T10_nfd_vietnamese_med` | medium | descriptive | edit | 1 | 100% | 1.0000 | - |
| `T11_impossible_narrow_col_refusal` | medium | descriptive | refuse | 1 | 100% | 1.0000 | - |
| `T12_pagination_ripple_awareness` | short | descriptive | edit | 1 | 100% | 1.0000 | - |
| `T13_screenshot_cropped_similar_tables` | medium | inferential | edit | 1 | 100% | 1.0000 | - |
| `T14_table_cell_margins_and_font_nd30` | short | descriptive | edit | 1 | 100% | 1.0000 | - |
| `T15_image_scale_aspect_ratio` | short | inferential | edit | 1 | 100% | 1.0000 | - |
| `T16_party_hd05_paragraph_format` | short | inferential | edit | 1 | 100% | 1.0000 | - |
| `T17_screenshot_ambiguous_identical_tables_refusal` | medium | inferential | refuse | 1 | 100% | 1.0000 | - |
| `T18_combined_multi_op_long_doc` | long | descriptive | edit | 1 | 100% | 1.0000 | - |
| `T19_table_pagination_repeat_header_cantsplit` | long | descriptive | edit | 1 | 100% | 1.0000 | - |
| `T20_table_borders_and_header_shading` | short | descriptive | edit | 1 | 100% | 1.0000 | - |
| `T21_style_update_normal_global` | medium | descriptive | edit | 1 | 100% | 1.0000 | - |
| `T22_section_landscape_appendix_layout` | medium | descriptive | edit | 1 | 100% | 1.0000 | - |
| `T23_paragraph_merge_fragmented_runs` | short | descriptive | edit | 1 | 100% | 1.0000 | - |
| `T24_track_changes_table_scale` | short | descriptive | edit | 1 | 100% | 1.0000 | - |
| `T25_auto_paginate_orphan_nd30_signature` | short | inferential | edit | 1 | 100% | 1.0000 | - |
| `T26_auto_paginate_target_2pages_report` | medium | inferential | edit | 1 | 100% | 1.0000 | - |
| `T27_complex_word365_namespaces_mc_alternate` | medium | descriptive | edit | 1 | 100% | 1.0000 | - |
| `T28_signed_docx_refusal` | short | inferential | refuse | 1 | 100% | 1.0000 | - |
| `T29_nested_header_footer_preservation` | medium | inferential | edit | 1 | 100% | 1.0000 | - |
| `T30_long_full_pipeline_nd30_report` | long | inferential | edit | 1 | 100% | 1.0000 | - |

---

## 3. Public Dataset & Reproduction

- **Dataset specification (30 cases & grading schema):** [`evals/layout_edit_cases.json`](layout_edit_cases.json)
- **Sanitized raw results (30 oracle cases + 50 Live Agent A/B trajectories):** [`evals/layout-benchmark-results.json`](layout-benchmark-results.json)
- **Reproduce locally (generates all 30 `.docx` sources + screenshot PNG crops without storing binary build artifacts in Git):**

```bash
python evals/run_layout_benchmark.py --generate-corpus artifacts/layout-corpus
python evals/run_layout_benchmark.py --measure-payload --corpus-dir artifacts/layout-corpus
python evals/run_layout_benchmark.py --run-oracle --write-doc evals/LAYOUT-BENCHMARK.md
python evals/run_layout_benchmark.py --grade-dir artifacts/submissions --corpus-dir artifacts/layout-corpus
```

---

## 4. Design & Evaluation References

The architecture (`map` -> `layout-edit` / `paginate` -> `diff-layout`), XML safety guardrails, and benchmark methodology draw on:

| Source | Venue / Repo | Key Takeaway Applied in `vndocs` |
|---|---|---|
| **WordInferBench** | [arXiv:2611.15189](https://arxiv.org/abs/2611.15189) | Table column width is #1 failure point (`36.0%` agent success); measures Descriptive vs Inferential instructions and non-target preservation. |
| **DocFormBench & DocFormFlow** | [arXiv:2606.01936](https://arxiv.org/abs/2606.01936) | Splitting Locate from Execute lifts Format Accuracy (`<26% -> 65.8%–80.4%`); stratified by Short/Medium/Long; tracks side effects (HFR). |
| **SkillsBench** | [arXiv:2602.12670](https://arxiv.org/abs/2602.12670) | Paired `without_skill` vs `with_skill` A/B design, zero task/skill leakage, Normalized Gain $g$, compact skill instructions + Complexity Contract. |
| **OfficeEval** | [arXiv:2606.10956](https://arxiv.org/abs/2606.10956) | Multi-checkpoint partial credit alongside binary hit; validates OOXML container path, enum, and unit correctness. |
| **LEDGER** | [arXiv:2606.28379](https://arxiv.org/abs/2606.28379) | Focused structural retrieval cuts context tokens `85%–92%` (`vndocs map` cuts `90.1%–98.0%`) and reduces long-document drift. |
| **DocEdit-v2** | [arXiv:2410.16472](https://arxiv.org/abs/2410.16472) | Set-of-Mark bounding-box overlays (`--annotate-marks`) for disambiguating target regions in cropped screenshots. |
| **VFLM & AutoPresent** | [arXiv:2603.22187](https://arxiv.org/abs/2603.22187) · [arXiv:2501.00912](https://arxiv.org/abs/2501.00912) | Unspecialized VLM self-checking degrades edits without objective metrics; visual loops saturate after 2–3 iterations (`max_iter = 2`). |
| **`kklimuk/docx-cli`** | [GitHub (MIT)](https://github.com/kklimuk/docx-cli) | Synced `tblGrid` + `tcW` column edits, `< 0.51 cm` text-wrap guardrail, ECMA-376 child element order validator, and Windows Word COM fallback. |
| **`ruwadgroup/docxengine`** | [GitHub (Apache-2.0)](https://github.com/ruwadgroup/docxengine) | Content-hashed NFC anchors (`T#h4`, `IMG#h4`, `P#h4`), stale-anchor protection (`--source-sha256`), and Markdown map projection. |
| **`anthropics/skills`** | [GitHub (Apache-2.0)](https://github.com/anthropics/skills) | OOXML container sanity checks, non-directory ZIP writing, pure-text `merge_runs`, and skill benchmark JSON reporting structure. |
| **`icip-cas/PPTAgent`** | [GitHub (MIT)](https://github.com/icip-cas/PPTAgent) · [arXiv:2501.03936](https://arxiv.org/abs/2501.03936) | Typed Edit APIs outperform free-form XML edits; SHA-256 snapshot freshness guard (`--cache-dir` manifest). |
