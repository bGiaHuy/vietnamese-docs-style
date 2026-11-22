# Layout Editing Benchmark Report (`evals/layout_edit_cases.json` v2.0)

- Generated: `2026-11-22T21:26:23.682590+07:00`
- Corpus cases: **30** (`short`: 14, `medium`: 10, `long`: 6)
- Deterministic Oracle Hit Rate: **30/30 (100.0%)**
- Mean Checkpoint Score: **100.0%**
- Mean `xml_ntpr`: **1.0000** | Mean ZIP non-target byte identity: **0.9960**

## 1. Stratification Summary

| Slice | Cases | Hits | Hit Rate | Mean Checkpoint | Mean `xml_ntpr` |
|---|---:|---:|---:|---:|---:|
| `length:short` | 14 | 14 | 100.0% | 100.0% | 1.0000 |
| `length:medium` | 10 | 10 | 100.0% | 100.0% | 1.0000 |
| `length:long` | 6 | 6 | 100.0% | 100.0% | 1.0000 |
| `type:descriptive` | 16 | 16 | 100.0% | 100.0% | 1.0000 |
| `type:inferential` | 14 | 14 | 100.0% | 100.0% | 1.0000 |
| `outcome:edit` | 26 | 26 | 100.0% | 100.0% | 1.0000 |
| `outcome:refuse` | 4 | 4 | 100.0% | 100.0% | 1.0000 |

## 2. Static Context Payload (`word/document.xml` vs `vndocs map`)

- `SKILL.md`: **9484 bytes** (~2371 est tok)
- `SKILL.md` + `references/tooling.md`: **24382 bytes** (~6096 est tok)

| Tier | Cases | Mean `document.xml` | Mean `map --format markdown` | Reduction |
|---|---:|---:|---:|---:|
| `short` | 13 | 5167 B (~1292 tok) | 513 B (~128 tok) | **90.1%** |
| `medium` | 10 | 16025 B (~4006 tok) | 503 B (~126 tok) | **96.9%** |
| `long` | 6 | 38742 B (~9686 tok) | 763 B (~191 tok) | **98.0%** |

## 3. Per-Case Oracle Results (`T01`–`T30`)

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

## 4. Reproduce

```bash
python evals/run_layout_benchmark.py --run-oracle --write-doc evals/LAYOUT-BENCHMARK.md
python evals/run_layout_benchmark.py --measure-payload
```
