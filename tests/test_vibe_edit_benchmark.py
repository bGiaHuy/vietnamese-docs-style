"""Unit and oracle grader tests for the Vibe Docs benchmark suite (`evals/`)."""
import importlib.util
import json
from pathlib import Path

import pytest

import layout_edit
import map_docx
from render_docx import find_renderer

ROOT = Path(__file__).resolve().parent.parent
CASES_PATH = ROOT / 'evals' / 'vibe_edit_cases.json'
BENCH_SCRIPT = ROOT / 'evals' / 'run_vibe_edit_benchmark.py'


def _load_bench():
    spec = importlib.util.spec_from_file_location('run_vibe_edit_benchmark', BENCH_SCRIPT)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def test_cases_schema_stratification_and_zero_leakage():
    data = json.loads(CASES_PATH.read_text(encoding='utf-8'))
    cases = data['cases']
    assert len(cases) == 30

    tiers = {}
    outcomes = {}
    inherit_spots = []
    forbidden = ('vndocs', 'layout-edit', 'layout_edit', 'map_docx', 'diff-layout', 'diff_layout')
    for c in cases:
        tiers[c['length_tier']] = tiers.get(c['length_tier'], 0) + 1
        outcomes[c['expected_outcome']] = outcomes.get(c['expected_outcome'], 0) + 1
        if c.get('spot_check_inherit'):
            inherit_spots.append(c['id'])
        low_prompt = c['prompt_vi'].lower()
        for kw in forbidden:
            assert kw not in low_prompt, f"Leakage keyword {kw!r} in {c['id']}"

    assert tiers == {'short': 14, 'medium': 10, 'long': 6}
    assert outcomes == {'edit': 26, 'refuse': 4}
    assert set(inherit_spots) == {
        'T03_nd30_budget_not_header',
        'T04_nd30_layout_only_refusal',
        'T07_merged_gridspan_vmerge_med',
        'T08_long_landscape_multisection',
        'T13_screenshot_cropped_similar_tables',
        'T16_party_hd05_paragraph_format',
        'T17_screenshot_ambiguous_identical_tables_refusal',
        'T25_auto_paginate_orphan_nd30_signature',
        'T27_complex_word365_namespaces_mc_alternate',
    }


def test_generate_corpus_and_measure_payload(tmp_path):
    bench = _load_bench()
    corpus_dir = tmp_path / 'corpus'
    manifest = bench.generate_corpus(corpus_dir)
    assert len(manifest) == 30
    for row in manifest:
        assert Path(row['path']).is_file()
        assert row['table_count'] >= 1 or row['image_count'] >= 1 or row['paragraph_count'] >= 1
        if find_renderer() and row['id'] in (
            'T13_screenshot_cropped_similar_tables',
            'T17_screenshot_ambiguous_identical_tables_refusal',
        ):
            assert row['screenshot_path'] and Path(row['screenshot_path']).is_file()

    payload = bench.measure_payload(corpus_dir)
    assert payload['by_length_tier']['long']['mean_raw_doc_xml_bytes'] > payload['by_length_tier']['short']['mean_raw_doc_xml_bytes']
    assert payload['by_length_tier']['long']['mean_reduction_pct'] > 85.0
    assert payload['instruction_overhead']['compact_vibe_skill_bytes'] < payload['instruction_overhead']['full_skill_md_bytes']


def test_grader_refusal_t04_and_transcript_parser(tmp_path):
    bench = _load_bench()
    corpus_dir = tmp_path / 'corpus'
    bench.generate_corpus(corpus_dir)
    cases = {c['id']: c for c in json.loads(CASES_PATH.read_text(encoding='utf-8'))['cases']}

    t04 = cases['T04_nd30_layout_only_refusal']
    src04 = corpus_dir / 'T04_nd30_layout_only_refusal.docx'

    # Clean refusal -> Hit = 1
    good_dir = tmp_path / 'sub_good' / t04['id']
    good_dir.mkdir(parents=True)
    (good_dir / f"{t04['id']}.result.json").write_text(
        json.dumps({'status': 'refused', 'reason': 'Chỉ có bảng bố cục NĐ30'}, ensure_ascii=False),
        encoding='utf-8',
    )
    res_good = bench.grade_case(t04, src04, good_dir)
    assert res_good['hit'] == 1
    assert res_good['checkpoint_score'] == 1.0

    # Illegal edit of layout table -> Hit = 0
    bad_dir = tmp_path / 'sub_bad' / t04['id']
    bad_dir.mkdir(parents=True)
    (bad_dir / f"{t04['id']}.result.json").write_text(
        json.dumps({'status': 'edited', 'reason': 'shrunk'}, ensure_ascii=False),
        encoding='utf-8',
    )
    m04 = map_docx.map_document(src04)
    layout_edit.apply_layout_plan(
        src04,
        bad_dir / f"{t04['id']}.docx",
        [{'op': 'table.scale_width', 'target': m04['tables'][0]['anchor'], 'factor': 0.85}],
        allow_layout_tables=True,
    )
    res_bad = bench.grade_case(t04, src04, bad_dir)
    assert res_bad['hit'] == 0

    # Transcript parser handles fresh/cache/output tokens + contamination marker
    tr_file = tmp_path / 'transcript.jsonl'
    tr_lines = [
        {'step_index': 0, 'type': 'USER_INPUT', 'created_at': '2026-11-22T10:00:00Z', 'content': 'hi'},
        {
            'step_index': 1,
            'type': 'PLANNER_RESPONSE',
            'created_at': '2026-11-22T10:00:05Z',
            'model': 'gemini-3-flash',
            'input_tokens': 1000,
            'cache_read_tokens': 0,
            'output_tokens': 200,
            'tool_calls': [{'name': 'run_command', 'args': {'CommandLine': 'vndocs map src.docx'}}],
        },
        {
            'step_index': 3,
            'type': 'PLANNER_RESPONSE',
            'created_at': '2026-11-22T10:00:10Z',
            'model': 'gemini-3-flash',
            'input_tokens': 300,
            'cache_read_tokens': 900,
            'output_tokens': 150,
            'tool_calls': [],
        },
    ]
    tr_file.write_text('\n'.join(json.dumps(x) for x in tr_lines), encoding='utf-8')
    parsed = bench.parse_transcript(tr_file)
    assert parsed['fresh_input_tokens'] == 1300
    assert parsed['cache_read_tokens'] == 900
    assert parsed['output_tokens'] == 350
    assert parsed['fresh_plus_output_tokens'] == 1650
    assert parsed['total_tokens'] == 2550
    assert parsed['elapsed_seconds'] == 10.0
    assert parsed['saw_vndocs_marker'] is True


@pytest.mark.skipif(not find_renderer(), reason='LibreOffice renderer not installed')
def test_grader_oracle_merged_and_ripple(tmp_path):
    bench = _load_bench()
    corpus_dir = tmp_path / 'corpus'
    bench.generate_corpus(corpus_dir)
    cases = {c['id']: c for c in json.loads(CASES_PATH.read_text(encoding='utf-8'))['cases']}

    # T07 oracle with vndocs layout-edit -> Hit = 1
    t07 = cases['T07_merged_gridspan_vmerge_med']
    src07 = corpus_dir / 'T07_merged_gridspan_vmerge_med.docx'
    m07 = map_docx.map_document(src07)
    anchor07 = m07['tables'][1]['anchor']
    sub07 = tmp_path / 'sub' / t07['id']
    sub07.mkdir(parents=True)
    layout_edit.apply_layout_plan(
        src07,
        sub07 / f"{t07['id']}.docx",
        [
            {'op': 'table.set_col_widths', 'target': anchor07, 'widths': ['2.5cm', '6.0cm', '2.5cm', '3.0cm']},
            {'op': 'table.set_alignment', 'target': anchor07, 'align': 'center'},
        ],
    )
    (sub07 / f"{t07['id']}.result.json").write_text(
        json.dumps({'status': 'edited', 'reason': 'ok', 'pagination_ripple': None}),
        encoding='utf-8',
    )
    res07 = bench.grade_case(t07, src07, sub07)
    assert res07['hit'] == 1, res07
    assert res07['telemetry']['xml_ntpr'] == 1.0

    # T12 oracle with ripple=True -> Hit = 1
    t12 = cases['T12_pagination_ripple_awareness']
    src12 = corpus_dir / 'T12_pagination_ripple_awareness.docx'
    m12 = map_docx.map_document(src12)
    anchor12 = m12['tables'][0]['anchor']
    sub12 = tmp_path / 'sub' / t12['id']
    sub12.mkdir(parents=True)
    layout_edit.apply_layout_plan(
        src12,
        sub12 / f"{t12['id']}.docx",
        [
            {'op': 'table.scale_width', 'target': anchor12, 'factor': 0.55},
            {'op': 'table.set_alignment', 'target': anchor12, 'align': 'center'},
        ],
    )
    (sub12 / f"{t12['id']}.result.json").write_text(
        json.dumps({'status': 'edited', 'reason': 'scaled 55% + centered', 'pagination_ripple': True}),
        encoding='utf-8',
    )
    res12 = bench.grade_case(t12, src12, sub12)
    assert res12['hit'] == 1, res12


@pytest.mark.skipif(not find_renderer(), reason='LibreOffice renderer not installed')
def test_grader_phase1b_oracles_and_negative_controls(tmp_path):
    from lxml import etree

    import docx_package

    bench = _load_bench()
    corpus_dir = tmp_path / 'corpus'
    bench.generate_corpus(corpus_dir)
    cases = {c['id']: c for c in json.loads(CASES_PATH.read_text(encoding='utf-8'))['cases']}

    # T14 oracle: cell margins + 12pt data font keeping 14pt bold header -> Hit = 1
    t14 = cases['T14_table_cell_margins_and_font_nd30']
    src14 = corpus_dir / f"{t14['id']}.docx"
    m14 = map_docx.map_document(src14)
    a14 = m14['tables'][1]['anchor']
    sub14 = tmp_path / 'sub' / t14['id']
    sub14.mkdir(parents=True)
    layout_edit.apply_layout_plan(
        src14,
        sub14 / f"{t14['id']}.docx",
        [
            {'op': 'table.set_cell_margins', 'target': a14, 'top_cm': 0.12, 'bottom_cm': 0.12, 'left_cm': 0.18, 'right_cm': 0.18},
            {'op': 'table.set_font_size', 'target': a14, 'size_pt': 12.0, 'preserve_header_size': True},
        ],
        profile='administrative',
    )
    (sub14 / f"{t14['id']}.result.json").write_text(json.dumps({'status': 'edited', 'reason': 'ok'}), encoding='utf-8')
    res14 = bench.grade_case(t14, src14, sub14)
    assert res14['hit'] == 1, res14
    assert res14['telemetry']['xml_ntpr'] == 1.0

    # T15 oracle: image.scale preserving aspect ratio + center -> Hit = 1
    t15 = cases['T15_image_scale_aspect_ratio']
    src15 = corpus_dir / f"{t15['id']}.docx"
    m15 = map_docx.map_document(src15)
    img15 = m15['images'][0]['anchor']
    sub15 = tmp_path / 'sub' / t15['id']
    sub15.mkdir(parents=True)
    layout_edit.apply_layout_plan(
        src15,
        sub15 / f"{t15['id']}.docx",
        [{'op': 'image.scale', 'target': img15, 'width_cm': 12.0, 'align': 'center'}],
    )
    (sub15 / f"{t15['id']}.result.json").write_text(json.dumps({'status': 'edited', 'reason': 'ok'}), encoding='utf-8')
    res15 = bench.grade_case(t15, src15, sub15)
    assert res15['hit'] == 1, res15
    assert res15['telemetry']['xml_ntpr'] == 1.0

    # T15 negative control: distorted aspect ratio (width changed without height) -> Hit = 0
    bad15_dir = tmp_path / 'sub_bad' / t15['id']
    bad15_dir.mkdir(parents=True)
    entries15, parts15 = docx_package.read_package(src15)
    root15 = etree.fromstring(parts15['word/document.xml'])
    ext15 = root15.xpath('.//w:drawing', namespaces={'w': docx_package.W})[0].find(
        f'{{{map_docx.WP}}}inline/{{{map_docx.WP}}}extent'
    )
    new_cx = str(round(12.0 * map_docx.EMU_PER_CM))
    ext15.set('cx', new_cx)
    for ax in root15.xpath('.//w:drawing', namespaces={'w': docx_package.W})[0].xpath('.//a:xfrm/a:ext', namespaces={'a': map_docx.A}):
        ax.set('cx', new_cx)
    docx_package.write_package(
        src15,
        bad15_dir / f"{t15['id']}.docx",
        [e for e in entries15 if not e.is_dir()],
        parts15,
        {'word/document.xml': docx_package.xml_bytes(root15)},
    )
    (bad15_dir / f"{t15['id']}.result.json").write_text(json.dumps({'status': 'edited', 'reason': 'bad ar'}), encoding='utf-8')
    res_bad15 = bench.grade_case(t15, src15, bad15_dir)
    assert res_bad15['hit'] == 0
    assert 'ooxml_target_image' in res_bad15['failure_reasons']

    # T16 oracle: Party HD 05 paragraph format on P4..P6 -> Hit = 1
    t16 = cases['T16_party_hd05_paragraph_format']
    src16 = corpus_dir / f"{t16['id']}.docx"
    m16 = map_docx.map_document(src16)
    p16_targets = [m16['paragraphs'][i]['anchor'] for i in (3, 4, 5)]
    sub16 = tmp_path / 'sub' / t16['id']
    sub16.mkdir(parents=True)
    layout_edit.apply_layout_plan(
        src16,
        sub16 / f"{t16['id']}.docx",
        [{
            'op': 'paragraph.set_format',
            'targets': p16_targets,
            'align': 'both',
            'first_line_indent_cm': 1.0,
            'line_spacing_exact_pt': 20.0,
            'space_after_pt': 6.0,
        }],
        profile='party',
    )
    (sub16 / f"{t16['id']}.result.json").write_text(json.dumps({'status': 'edited', 'reason': 'ok'}), encoding='utf-8')
    res16 = bench.grade_case(t16, src16, sub16)
    assert res16['hit'] == 1, res16
    assert res16['telemetry']['xml_ntpr'] == 1.0

    # T16 negative control: 1.15 auto line spacing instead of exact 18-22pt -> Hit = 0
    bad16_dir = tmp_path / 'sub_bad' / t16['id']
    bad16_dir.mkdir(parents=True)
    layout_edit.apply_layout_plan(
        src16,
        bad16_dir / f"{t16['id']}.docx",
        [{
            'op': 'paragraph.set_format',
            'targets': p16_targets,
            'align': 'both',
            'first_line_indent_cm': 1.0,
            'line_spacing': 1.15,
            'space_after_pt': 6.0,
        }],
    )
    (bad16_dir / f"{t16['id']}.result.json").write_text(json.dumps({'status': 'edited', 'reason': 'auto'}), encoding='utf-8')
    res_bad16 = bench.grade_case(t16, src16, bad16_dir)
    assert res_bad16['hit'] == 0
    assert 'ooxml_target_paragraph' in res_bad16['failure_reasons']

    # T17 oracle: fail-closed refusal on indistinguishable identical tables -> Hit = 1
    t17 = cases['T17_screenshot_ambiguous_identical_tables_refusal']
    src17 = corpus_dir / f"{t17['id']}.docx"
    sub17 = tmp_path / 'sub' / t17['id']
    sub17.mkdir(parents=True)
    (sub17 / f"{t17['id']}.result.json").write_text(
        json.dumps({'status': 'refused', 'reason': 'Hai bảng giống hệt nhau, ảnh cắt không đủ phân biệt'}),
        encoding='utf-8',
    )
    res17 = bench.grade_case(t17, src17, sub17)
    assert res17['hit'] == 1, res17

    # T18 oracle: combined multi-op (table + image + caption paragraph) on long doc -> Hit = 1
    t18 = cases['T18_combined_multi_op_long_doc']
    src18 = corpus_dir / f"{t18['id']}.docx"
    m18 = map_docx.map_document(src18)
    tbl18 = m18['tables'][1]['anchor']
    img18 = m18['images'][0]['anchor']
    cap18 = next(
        p['anchor'] for p in m18['paragraphs']
        if p['text_preview'].startswith('Hình 1: Sơ đồ kiến trúc tổng thể hệ thống')
    )
    sub18 = tmp_path / 'sub' / t18['id']
    sub18.mkdir(parents=True)
    layout_edit.apply_layout_plan(
        src18,
        sub18 / f"{t18['id']}.docx",
        [
            {'op': 'table.scale_width', 'target': tbl18, 'factor': 0.8},
            {'op': 'table.set_alignment', 'target': tbl18, 'align': 'center'},
            {'op': 'table.set_font_size', 'target': tbl18, 'size_pt': 12.0, 'preserve_header_size': True},
            {'op': 'image.scale', 'target': img18, 'width_cm': 11.5, 'align': 'center'},
            {'op': 'paragraph.set_format', 'target': cap18, 'align': 'center', 'keep_with_next': True},
        ],
    )
    (sub18 / f"{t18['id']}.result.json").write_text(json.dumps({'status': 'edited', 'reason': 'ok'}), encoding='utf-8')
    res18 = bench.grade_case(t18, src18, sub18)
    assert res18['hit'] == 1, res18
    assert res18['telemetry']['xml_ntpr'] == 1.0


@pytest.mark.skipif(not find_renderer(), reason='LibreOffice renderer not installed')
def test_grader_phase2_and_phase3_oracles(tmp_path):
    bench = _load_bench()
    corpus_dir = tmp_path / 'corpus'
    bench.generate_corpus(corpus_dir)
    cases = {c['id']: c for c in json.loads(CASES_PATH.read_text(encoding='utf-8'))['cases']}

    for cid in (
        'T21_style_update_normal_global',
        'T24_track_changes_table_scale',
        'T25_auto_paginate_orphan_nd30_signature',
        'T27_complex_word365_namespaces_mc_alternate',
        'T28_signed_docx_refusal',
    ):
        case = cases[cid]
        src = corpus_dir / f'{cid}.docx'
        sub_dir = tmp_path / 'sub_p3' / cid
        exec_info = bench._execute_oracle_case(case, src, sub_dir)
        res = bench.grade_case(case, src, sub_dir)
        assert res['hit'] == 1, res
        assert res['telemetry']['xml_ntpr'] == 1.0
        if cid == 'T25_auto_paginate_orphan_nd30_signature':
            assert exec_info['report']['orphan_signature_before'] is True
            assert exec_info['report']['pages_before'] == 2
            assert exec_info['report']['pages_after'] == 1
            assert exec_info['report']['orphan_signature_after'] is False
        elif cid == 'T28_signed_docx_refusal':
            assert exec_info['mode'] == 'refuse'
            assert not (sub_dir / f'{cid}.docx').exists()
