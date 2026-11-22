"""Reproducible benchmark corpus builder, payload measurer, blind grader, and Live A/B analyzer."""
import argparse
import hashlib
import json
import math
import shutil
import statistics
import struct
import sys
import tempfile
import unicodedata
import zipfile
import zlib
from datetime import datetime
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT, WD_SECTION_START
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt
from lxml import etree

REPO_ROOT = Path(__file__).resolve().parents[1]
CASES_PATH = Path(__file__).resolve().parent / 'layout_edit_cases.json'
ASSETS_TEMPLATES = REPO_ROOT / 'assets' / 'templates'

COMPACT_LAYOUT_SKILL_BLOCK = """Controlled Layout Editing Workflow (`vndocs` — NO raw XML surgery):
1. Map first: `vndocs map <src.docx> --format markdown` (add `--include-paragraphs` for paragraph anchors `P#h4`, `--format json` only if `source_sha256` or DXA needed).
2. Fail-closed safety & Screenshot Grounding:
   - NEVER modify `[LAYOUT:...]` tables unless explicitly asked. If no valid data target exists or `< 0.51 cm` text col width is requested, STOP & report `status: "refused"`.
   - If `user_screenshot.png` is present, inspect it via `read_file` and match unique row IDs/text/structure against `map` FIRST. Only run `vndocs render <src.docx> --output-dir _render --png --annotate-marks` (`page-N-marks.png`) if `>=2` candidates remain ambiguous or map text is insufficient. If `>=2` objects are structurally & textually identical and the crop omits headings/page numbers, STOP & report `status: "refused"` (do NOT guess).
3. Plan JSON (`{"operations": [...]}`), then run `vndocs layout-edit <src.docx> <dst.docx> --plan plan.json`:
   - `table.scale_width`: `{op, target: "T#h4", factor | target_width_pct | target_width_cm}`
   - `table.set_col_widths`: `{op, target, widths: ["20%", "50%", "30%"] | ["2.5cm", ...]}` (1 per `gridCol`)
   - `table.set_alignment`: `{op, target, align: "left|center|right"}`
   - `table.set_cell_margins`: `{op, target, top_cm?, bottom_cm?, left_cm?, right_cm?, clear_cell_overrides?: true}`
   - `table.set_font_size`: `{op, target, size_pt, preserve_header_size?: true}` (NĐ30 tables `>= 12pt`)
   - `image.scale`: `{op, target: "IMG#h4", factor | width_cm | width_pct, align?}` (auto preserves aspect ratio)
   - `paragraph.set_format`: `{op, target | targets: ["P#h4",...], align?, space_before_pt?, space_after_pt?, line_spacing? | line_spacing_exact_pt?, first_line_indent_cm?, keep_with_next?}`
4. Vietnamese Standards & Verify:
   - NĐ30: pass `--profile administrative`, protect Quốc hiệu/Nơi nhận, tables `>= 12pt`.
   - Party (HD 05-HD/VPTW 2026): pass `--profile party`; body paragraphs use `align="both"`, `first_line_indent_cm=1.0`, `line_spacing_exact_pt=20.0` (18-22pt exact), `space_after_pt=6.0`.
   - Complexity Contract (`max_iter=2`): skip `diff-layout` on clean explicit numeric edits with zero warnings; run `vndocs diff-layout <src.docx> <dst.docx> --target <anchor>` when checking overflow, ripple, or relative sizing."""


def _tools():
    scripts_dir = str(REPO_ROOT / 'scripts')
    if scripts_dir not in sys.path:
        sys.path.insert(0, scripts_dir)
    import auto_paginate
    import build_docx
    import diff_layout
    import docx_package
    import generate_docx
    import layout_edit
    import map_docx
    import ooxml_order
    import render_docx
    import validate_docx
    return {
        'auto_paginate': auto_paginate,
        'build_docx': build_docx,
        'diff_layout': diff_layout,
        'docx_package': docx_package,
        'generate_docx': generate_docx,
        'layout_edit': layout_edit,
        'map_docx': map_docx,
        'ooxml_order': ooxml_order,
        'render_docx': render_docx,
        'validate_docx': validate_docx,
    }


def _sha256(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def _set_col_widths_cm(tbl, widths_cm):
    t = _tools()
    w_ns = f"{{{t['docx_package'].W}}}"
    tbl.autofit = False
    dxa_list = [round(w * t['map_docx'].DXA_PER_CM) for w in widths_cm]
    for idx, w_cm in enumerate(widths_cm):
        width = Cm(w_cm)
        tbl.columns[idx].width = width
        for row in tbl.rows:
            row.cells[idx].width = width
    tbl_grid = tbl._tbl.find(f'{w_ns}tblGrid')
    if tbl_grid is not None:
        for gcol, w_dxa in zip(tbl_grid.findall(f'{w_ns}gridCol'), dxa_list):
            gcol.set(f'{w_ns}w', str(w_dxa))
    tbl_pr = tbl._tbl.find(f'{w_ns}tblPr')
    if tbl_pr is not None:
        tbl_w = tbl_pr.find(f'{w_ns}tblW')
        if tbl_w is not None:
            tbl_w.set(f'{w_ns}w', str(sum(dxa_list)))
            tbl_w.set(f'{w_ns}type', 'dxa')
    return tbl


def _add_prose_block(doc, heading, count, topic, *, nfd=False):
    t = _tools()
    norm = (lambda s: unicodedata.normalize('NFD', s)) if nfd else (lambda s: s)
    h_para = t['build_docx'].add_body(doc, norm(heading))
    for run in h_para.runs:
        run.bold = True
    for idx in range(1, count + 1):
        sentence = (
            f'Đoạn {idx} ({topic}): Báo cáo ghi nhận tiến độ triển khai công tác rà soát dữ liệu, '
            f'chuẩn hóa biểu mẫu hành chính và kiểm tra tính nhất quán số liệu giữa các đơn vị trực thuộc '
            f'theo đúng kế hoạch đã được lãnh đạo phê duyệt trong kỳ đánh giá hiện hành.'
        )
        p = t['build_docx'].add_body(doc, norm(sentence))
        p.paragraph_format.space_after = Pt(4)


def _make_png(path, width=800, height=400, rgb=(24, 92, 168)):
    """Create a valid truecolor PNG with an inner contrast frame using pure Python stdlib."""
    def _chunk(tag, data):
        crc = zlib.crc32(tag + data) & 0xFFFFFFFF
        return struct.pack('!I', len(data)) + tag + data + struct.pack('!I', crc)

    border = bytes((max(0, rgb[0] - 18), max(0, rgb[1] - 18), max(0, rgb[2] - 18)))
    accent = bytes((min(255, rgb[0] + 150), min(255, rgb[1] + 130), min(255, rgb[2] + 90)))
    fill = bytes(rgb)
    raw = bytearray()
    for y in range(height):
        raw.append(0)
        if y < 6 or y >= height - 6:
            raw.extend(border * width)
        elif height // 3 <= y <= (2 * height) // 3:
            side = width // 8
            raw.extend(fill * side + accent * (width - 2 * side) + fill * side)
        else:
            raw.extend(fill * width)
    ihdr = struct.pack('!IIBBBBB', width, height, 8, 2, 0, 0, 0)
    payload = (
        b'\x89PNG\r\n\x1a\n'
        + _chunk(b'IHDR', ihdr)
        + _chunk(b'IDAT', zlib.compress(bytes(raw), 6))
        + _chunk(b'IEND', b'')
    )
    out = Path(path)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_bytes(payload)
    return out


def _generate_table_screenshot(docx_path, table_ordinal, out_png):
    """Render DOCX via LibreOffice and extract a tight table-only PNG crop without headings or page numbers."""
    t = _tools()
    if not t['render_docx'].find_renderer():
        return None
    import pymupdf

    with tempfile.TemporaryDirectory(prefix='vndocs-shot-') as tmp_dir:
        ren = t['render_docx'].render_doc(docx_path, output_dir=tmp_dir, png=False)
        if ren.get('status') != 'rendered' or not ren.get('pdf'):
            return None
        mapping = t['map_docx'].map_document(docx_path, pdf_path=ren['pdf'], include_row_texts=True)
        if len(mapping['tables']) < table_ordinal:
            return None
        pdf_info = mapping['tables'][table_ordinal - 1].get('pdf') or {}
        if not pdf_info.get('matched') or not pdf_info.get('pdf_pages') or not pdf_info.get('bbox_pt'):
            return None
        page_num = pdf_info['pdf_pages'][0]
        x0, y0, x1, y1 = pdf_info['bbox_pt']
        with pymupdf.open(ren['pdf']) as pdf_doc:
            page = pdf_doc[page_num - 1]
            clip = pymupdf.Rect(
                max(page.rect.x0, x0 - 6.0),
                max(page.rect.y0, y0 - 4.0),
                min(page.rect.x1, x1 + 6.0),
                min(page.rect.y1, y1 + 4.0),
            )
            pix = page.get_pixmap(matrix=pymupdf.Matrix(2.0, 2.0), clip=clip)
            out_path = Path(out_png)
            out_path.parent.mkdir(parents=True, exist_ok=True)
            pix.save(str(out_path))
            return out_path


def _add_nd30_header(doc):
    t = _tools()
    tbl = doc.add_table(rows=1, cols=2)
    tbl.cell(0, 0).text = 'BỘ KHOA HỌC VÀ CÔNG NGHỆ\nCỤC CÔNG NGHỆ THÔNG TIN'
    tbl.cell(0, 1).text = 'CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\nĐộc lập - Tự do - Hạnh phúc'
    t['generate_docx']._borderless(tbl)


def _add_nd30_signature(doc):
    t = _tools()
    tbl = doc.add_table(rows=1, cols=2)
    tbl.cell(0, 0).text = 'Nơi nhận:\n- Như trên;\n- Lưu: VT, KHTC.'
    tbl.cell(0, 1).text = 'KT. CỤC TRƯỞNG\nPHÓ CỤC TRƯỞNG\n\n\nNguyễn Văn Minh'
    t['generate_docx']._borderless(tbl)


def _build_generated_doc(generator_name, dst_path):
    t = _tools()
    build = t['build_docx']
    doc = Document()

    if generator_name == 'nd30_with_budget':
        build.setup_document(doc, 'administrative')
        _add_nd30_header(doc)
        build.add_body(doc, 'Kính gửi: Các đơn vị trực thuộc')
        build.add_body(doc, 'Tổng hợp dự toán kinh phí triển khai nhiệm vụ chuyển đổi số năm 2026 như sau:')
        t_budget = build.add_table(
            doc,
            ['TT', 'Nội dung chi', 'Đơn vị tính', 'Dự toán (đồng)'],
            [
                ['1', 'Khảo sát hiện trạng hạ tầng số tại các đơn vị', 'Gói', '120.000.000'],
                ['2', 'Thẩm định, kiểm thử và nghiệm thu hệ thống', 'Hạng mục', '85.500.000'],
                ['3', 'Đào tạo chuyển giao quy trình vận hành chuẩn hóa', 'Lớp', '42.000.000'],
            ],
            profile='administrative',
        )
        _set_col_widths_cm(t_budget, [1.3, 8.7, 2.5, 4.0])
        build.add_body(doc, 'Đề nghị các đơn vị phối hợp thực hiện đúng tiến độ đã phân công.')
        _add_nd30_signature(doc)

    elif generator_name == 'long_multi_similar_tables':
        build.setup_document(doc, 'academic')
        for sec_num in range(1, 5):
            if sec_num == 1:
                heading = 'Mục 1. Khảo sát hiện trạng và tổng hợp nhân lực'
                hdr = ['Mã', 'Hạng mục nhân sự', 'Giá trị (triệu đồng)']
                rows = [['M1-01', 'Phân tích nhu cầu đào tạo chuyên môn', '35.0'], ['M1-02', 'Đánh giá mức độ sẵn sàng vận hành', '40.0']]
            elif sec_num == 2:
                heading = 'Mục 2. Kế hoạch triển khai kỹ thuật'
                hdr = ['Mã', 'Hạng mục kỹ thuật', 'Giá trị (triệu đồng)']
                rows = [['M2-01', 'Chuẩn hóa kho dữ liệu dùng chung', '90.0'], ['M2-02', 'Kiểm thử an toàn thông tin định kỳ', '65.0']]
            elif sec_num == 3:
                heading = 'Mục 3. Dự toán kinh phí chi tiết năm 2026'
                hdr = ['Mã', 'Hạng mục kinh phí', 'Giá trị (triệu đồng)']
                rows = [['M3-01', 'Mua sắm thiết bị lưu trữ dự phòng', '150.0'], ['M3-02', 'Thuê dịch vụ kiểm định phần mềm', '80.0']]
            else:
                heading = 'Mục 4. Tổng hợp nghiệm thu đầu ra'
                hdr = ['Mã', 'Hạng mục nghiệm thu', 'Giá trị (triệu đồng)']
                rows = [['M4-01', 'Hồ sơ tài liệu bàn giao kỹ thuật', '25.0'], ['M4-02', 'Biên bản đánh giá sau vận hành', '15.0']]
            _add_prose_block(doc, heading, 18, f'Mục {sec_num}')
            tbl = build.add_table(doc, hdr, rows, profile='academic')
            _set_col_widths_cm(tbl, [2.2, 9.8, 4.5])
        _add_prose_block(doc, 'Kết luận chung toàn báo cáo', 8, 'Kết luận')

    elif generator_name == 'overflow_wide_table':
        build.setup_document(doc, 'academic')
        build.add_body(doc, 'Báo cáo kiểm tra bảng biểu bị tràn lề phải trên khổ giấy A4:')
        t_wide = build.add_table(
            doc,
            ['Mã', 'Hạng mục', 'Phụ trách', 'Kết quả'],
            [
                ['A1', 'Phân tích hiệu năng render tài liệu hành chính', 'Nhóm 1', 'Hoàn thành'],
                ['A2', 'Kiểm tra tràn khung trang khi cột quá rộng', 'Nhóm 2', 'Cần chỉnh lại'],
            ],
            profile='academic',
        )
        _set_col_widths_cm(t_wide, [2.5, 9.5, 4.0, 4.0])  # 20.0 cm > 16.5 cm

    elif generator_name == 'merged_gridspan_vmerge_med':
        build.setup_document(doc, 'academic')
        _add_prose_block(doc, '1. Thông tin chung dự án', 14, 'Tổng quan')
        t_sum = build.add_table(
            doc,
            ['Thông tin', 'Chi tiết'],
            [['Tên nhiệm vụ', 'Chuẩn hóa quy trình biên tập tài liệu điện tử'], ['Đơn vị chủ trì', 'Trung tâm Công nghệ Thông tin']],
            profile='academic',
        )
        _set_col_widths_cm(t_sum, [4.5, 12.0])
        _add_prose_block(doc, '2. Bảng nghiệm thu theo giai đoạn (có ô gộp)', 16, 'Nghiệm thu')
        t_merged = doc.add_table(rows=4, cols=4)
        t_merged.style = 'Table Grid'
        t_merged.cell(0, 0).merge(t_merged.cell(0, 1))
        t_merged.cell(0, 0).text = 'Thông tin hạng mục'
        t_merged.cell(0, 2).merge(t_merged.cell(0, 3))
        t_merged.cell(0, 2).text = 'Giá trị nghiệm thu'
        t_merged.cell(1, 0).merge(t_merged.cell(2, 0))
        t_merged.cell(1, 0).text = 'Giai đoạn 1'
        t_merged.cell(1, 1).text = 'Thiết kế sơ bộ'
        t_merged.cell(1, 2).text = 'Quý I'
        t_merged.cell(1, 3).text = '50.000.000'
        t_merged.cell(2, 1).text = 'Kiểm thử hệ thống'
        t_merged.cell(2, 2).text = 'Quý II'
        t_merged.cell(2, 3).text = '75.000.000'
        t_merged.cell(3, 0).text = 'Giai đoạn 2'
        t_merged.cell(3, 1).text = 'Triển khai diện rộng'
        t_merged.cell(3, 2).text = 'Quý III'
        t_merged.cell(3, 3).text = '120.000.000'
        _add_prose_block(doc, '3. Kiến nghị tiếp theo', 10, 'Kiến nghị')

    elif generator_name == 'long_landscape_multisection':
        build.setup_document(doc, 'academic')
        _add_prose_block(doc, 'Phần I. Thuyết minh kết quả khối đứng (Portrait)', 28, 'Khổ đứng 1')
        t1 = build.add_table(
            doc,
            ['TT', 'Hạng mục khổ đứng 1', 'Ghi chú'],
            [['1', 'Rà soát quy chuẩn dữ liệu văn bản', 'Đạt'], ['2', 'Đánh giá thời gian xử lý trung bình', 'Đạt']],
            profile='academic',
        )
        _set_col_widths_cm(t1, [2.0, 10.5, 4.0])
        _add_prose_block(doc, 'Phần II. Tổng hợp đề xuất trước phụ lục', 28, 'Khổ đứng 2')
        t2 = build.add_table(
            doc,
            ['TT', 'Hạng mục khổ đứng 2', 'Trạng thái'],
            [['1', 'Kiểm tra biểu mẫu phòng ban', 'Hoàn tất'], ['2', 'Phê duyệt danh mục phụ lục', 'Hoàn tất']],
            profile='academic',
        )
        _set_col_widths_cm(t2, [2.0, 10.5, 4.0])
        sec2 = doc.add_section(WD_SECTION_START.NEW_PAGE)
        sec2.orientation = WD_ORIENT.LANDSCAPE
        sec2.page_width = Cm(29.7)
        sec2.page_height = Cm(21.0)
        sec2.left_margin = Cm(2.0)
        sec2.right_margin = Cm(2.0)
        sec2.top_margin = Cm(2.0)
        sec2.bottom_margin = Cm(2.0)
        _add_prose_block(doc, 'Phụ lục A. Bảng chỉ tiêu KPI (Trang khổ ngang A4)', 8, 'Khổ ngang')
        t3 = build.add_table(
            doc,
            ['STT', 'Tên đơn vị phụ trách', 'KPI Q1', 'KPI Q2', 'Đánh giá chung'],
            [
                ['1', 'Trung tâm Dữ liệu Quốc gia', '98%', '96%', 'Đạt chuẩn xuất sắc'],
                ['2', 'Phòng Hạ tầng và An toàn mạng', '92%', '95%', 'Đạt yêu cầu'],
            ],
            profile='academic',
        )
        _set_col_widths_cm(t3, [1.7, 10.0, 4.5, 4.5, 5.0])  # 25.7 cm printable width

    elif generator_name == 'long_multipage_table':
        build.setup_document(doc, 'academic')
        _add_prose_block(doc, 'Phần 1. Tóm tắt điều hành dự án', 24, 'Mở đầu')
        t_intro = build.add_table(
            doc,
            ['Chỉ tiêu', 'Kết quả tổng hợp'],
            [['Số đơn vị tham gia', '12 đơn vị'], ['Tỷ lệ hoàn thành đúng hạn', '94.5%']],
            profile='academic',
        )
        _set_col_widths_cm(t_intro, [6.5, 10.0])
        _add_prose_block(doc, 'Phần 2. Thuyết minh chi tiết trước phụ lục thống kê', 28, 'Thuyết minh')
        rows28 = [
            [str(i), f'Đầu việc thống kê số {i:02d} thuộc kế hoạch chuẩn hóa dữ liệu hành chính', f'{min(100, i * 4)}%', 'Hoàn thành']
            for i in range(1, 29)
        ]
        t_long = build.add_table(doc, ['TT', 'Mô tả đầu việc', 'Tiến độ', 'Trạng thái'], rows28, profile='academic')
        _set_col_widths_cm(t_long, [1.4, 9.5, 2.5, 3.1])
        build.add_body(doc, 'KẾT THÚC BẢNG THỐNG KÊ 28 DÒNG ĐA TRANG.')

    elif generator_name == 'nfd_vietnamese_med':
        build.setup_document(doc, 'academic')
        _add_prose_block(doc, '1. Giới thiệu văn bản mã tổ hợp macOS (NFD)', 16, 'NFD mở đầu', nfd=True)
        nfd = lambda s: unicodedata.normalize('NFD', s)  # noqa: E731
        t_nfd = build.add_table(
            doc,
            [nfd('STT'), nfd('Tên đơn vị phối hợp'), nfd('Trạng thái xử lý')],
            [
                ['1', nfd('Sở Khoa học và Công nghệ Thành phố Hà Nội'), nfd('Đã tiếp nhận hồ sơ')],
                ['2', nfd('Trung tâm Chuyển đổi số và Thống kê địa phương'), nfd('Đang rà soát định dạng')],
            ],
            profile='academic',
        )
        _set_col_widths_cm(t_nfd, [1.5, 10.5, 4.5])
        _add_prose_block(doc, '2. Lịch dự kiến các phiên làm việc tiếp theo', 14, 'NFD lịch họp', nfd=True)
        t_sched = build.add_table(
            doc,
            [nfd('Phiên họp'), nfd('Nội dung thảo luận chính')],
            [[nfd('Phiên 1'), nfd('Thống nhất tiêu chuẩn biên soạn nội bộ')]],
            profile='academic',
        )
        _set_col_widths_cm(t_sched, [4.5, 12.0])

    elif generator_name == 'impossible_narrow_col_med':
        build.setup_document(doc, 'academic')
        _add_prose_block(doc, '1. Nội dung thuyết minh trước bảng số liệu', 24, 'Thuyết minh')
        t_dense = build.add_table(
            doc,
            ['Mã', 'Hạng mục công việc', 'Ghi chú', 'Kết quả'],
            [
                ['01', 'Kiểm tra giới hạn độ rộng cột tối thiểu cho văn bản', 'Bắt buộc giữ chữ đọc được', 'Đạt'],
                ['02', 'Xác minh giữ nguyên các đoạn ngoài bảng', 'Không biến đổi văn bản', 'Đạt'],
            ],
            profile='academic',
        )
        _set_col_widths_cm(t_dense, [2.0, 7.5, 4.5, 2.5])

    elif generator_name == 'nearly_full_page_ripple':
        build.setup_document(doc, 'academic')
        for i in range(13):
            p = build.add_body(doc, f'Đoạn trình bày số {i + 1}: tổng hợp điều kiện kích hoạt nhảy trang khi bảng thay đổi chiều cao.')
            p.paragraph_format.space_after = Pt(6)
        t_rip = build.add_table(
            doc,
            ['TT', 'Nội dung chi tiết hạng mục dễ bị xuống dòng khi thu hẹp cột', 'Giá trị'],
            [
                [str(i), f'Hạng mục kỹ thuật số {i} có mô tả đủ dài để xuống dòng nhiều hơn khi thu hẹp bề rộng cột giữa', '10.000.000']
                for i in range(1, 6)
            ],
            profile='academic',
        )
        _set_col_widths_cm(t_rip, [1.5, 11.5, 3.5])
        build.add_body(doc, 'ĐOẠN XÁC NHẬN CUỐI TRANG ĐỂ ĐO PAGINATION RIPPLE')

    elif generator_name == 'screenshot_three_similar_tables':
        build.setup_document(doc, 'academic')
        entries = [
            ('Mục 1. Thống kê hạng mục Quý I', [['Q1-01', 'Khảo sát dữ liệu cơ sở tại đơn vị', '31.400.000'], ['Q1-02', 'Chuẩn hóa biểu mẫu hành chính Quý I', '42.100.000']]),
            ('Mục 2. Thống kê hạng mục Quý II', [['Q2-01', 'Kiểm thử luồng xử lý văn bản điện tử', '48.200.000'], ['Q2-02', 'Đánh giá tương thích đa thiết bị Quý II', '63.500.000']]),
            ('Mục 3. Thống kê hạng mục Quý III', [['Q3-01', 'Nghiệm thu triển khai toàn hệ thống', '91.000.000'], ['Q3-02', 'Đào tạo chuyển giao quy trình vận hành', '54.800.000']]),
        ]
        for idx, (heading, rows) in enumerate(entries, start=1):
            _add_prose_block(doc, heading, 10, f'Phần {idx}')
            doc.paragraphs[-1].paragraph_format.space_after = Pt(12)
            tbl = build.add_table(doc, ['Mã', 'Hạng mục công việc', 'Kinh phí (đồng)'], rows, profile='academic')
            _set_col_widths_cm(tbl, [2.4, 9.6, 4.5])
            spacer = build.add_body(doc, f'Nhận xét tóm tắt cho {heading.lower()}.')
            spacer.paragraph_format.space_before = Pt(10)
        _add_prose_block(doc, 'Mục 4. Tổng hợp chung toàn báo cáo', 6, 'Kết')

    elif generator_name == 'nd30_dense_table_font_margins':
        build.setup_document(doc, 'administrative')
        _add_nd30_header(doc)
        build.add_body(doc, 'Kính gửi: Các đơn vị trực thuộc')
        build.add_body(doc, 'Tổng hợp danh mục nhiệm vụ và kinh phí triển khai năm 2026 như sau:')
        t_data = build.add_table(
            doc,
            ['TT', 'Hạng mục triển khai', 'Đơn vị', 'Dự toán (đồng)'],
            [
                ['1', 'Chuẩn hóa quy trình tiếp nhận hồ sơ trực tuyến', 'Văn phòng', '145.000.000'],
                ['2', 'Nâng cấp hạ tầng lưu trữ và sao lưu định kỳ', 'Phòng Kỹ thuật', '210.000.000'],
                ['3', 'Tập huấn sử dụng chữ ký số chuyên dùng', 'Phòng Tổ chức', '68.000.000'],
            ],
            profile='administrative',
        )
        _set_col_widths_cm(t_data, [1.5, 8.0, 2.5, 4.5])
        for r_idx, row in enumerate(t_data.rows):
            for cell in row.cells:
                for p in cell.paragraphs:
                    for run in p.runs:
                        run.font.size = Pt(14)
                        run.bold = (r_idx == 0)
        # Seed initial narrow tblCellMar and a stray tcMar override to test clean harmonization.
        w_ns = f"{{{t['docx_package'].W}}}"
        tbl_pr = t['layout_edit']._ensure_pr(t_data._tbl, 'tblPr')
        init_mar = etree.Element(f'{w_ns}tblCellMar')
        for side in ('top', 'left', 'bottom', 'right'):
            m_el = etree.SubElement(init_mar, f'{w_ns}{side}')
            m_el.set(f'{w_ns}w', '28')
            m_el.set(f'{w_ns}type', 'dxa')
        t['ooxml_order'].insert_ordered_child(tbl_pr, init_mar)
        tc_pr = t['layout_edit']._ensure_pr(t_data.rows[1].cells[1]._tc, 'tcPr')
        stale_mar = etree.Element(f'{w_ns}tcMar')
        stale_left = etree.SubElement(stale_mar, f'{w_ns}left')
        stale_left.set(f'{w_ns}w', '420')
        stale_left.set(f'{w_ns}type', 'dxa')
        t['ooxml_order'].insert_ordered_child(tc_pr, stale_mar)

        build.add_body(doc, 'Đề nghị các đơn vị khẩn trương thực hiện đúng kế hoạch.')
        _add_nd30_signature(doc)

    elif generator_name == 'overflow_image_and_icon':
        build.setup_document(doc, 'academic')
        with tempfile.TemporaryDirectory(prefix='vndocs-img-asset-') as asset_dir:
            diag = _make_png(Path(asset_dir) / 'architecture.png', 900, 450, (22, 95, 165))
            icon = _make_png(Path(asset_dir) / 'badge.png', 180, 180, (185, 55, 45))
            build.add_body(doc, '1. Sơ đồ kiến trúc tổng thể (hình dưới đang quá rộng so với lề giấy):')
            doc.add_picture(str(diag), width=Cm(18.5))
            _add_prose_block(doc, '2. Biểu tượng xác nhận kèm theo', 3, 'Mục ảnh nhỏ')
            doc.add_picture(str(icon), width=Cm(3.0))

    elif generator_name == 'party_hd05_misformatted_body':
        build.setup_document(doc, 'party')
        t_hdr = doc.add_table(rows=1, cols=2)
        t_hdr.cell(0, 0).text = 'BAN CHẤP HÀNH TRUNG ƯƠNG\nVĂN PHÒNG'
        t_hdr.cell(0, 1).text = 'ĐẢNG CỘNG SẢN VIỆT NAM\n*'
        t['generate_docx']._borderless(t_hdr)

        p1 = build.add_body(doc, 'BÁO CÁO')
        p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in p1.runs:
            r.bold = True
            r.font.size = Pt(15)
        p2 = build.add_body(doc, 'Kết quả công tác chuyển đổi số quý I năm 2026')
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in p2.runs:
            r.bold = True
            r.font.size = Pt(14)
        p3 = build.add_body(doc, 'I. KẾT QUẢ THỰC HIỆN')
        p3.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p3.paragraph_format.space_before = Pt(6)
        p3.paragraph_format.space_after = Pt(6)
        for r in p3.runs:
            r.bold = True
            r.font.size = Pt(14)

        body_texts = [
            'Trong quý I năm 2026, các cơ quan tham mưu đã hoàn thành rà soát hạ tầng truyền dẫn và triển khai ký số văn bản điện tử thông suốt tới toàn bộ chi bộ trực thuộc.',
            'Công tác lưu trữ hồ sơ công việc trên môi trường số bảo đảm tính chính xác, thống nhất theo danh mục hồ sơ lưu trữ cơ quan Đảng và quy trình bảo mật hai lớp.',
            'Trong thời gian tới, các đơn vị tiếp tục hoàn thiện quy chuẩn dữ liệu dùng chung, tăng cường tập huấn kỹ năng khai thác phần mềm xử lý văn bản cho cán bộ chuyên trách.',
        ]
        for txt in body_texts:
            p = build.add_body(doc, txt)
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.first_line_indent = Cm(0)
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_after = Pt(0)
            for r in p.runs:
                r.font.size = Pt(14)

        t_sig = doc.add_table(rows=1, cols=2)
        t_sig.cell(0, 0).text = 'Nơi nhận:\n- Ban Bí thư;\n- Lưu Văn phòng Trung ương.'
        t_sig.cell(0, 1).text = 'CHÁNH VĂN PHÒNG\n\n\nTrần Quốc Hưng'
        t['generate_docx']._borderless(t_sig)

    elif generator_name == 'screenshot_two_identical_tables':
        build.setup_document(doc, 'academic')
        ident_rows = [
            ['01', 'Kiểm tra đồng bộ dữ liệu liên thông', '100%'],
            ['02', 'Rà soát tuân thủ định dạng biểu mẫu', '98%'],
        ]
        _add_prose_block(doc, 'Phụ lục 1. Bảng tổng hợp kiểm tra định kỳ', 14, 'Phụ lục 1')
        doc.paragraphs[-1].paragraph_format.space_after = Pt(12)
        t1 = build.add_table(doc, ['STT', 'Tiêu chí đánh giá', 'Tỷ lệ đạt'], ident_rows, profile='academic')
        _set_col_widths_cm(t1, [2.0, 11.0, 3.5])
        spacer1 = build.add_body(doc, 'Kết thúc bảng phụ lục thứ nhất.')
        spacer1.paragraph_format.space_before = Pt(10)
        doc.add_page_break()
        _add_prose_block(doc, 'Phụ lục 2. Bảng tổng hợp kiểm tra định kỳ', 14, 'Phụ lục 2')
        doc.paragraphs[-1].paragraph_format.space_after = Pt(12)
        t2 = build.add_table(doc, ['STT', 'Tiêu chí đánh giá', 'Tỷ lệ đạt'], ident_rows, profile='academic')
        _set_col_widths_cm(t2, [2.0, 11.0, 3.5])
        spacer2 = build.add_body(doc, 'Kết thúc bảng phụ lục thứ hai.')
        spacer2.paragraph_format.space_before = Pt(10)
        _add_prose_block(doc, 'Phần III. Kiến nghị chung', 6, 'Kết')

    elif generator_name == 'long_combined_table_img_para':
        build.setup_document(doc, 'academic')
        _add_prose_block(doc, 'Phần 1. Tổng quan điều hành dự án', 22, 'Phần 1')
        t1 = build.add_table(
            doc,
            ['Chỉ tiêu', 'Kết quả'],
            [['Số hạng mục', '12'], ['Tiến độ', 'Đạt kế hoạch']],
            profile='academic',
        )
        _set_col_widths_cm(t1, [6.5, 10.0])
        _add_prose_block(doc, 'Phần 2. Kiến trúc hệ thống xử lý tài liệu', 24, 'Phần 2')
        p_cap = build.add_body(doc, 'Hình 1: Sơ đồ kiến trúc tổng thể hệ thống')
        p_cap.alignment = WD_ALIGN_PARAGRAPH.LEFT
        with tempfile.TemporaryDirectory(prefix='vndocs-combo-img-') as asset_dir:
            diag = _make_png(Path(asset_dir) / 'system_arch.png', 800, 400, (30, 115, 90))
            doc.add_picture(str(diag), width=Cm(18.0))
        _add_prose_block(doc, 'Phần 3. Bảng phân bổ kinh phí chi tiết', 26, 'Phần 3')
        t2 = build.add_table(
            doc,
            ['Mã', 'Nội dung chi tiết', 'Ghi chú', 'Kinh phí (đồng)'],
            [
                ['C1', 'Phân tích cấu trúc XML và đo lường render', 'Đợt 1', '95.000.000'],
                ['C2', 'Kiểm thử hồi quy và đánh giá độc lập', 'Đợt 2', '110.000.000'],
            ],
            profile='academic',
        )
        _set_col_widths_cm(t2, [1.8, 7.7, 3.0, 4.0])
        for r_idx, row in enumerate(t2.rows):
            for cell in row.cells:
                for p in cell.paragraphs:
                    for run in p.runs:
                        run.font.size = Pt(14)
                        run.bold = (r_idx == 0)
        _add_prose_block(doc, 'Phần 4. Kết luận và kiến nghị', 14, 'Phần 4')

    elif generator_name == 'table_borders_shading_src':
        build.setup_document(doc, 'academic')
        build.add_body(doc, 'Bảng số liệu tổng hợp trước khi định dạng viền và màu nền:')
        tbl = build.add_table(
            doc,
            ['TT', 'Hạng mục công việc', 'Kết quả'],
            [['1', 'Kiểm thử luồng xử lý', 'Đạt'], ['2', 'Rà soát tính nhất quán XML', 'Đạt']],
            profile='academic',
        )
        _set_col_widths_cm(tbl, [2.0, 10.5, 4.0])

    elif generator_name == 'style_update_medium_src':
        build.setup_document(doc, 'academic')
        _add_prose_block(doc, '1. Tổng quan sử dụng style chung', 12, 'Style chung')
        tbl = build.add_table(
            doc,
            ['STT', 'Tên thành phần', 'Trạng thái'],
            [['1', 'Khối nội dung chuẩn', 'Đạt'], ['2', 'Khối bảng số liệu', 'Đạt']],
            profile='academic',
        )
        _set_col_widths_cm(tbl, [2.0, 10.5, 4.0])
        _add_prose_block(doc, '2. Đề xuất hoàn thiện mẫu chuẩn', 10, 'Đề xuất')

    elif generator_name == 'two_portrait_sections_appendix':
        build.setup_document(doc, 'academic')
        _add_prose_block(doc, 'Phần I. Nội dung chính (Section 1 khổ đứng)', 14, 'Section 1')
        t1 = build.add_table(
            doc,
            ['Chỉ tiêu', 'Kết quả'],
            [['Đơn vị tham gia', '10'], ['Mức độ đáp ứng', 'Tốt']],
            profile='academic',
        )
        _set_col_widths_cm(t1, [6.5, 10.0])
        sec2 = doc.add_section(WD_SECTION_START.NEW_PAGE)
        sec2.orientation = WD_ORIENT.PORTRAIT
        sec2.page_width = Cm(21.0)
        sec2.page_height = Cm(29.7)
        sec2.top_margin = Cm(2.5)
        sec2.bottom_margin = Cm(2.5)
        sec2.left_margin = Cm(3.0)
        sec2.right_margin = Cm(2.0)
        _add_prose_block(doc, 'Phụ lục A. Bảng tổng hợp đa cột (Section 2 cần xoay ngang)', 10, 'Section 2')
        t2 = build.add_table(
            doc,
            ['TT', 'Hạng mục phụ lục', 'Q1', 'Q2', 'Ghi chú'],
            [['1', 'Tối ưu luồng render', '95%', '98%', 'Ổn định']],
            profile='academic',
        )
        _set_col_widths_cm(t2, [1.5, 7.0, 2.5, 2.5, 3.0])

    elif generator_name == 'fragmented_runs_src':
        build.setup_document(doc, 'academic')
        p = doc.add_paragraph()
        p.add_run('Báo cáo ')
        p.add_run('tổng kết ')
        p.add_run('quý I ')
        b_run = p.add_run('2026')
        b_run.bold = True
        p.add_run(' đã ')
        p.add_run('hoàn tất.')
        build.add_body(doc, 'Đoạn thứ hai trình bày kế hoạch quý II giữ nguyên không thay đổi.')
        build.add_body(doc, 'Đoạn thứ ba trình bày kiến nghị kỹ thuật giữ nguyên không thay đổi.')

    elif generator_name == 'nd30_orphan_signature_2p':
        build.setup_document(doc, 'administrative')
        t_hdr = doc.add_table(rows=1, cols=2)
        t_hdr.cell(0, 0).text = 'BỘ KHOA HỌC VÀ CÔNG NGHỆ\nVĂN PHÒNG'
        t_hdr.cell(0, 1).text = 'CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\nĐộc lập - Tự do - Hạnh phúc'
        t['generate_docx']._borderless(t_hdr)
        sec = doc.sections[0]
        sec.top_margin = Cm(2.45)
        sec.bottom_margin = Cm(2.45)
        for idx in range(1, 10):
            p = build.add_body(
                doc,
                f'Căn cứ nội dung công văn số {idx:02d}: Văn phòng Bộ đề nghị các đơn vị trực thuộc khẩn trương '
                f'rà soát lại toàn bộ biểu mẫu báo cáo hành chính định kỳ quý I năm 2026 trước khi trình duyệt.',
            )
            p.paragraph_format.space_after = Pt(8)
        tbl = build.add_table(
            doc,
            ['STT', 'Đơn vị thực hiện', 'Hạn hoàn thành'],
            [['1', 'Phòng Tổng hợp', '25/11/2026'], ['2', 'Phòng Kế hoạch', '28/11/2026']],
            profile='administrative',
        )
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for run in p.runs:
                        run.font.size = Pt(14)
        t_sig = doc.add_table(rows=1, cols=2)
        t_sig.cell(0, 0).text = 'Nơi nhận:\n- Như trên;\n- Lưu: VT.'
        t_sig.cell(0, 1).text = 'CHÁNH VĂN PHÒNG\n\n\nNguyễn Minh Anh'
        t['generate_docx']._borderless(t_sig)

    elif generator_name == 'academic_3pages_src':
        build.setup_document(doc, 'academic')
        doc.sections[0].top_margin = Cm(2.5)
        doc.sections[0].bottom_margin = Cm(2.5)
        for i in range(1, 36):
            p = build.add_body(
                doc,
                f'Đoạn phân tích số {i:02d}: Nội dung tổng hợp đánh giá hiệu năng hệ thống xử lý văn bản điện tử '
                f'và mức độ bảo toàn cấu trúc XML qua các lượt thử nghiệm thực nghiệm độc lập.',
            )
            p.paragraph_format.space_after = Pt(12)

    elif generator_name == 'word365_complex_src':
        build.setup_document(doc, 'academic')
        _add_prose_block(doc, '1. Văn bản Word 365 phức hợp', 12, 'Word 365')
        tbl = build.add_table(
            doc,
            ['Mã', 'Hạng mục kiểm thử', 'Kết quả'],
            [['W1', 'Bảo toàn w14/w15/mc:Ignorable', 'Đạt'], ['W2', 'Bảo toàn mc:AlternateContent', 'Đạt']],
            profile='academic',
        )
        _set_col_widths_cm(tbl, [2.0, 10.5, 4.0])
        _add_prose_block(doc, '2. Phần ghi chú cuối văn bản', 8, 'Ghi chú')
        dst_path.parent.mkdir(parents=True, exist_ok=True)
        tmp_base = dst_path.with_suffix('.tmp.docx')
        doc.save(tmp_base)
        with zipfile.ZipFile(tmp_base, 'r') as zin:
            parts = {info.filename: zin.read(info.filename) for info in zin.infolist() if not info.is_dir()}
        tmp_base.unlink()
        w_ns = t['docx_package'].W
        doc_root = etree.fromstring(parts['word/document.xml'])
        merged_ns = dict(doc_root.nsmap)
        merged_ns.update({
            'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
            'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
            'w15': 'http://schemas.microsoft.com/office/word/2012/wordml',
            'wps': 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
        })
        rebuilt = etree.Element(doc_root.tag, attrib=dict(doc_root.attrib), nsmap=merged_ns)
        rebuilt.set('{http://schemas.openxmlformats.org/markup-compatibility/2006}Ignorable', 'w14 w15')
        for ch in list(doc_root):
            rebuilt.append(ch)
        alt_p = etree.fromstring(
            '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
            ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
            ' xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"'
            ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"'
            ' w14:paraId="55AA77BB">'
            '<w:r><mc:AlternateContent><mc:Choice Requires="wps">'
            '<w:t>ChoiceBlockWord365</w:t></mc:Choice><mc:Fallback><w:t>FallbackBlockWord365</w:t></mc:Fallback>'
            '</mc:AlternateContent></w:r></w:p>'
        )
        rebuilt.find(f'{{{w_ns}}}body').insert(0, alt_p)
        parts['word/document.xml'] = etree.tostring(rebuilt, encoding='UTF-8', xml_declaration=True, standalone=True)
        with zipfile.ZipFile(dst_path, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
            zout.writestr('word/', b'')
            for fname, payload in parts.items():
                zout.writestr(fname, payload)
        return

    elif generator_name == 'signed_docx_src':
        build.setup_document(doc, 'academic')
        build.add_body(doc, 'Tài liệu đã ký số OOXML (không được phép chỉnh sửa nội dung).')
        tbl = build.add_table(doc, ['TT', 'Mục', 'Giá trị'], [['1', 'Hợp đồng', '100.000.000']], profile='academic')
        _set_col_widths_cm(tbl, [2.0, 9.5, 5.0])
        dst_path.parent.mkdir(parents=True, exist_ok=True)
        tmp_base = dst_path.with_suffix('.tmp.docx')
        doc.save(tmp_base)
        with zipfile.ZipFile(tmp_base, 'r') as zin:
            parts = {info.filename: zin.read(info.filename) for info in zin.infolist() if not info.is_dir()}
        tmp_base.unlink()
        parts['_xmlsignatures/sig1.xml'] = (
            b'<?xml version="1.0" encoding="UTF-8"?>'
            b'<Signature xmlns="http://www.w3.org/2000/09/xmldsig#"><SignedInfo/></Signature>'
        )
        with zipfile.ZipFile(dst_path, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
            for fname, payload in parts.items():
                zout.writestr(fname, payload)
        return

    elif generator_name == 'header_footer_table_src':
        build.setup_document(doc, 'academic')
        sec = doc.sections[0]
        sec.header.paragraphs[0].text = 'CƠ QUAN CHỦ QUẢN - BIỂU MẪU CHUẨN NỘI BỘ'
        sec.footer.paragraphs[0].text = 'Tài liệu lưu hành nội bộ - Năm 2026'
        _add_prose_block(doc, '1. Nội dung kèm bảng có ô phụ lồng nhau', 14, 'HeaderFooter')
        tbl = build.add_table(
            doc,
            ['TT', 'Chi tiết nhiệm vụ', 'Trạng thái'],
            [['1', 'Gói công việc chính', 'Hoàn thành'], ['2', 'Gói công việc mở rộng', 'Đang chạy']],
            profile='academic',
        )
        _set_col_widths_cm(tbl, [2.0, 10.5, 4.0])
        sub_tbl = tbl.rows[1].cells[1].add_table(rows=1, cols=2)
        sub_tbl.cell(0, 0).text = 'Nhánh A'
        sub_tbl.cell(0, 1).text = 'Đạt'
        _add_prose_block(doc, '2. Kết luận cuối trang', 8, 'Kết')

    elif generator_name == 'long_nd30_full_report_src':
        build.setup_document(doc, 'administrative')
        _add_nd30_header(doc)
        _add_prose_block(doc, 'I. TÌNH HÌNH THỰC HIỆN NHIỆM VỤ', 20, 'NĐ30 Mục I')
        t_main = build.add_table(
            doc,
            ['TT', 'Hạng mục triển khai', 'Đơn vị', 'Kinh phí (đồng)'],
            [
                ['1', 'Chuẩn hóa quy trình tiếp nhận văn bản', 'Văn phòng', '120.000.000'],
                ['2', 'Vận hành hệ thống kiểm thử tự động', 'Trung tâm CNTT', '185.000.000'],
            ],
            profile='administrative',
        )
        _set_col_widths_cm(t_main, [1.5, 7.5, 3.0, 4.5])
        for r_idx, row in enumerate(t_main.rows):
            for cell in row.cells:
                for p in cell.paragraphs:
                    for run in p.runs:
                        run.font.size = Pt(14)
                        run.bold = (r_idx == 0)
        _add_prose_block(doc, 'II. SƠ ĐỒ KIẾN TRÚC XỬ LÝ VĂN BẢN', 20, 'NĐ30 Mục II')
        p_cap = build.add_body(doc, 'Hình 1: Sơ đồ quy trình xử lý văn bản điện tử')
        p_cap.alignment = WD_ALIGN_PARAGRAPH.LEFT
        with tempfile.TemporaryDirectory(prefix='vndocs-nd30-img-') as asset_dir:
            diag = _make_png(Path(asset_dir) / 'workflow_diag.png', 800, 400, (18, 82, 148))
            doc.add_picture(str(diag), width=Cm(18.0))
        _add_prose_block(doc, 'III. KIẾN NGHỊ VÀ TỔ CHỨC THỰC HIỆN', 10, 'NĐ30 Mục III')
        _add_nd30_signature(doc)

    else:
        raise ValueError(f'Unknown generator: {generator_name}')

    dst_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(dst_path)


def generate_corpus(output_dir):
    t = _tools()
    cases_cfg = json.loads(CASES_PATH.read_text(encoding='utf-8'))
    out_root = Path(output_dir)
    out_root.mkdir(parents=True, exist_ok=True)
    manifest = []

    for case in cases_cfg['cases']:
        cid = case['id']
        dst = out_root / f'{cid}.docx'
        spec = case['source_spec']
        if spec['kind'] == 'starter':
            src_tpl = ASSETS_TEMPLATES / spec['template']
            if not src_tpl.is_file():
                raise FileNotFoundError(f'Starter template missing: {src_tpl}')
            shutil.copy2(src_tpl, dst)
        elif spec['kind'] == 'generated':
            _build_generated_doc(spec['generator'], dst)
        else:
            raise ValueError(f'Unsupported source_spec kind: {spec}')

        screenshot_path = None
        if case.get('has_screenshot'):
            shot_ord = case.get('screenshot_table_ordinal', case.get('grading', {}).get('target_ordinal', 1))
            shot_out = out_root / f'{cid}.screenshot.png'
            generated_shot = _generate_table_screenshot(dst, shot_ord, shot_out)
            if generated_shot is not None:
                screenshot_path = str(generated_shot)

        with zipfile.ZipFile(dst, 'r') as zf:
            names = zf.namelist()
            doc_xml_bytes = len(zf.read('word/document.xml')) if 'word/document.xml' in names else 0
            styles_xml_bytes = len(zf.read('word/styles.xml')) if 'word/styles.xml' in names else 0
            is_signed = any(n.startswith('_xmlsignatures/') for n in names)

        if is_signed:
            tbl_cnt, img_cnt, para_cnt = 1, 0, 1
        else:
            mapped = t['map_docx'].map_document(dst)
            tbl_cnt = mapped['summary']['tables']
            img_cnt = mapped['summary'].get('images', 0)
            para_cnt = mapped['summary']['paragraphs']

        manifest.append({
            'id': cid,
            'path': str(dst),
            'screenshot_path': screenshot_path,
            'signed': is_signed,
            'source_sha256': _sha256(dst),
            'length_tier': case['length_tier'],
            'instruction_type': case['instruction_type'],
            'expected_outcome': case['expected_outcome'],
            'doc_xml_bytes': doc_xml_bytes,
            'styles_xml_bytes': styles_xml_bytes,
            'table_count': tbl_cnt,
            'image_count': img_cnt,
            'paragraph_count': para_cnt,
            'target_ordinal': case['grading'].get('target_ordinal'),
        })

    (out_root / 'corpus_manifest.json').write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2) + '\n',
        encoding='utf-8',
    )
    return manifest


def measure_payload(corpus_dir=None):
    t = _tools()
    tmp_ctx = None
    if corpus_dir is None:
        tmp_ctx = tempfile.TemporaryDirectory(prefix='vndocs-payload-')
        cdir = Path(tmp_ctx.name)
    else:
        cdir = Path(corpus_dir)

    try:
        manifest_file = cdir / 'corpus_manifest.json'
        if not manifest_file.is_file():
            manifest = generate_corpus(cdir)
        else:
            manifest = json.loads(manifest_file.read_text(encoding='utf-8'))

        per_case = []
        for entry in manifest:
            raw_xml_b = entry['doc_xml_bytes']
            raw_both_b = entry['doc_xml_bytes'] + entry['styles_xml_bytes']
            if entry.get('signed'):
                continue
            src = Path(entry['path'])
            mapped = t['map_docx'].map_document(src)
            map_md = t['map_docx'].format_markdown(mapped)
            map_json = json.dumps(mapped, ensure_ascii=False)
            raw_xml_b = entry['doc_xml_bytes']
            raw_both_b = entry['doc_xml_bytes'] + entry['styles_xml_bytes']
            md_b = len(map_md.encode('utf-8'))
            json_b = len(map_json.encode('utf-8'))
            per_case.append({
                'id': entry['id'],
                'length_tier': entry['length_tier'],
                'raw_doc_xml_bytes': raw_xml_b,
                'raw_doc_plus_styles_bytes': raw_both_b,
                'map_markdown_bytes': md_b,
                'map_json_bytes': json_b,
                'est_tok_raw_doc_xml': round(raw_xml_b / 4),
                'est_tok_map_markdown': round(md_b / 4),
                'est_tok_map_json': round(json_b / 4),
                'md_reduction_vs_raw_pct': round((1.0 - md_b / max(1, raw_xml_b)) * 100.0, 1),
            })

        by_tier = {}
        for tier in ('short', 'medium', 'long'):
            sub = [r for r in per_case if r['length_tier'] == tier]
            if sub:
                raw_avg = statistics.mean(r['raw_doc_xml_bytes'] for r in sub)
                md_avg = statistics.mean(r['map_markdown_bytes'] for r in sub)
                by_tier[tier] = {
                    'count': len(sub),
                    'mean_raw_doc_xml_bytes': round(raw_avg),
                    'mean_map_markdown_bytes': round(md_avg),
                    'mean_est_tok_raw_xml': round(raw_avg / 4),
                    'mean_est_tok_map_md': round(md_avg / 4),
                    'mean_reduction_pct': round((1.0 - md_avg / max(1, raw_avg)) * 100.0, 1),
                }

        skill_md_path = REPO_ROOT / 'SKILL.md'
        tooling_md_path = REPO_ROOT / 'references' / 'tooling.md'
        compact_b = len(COMPACT_LAYOUT_SKILL_BLOCK.encode('utf-8'))
        full_skill_b = skill_md_path.stat().st_size if skill_md_path.is_file() else 0
        tooling_b = tooling_md_path.stat().st_size if tooling_md_path.is_file() else 0

        return {
            'token_estimation_note': 'Heuristic bytes/4 proxy for static payload only; live A/B measures exact model tokens from transcript.jsonl.',
            'instruction_overhead': {
                'compact_layout_skill_bytes': compact_b,
                'compact_layout_skill_est_tok': round(compact_b / 4),
                'full_skill_md_bytes': full_skill_b,
                'full_skill_md_est_tok': round(full_skill_b / 4),
                'full_skill_plus_tooling_bytes': full_skill_b + tooling_b,
                'full_skill_plus_tooling_est_tok': round((full_skill_b + tooling_b) / 4),
            },
            'by_length_tier': by_tier,
            'cases': per_case,
        }
    finally:
        if tmp_ctx is not None:
            tmp_ctx.cleanup()


def _body_paragraphs_signature(doc_root):
    t = _tools()
    w = t['docx_package'].W
    body = doc_root.find(f'{{{w}}}body')
    sigs = []
    for p in body:
        if etree.QName(p).localname != 'p':
            continue
        txt = t['map_docx'].normalize_nfc(''.join(node.text or '' for node in p.iter(f'{{{w}}}t')))
        p_pr = p.find(f'{{{w}}}pPr')
        p_style = p_pr.find(f'{{{w}}}pStyle') if p_pr is not None else None
        style_val = p_style.get(f'{{{w}}}val') if p_style is not None else None
        sigs.append((style_val, txt))
    return sigs


def _check_grid_span_tcw(tbl_node, w_ns):
    """Verify every <w:tcW> equals sum of spanned <w:gridCol w:w> within 30 dxa."""
    t = _tools()
    grid_cols = [
        t['map_docx']._int_attr(c, 'w')
        for c in tbl_node.xpath('./w:tblGrid/w:gridCol', namespaces={'w': w_ns})
    ]
    if not grid_cols or any(g is None or g <= 0 for g in grid_cols):
        return False
    for tr in tbl_node.findall(f'{{{w_ns}}}tr'):
        tr_pr = tr.find(f'{{{w_ns}}}trPr')
        cursor = t['map_docx']._int_attr(tr_pr.find(f'{{{w_ns}}}gridBefore') if tr_pr is not None else None, 'val', 0)
        for tc in tr.findall(f'{{{w_ns}}}tc'):
            tc_pr = tc.find(f'{{{w_ns}}}tcPr')
            span = max(1, t['map_docx']._int_attr(tc_pr.find(f'{{{w_ns}}}gridSpan') if tc_pr is not None else None, 'val', 1))
            tc_w = tc_pr.find(f'{{{w_ns}}}tcW') if tc_pr is not None else None
            tc_w_val = t['map_docx']._int_attr(tc_w, 'w')
            tc_w_type = tc_w.get(f'{{{w_ns}}}type') if tc_w is not None else None
            if tc_w_val is None or tc_w_type != 'dxa':
                return False
            expected_sum = sum(grid_cols[cursor:cursor + span])
            if abs(tc_w_val - expected_sum) > 30:
                return False
            cursor += span
    return True


def _check_target_tbl_ecma_order(tbl_node, w_ns):
    t = _tools()
    violations = []
    for idx, ch in enumerate(tbl_node):
        if isinstance(ch.tag, str) and etree.QName(ch).localname == 'tblPr' and idx > 0:
            violations.append('tblPr not first child of tbl')
    tbl_pr = tbl_node.find(f'{{{w_ns}}}tblPr')
    if tbl_pr is not None:
        violations.extend(v['message'] for v in t['ooxml_order'].check_element_order(tbl_pr))
        tbl_mar = tbl_pr.find(f'{{{w_ns}}}tblCellMar')
        if tbl_mar is not None:
            violations.extend(v['message'] for v in t['ooxml_order'].check_element_order(tbl_mar))
        tbl_borders = tbl_pr.find(f'{{{w_ns}}}tblBorders')
        if tbl_borders is not None:
            violations.extend(v['message'] for v in t['ooxml_order'].check_element_order(tbl_borders))
    for tr in tbl_node.findall(f'{{{w_ns}}}tr'):
        for idx, ch in enumerate(tr):
            if isinstance(ch.tag, str) and etree.QName(ch).localname == 'trPr' and idx > 0:
                violations.append('trPr not first child of tr')
        tr_pr = tr.find(f'{{{w_ns}}}trPr')
        if tr_pr is not None:
            violations.extend(v['message'] for v in t['ooxml_order'].check_element_order(tr_pr))
    for tc in tbl_node.xpath('./w:tr/w:tc', namespaces={'w': w_ns}):
        for idx, ch in enumerate(tc):
            if isinstance(ch.tag, str) and etree.QName(ch).localname == 'tcPr' and idx > 0:
                violations.append('tcPr not first child of tc')
        tc_pr = tc.find(f'{{{w_ns}}}tcPr')
        if tc_pr is not None:
            violations.extend(v['message'] for v in t['ooxml_order'].check_element_order(tc_pr))
            tc_mar = tc_pr.find(f'{{{w_ns}}}tcMar')
            if tc_mar is not None:
                violations.extend(v['message'] for v in t['ooxml_order'].check_element_order(tc_mar))
            tc_borders = tc_pr.find(f'{{{w_ns}}}tcBorders')
            if tc_borders is not None:
                violations.extend(v['message'] for v in t['ooxml_order'].check_element_order(tc_borders))
    for run in tbl_node.xpath('.//w:r', namespaces={'w': w_ns}):
        for idx, ch in enumerate(run):
            if isinstance(ch.tag, str) and etree.QName(ch).localname == 'rPr' and idx > 0:
                violations.append('rPr not first child of r')
    return violations


def _check_ppr_ecma_order(p_node, w_ns):
    t = _tools()
    violations = []
    for idx, ch in enumerate(p_node):
        if isinstance(ch.tag, str) and etree.QName(ch).localname == 'pPr' and idx > 0:
            violations.append('pPr not first child of p')
    p_pr = p_node.find(f'{{{w_ns}}}pPr')
    if p_pr is not None:
        violations.extend(v['message'] for v in t['ooxml_order'].check_element_order(p_pr))
    return violations


def _extract_run_size_hp(run, p_node, w_ns):
    t = _tools()
    r_pr = run.find(f'{{{w_ns}}}rPr')
    sz = r_pr.find(f'{{{w_ns}}}sz') if r_pr is not None else None
    val = t['map_docx']._int_attr(sz, 'val')
    if val is not None:
        return val
    p_pr = p_node.find(f'{{{w_ns}}}pPr')
    p_rpr = p_pr.find(f'{{{w_ns}}}rPr') if p_pr is not None else None
    p_sz = p_rpr.find(f'{{{w_ns}}}sz') if p_rpr is not None else None
    return t['map_docx']._int_attr(p_sz, 'val')


def _is_run_bold(run, p_node, w_ns):
    for holder in (run.find(f'{{{w_ns}}}rPr'), p_node.find(f'{{{w_ns}}}pPr/{{{w_ns}}}rPr')):
        if holder is None:
            continue
        b_el = holder.find(f'{{{w_ns}}}b')
        if b_el is not None:
            val = (b_el.get(f'{{{w_ns}}}val') or '1').lower()
            return val not in ('0', 'false', 'off')
    return False


def _check_table_margins_and_fonts(tbl_node, grading, w_ns):
    t = _tools()
    ok = True
    exp_mar = grading.get('expected_cell_margins_cm')
    if exp_mar:
        tol = grading.get('margin_tolerance_cm', 0.03)
        tbl_mar = tbl_node.find(f'{{{w_ns}}}tblPr/{{{w_ns}}}tblCellMar')
        if tbl_mar is None:
            return False
        aliases = {'top': ('top',), 'bottom': ('bottom',), 'left': ('left', 'start'), 'right': ('right', 'end')}
        for side, target_cm in exp_mar.items():
            node = None
            for cand in aliases[side]:
                node = tbl_mar.find(f'{{{w_ns}}}{cand}')
                if node is not None:
                    break
            dxa = t['map_docx']._int_attr(node, 'w')
            if dxa is None or abs((dxa / t['map_docx'].DXA_PER_CM) - target_cm) > tol:
                ok = False
                break
    if grading.get('require_no_cell_margin_overrides') and tbl_node.xpath('.//w:tcPr/w:tcMar', namespaces={'w': w_ns}):
        ok = False

    rows = tbl_node.findall(f'{{{w_ns}}}tr')
    exp_hdr_pt = grading.get('expected_header_font_pt')
    req_hdr_bold = grading.get('require_header_bold', False)
    if rows and (exp_hdr_pt is not None or req_hdr_bold):
        exp_hp = round(exp_hdr_pt * 2) if exp_hdr_pt is not None else None
        for p_node in rows[0].xpath('./w:tc//w:p', namespaces={'w': w_ns}):
            for run in p_node.findall(f'{{{w_ns}}}r'):
                txt = ''.join(node.text or '' for node in run.findall(f'{{{w_ns}}}t')).strip()
                if not txt:
                    continue
                if exp_hp is not None and _extract_run_size_hp(run, p_node, w_ns) != exp_hp:
                    return False
                if req_hdr_bold and not _is_run_bold(run, p_node, w_ns):
                    return False

    exp_data_pt = grading.get('expected_data_font_pt')
    if exp_data_pt is not None and len(rows) > 1:
        exp_hp = round(exp_data_pt * 2)
        for tr in rows[1:]:
            for p_node in tr.xpath('./w:tc//w:p', namespaces={'w': w_ns}):
                for run in p_node.findall(f'{{{w_ns}}}r'):
                    txt = ''.join(node.text or '' for node in run.findall(f'{{{w_ns}}}t')).strip()
                    if not txt:
                        continue
                    if _extract_run_size_hp(run, p_node, w_ns) != exp_hp:
                        return False

    if grading.get('expected_repeat_header') is not None:
        has_hdr = bool(rows) and (rows[0].find(f'{{{w_ns}}}trPr/{{{w_ns}}}tblHeader') is not None)
        if has_hdr is not grading['expected_repeat_header']:
            return False

    if grading.get('expected_cant_split') is not None:
        all_cs = bool(rows) and all(tr.find(f'{{{w_ns}}}trPr/{{{w_ns}}}cantSplit') is not None for tr in rows)
        if all_cs is not grading['expected_cant_split']:
            return False

    if 'expected_border_style' in grading or 'expected_border_color' in grading:
        borders = tbl_node.find(f'{{{w_ns}}}tblPr/{{{w_ns}}}tblBorders')
        if borders is None:
            return False
        exp_b_style = grading.get('expected_border_style')
        exp_b_color = grading.get('expected_border_color')
        for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            b_el = borders.find(f'{{{w_ns}}}{side}')
            if b_el is None:
                return False
            if exp_b_style and b_el.get(f'{{{w_ns}}}val') != exp_b_style:
                return False
            if exp_b_color and (b_el.get(f'{{{w_ns}}}color') or '').upper() != exp_b_color.upper():
                return False

    if 'expected_header_shading_fill' in grading:
        exp_fill = grading['expected_header_shading_fill'].upper()
        if not rows:
            return False
        for tc in rows[0].findall(f'{{{w_ns}}}tc'):
            shd = tc.find(f'{{{w_ns}}}tcPr/{{{w_ns}}}shd')
            if shd is None or (shd.get(f'{{{w_ns}}}fill') or '').upper() != exp_fill:
                return False
        for tr in rows[1:]:
            for tc in tr.findall(f'{{{w_ns}}}tc'):
                shd = tc.find(f'{{{w_ns}}}tcPr/{{{w_ns}}}shd')
                if shd is not None and (shd.get(f'{{{w_ns}}}fill') or '').upper() == exp_fill:
                    return False
    return ok


def _check_style_xml(styles_root, grading, t, w_ns):
    if styles_root is None:
        return False, ['styles.xml missing']
    st_node = t['layout_edit']._find_style_node(styles_root, grading.get('target_style_name'))
    if st_node is None:
        return False, ['target style not found']
    order_viols = [v['message'] for v in t['ooxml_order'].check_element_order(st_node)]

    ok = True
    r_pr = st_node.find(f'{{{w_ns}}}rPr')
    p_pr = st_node.find(f'{{{w_ns}}}pPr')
    if 'expected_style_font' in grading:
        rf = r_pr.find(f'{{{w_ns}}}rFonts') if r_pr is not None else None
        want_f = grading['expected_style_font']
        if rf is None or rf.get(f'{{{w_ns}}}ascii') != want_f or rf.get(f'{{{w_ns}}}hAnsi') != want_f:
            ok = False
    if 'expected_style_size_pt' in grading:
        sz = r_pr.find(f'{{{w_ns}}}sz') if r_pr is not None else None
        if t['map_docx']._int_attr(sz, 'val') != round(grading['expected_style_size_pt'] * 2):
            ok = False
    if 'expected_style_color' in grading:
        col = r_pr.find(f'{{{w_ns}}}color') if r_pr is not None else None
        if col is None or (col.get(f'{{{w_ns}}}val') or '').upper() != grading['expected_style_color'].upper():
            ok = False
    sp = p_pr.find(f'{{{w_ns}}}spacing') if p_pr is not None else None
    if 'expected_style_line_rule' in grading:
        if sp is None or sp.get(f'{{{w_ns}}}lineRule') != grading['expected_style_line_rule']:
            ok = False
    if 'expected_style_line_pt' in grading:
        line_val = t['map_docx']._int_attr(sp, 'line') if sp is not None else None
        if line_val is None or abs((line_val / 20.0) - grading['expected_style_line_pt']) > 0.1:
            ok = False
    if 'expected_style_space_after_pt' in grading:
        after_val = t['map_docx']._int_attr(sp, 'after') if sp is not None else None
        if after_val is None or abs((after_val / 20.0) - grading['expected_style_space_after_pt']) > 0.1:
            ok = False
    return ok, order_viols


def _check_merged_runs(before_p, after_p, grading, t, w_ns):
    runs_after = after_p.findall(f'{{{w_ns}}}r')
    if len(runs_after) > grading.get('max_run_count_after', 100):
        return False
    if t['docx_package'].text_of(before_p) != t['docx_package'].text_of(after_p):
        return False

    def _char_bold_seq(p_node):
        seq = []
        for run in p_node.findall(f'{{{w_ns}}}r'):
            r_pr = run.find(f'{{{w_ns}}}rPr')
            b_el = r_pr.find(f'{{{w_ns}}}b') if r_pr is not None else None
            is_bold = b_el is not None and (b_el.get(f'{{{w_ns}}}val') or '1').lower() not in ('0', 'false', 'off')
            txt = ''.join(node.text or '' for node in run.findall(f'{{{w_ns}}}t'))
            for ch in txt:
                seq.append((ch, is_bold))
        return seq

    return _char_bold_seq(before_p) == _char_bold_seq(after_p)


def _check_section_layout_meta(before_map, after_map, grading):
    sec_idx = grading.get('target_section_index', 1)
    after_sec = next((s for s in after_map['sections'] if s['section_index'] == sec_idx), None)
    if after_sec is None:
        return False
    if 'expected_orientation' in grading and after_sec['orientation'] != grading['expected_orientation']:
        return False
    for side, exp_cm in grading.get('expected_section_margins_cm', {}).items():
        if abs(after_sec['margins_cm'].get(side, -999) - exp_cm) > 0.05:
            return False
    if grading.get('keep_section1_portrait'):
        b1, a1 = before_map['sections'][0], after_map['sections'][0]
        if a1['orientation'] != 'portrait' or b1['margins_cm'] != a1['margins_cm']:
            return False
    return True


def _check_image_xml(after_root, img_before, img_after, grading, w_ns):
    t = _tools()
    body = after_root.find(f'{{{w_ns}}}body')
    body_children = list(body)
    if img_after['body_index'] >= len(body_children):
        return False
    p_node = body_children[img_after['body_index']]
    drawings = p_node.xpath('.//w:drawing', namespaces={'w': w_ns})
    if img_after['drawing_index_in_p'] >= len(drawings):
        return False
    d_node = drawings[img_after['drawing_index_in_p']]
    box = d_node.find(f'{{{t["map_docx"].WP}}}inline')
    if box is None:
        box = d_node.find(f'{{{t["map_docx"].WP}}}anchor')
    if box is None:
        return False
    extent = box.find(f'{{{t["map_docx"].WP}}}extent')
    if extent is None:
        return False
    try:
        wp_cx = int(extent.get('cx') or '0')
        wp_cy = int(extent.get('cy') or '0')
    except ValueError:
        return False
    a_exts = box.xpath('.//a:xfrm/a:ext', namespaces={'a': t['map_docx'].A})
    if not a_exts:
        return False
    for aext in a_exts:
        try:
            acx = int(aext.get('cx') or '0')
            acy = int(aext.get('cy') or '0')
        except ValueError:
            return False
        if abs(acx - wp_cx) > 1000 or abs(acy - wp_cy) > 1000:
            return False

    w_ok = True
    if 'expected_image_width_cm' in grading:
        tol = grading.get('image_width_tolerance_cm', 0.25)
        w_ok = abs(img_after['width_cm'] - grading['expected_image_width_cm']) <= tol

    ar_ok = True
    ar_b = img_before.get('aspect_ratio')
    ar_a = img_after.get('aspect_ratio')
    if ar_b and ar_a and 'max_aspect_ratio_delta_pct' in grading:
        delta_pct = abs(ar_a - ar_b) / ar_b * 100.0
        ar_ok = delta_pct <= grading['max_aspect_ratio_delta_pct']

    align_ok = True
    if 'expected_image_alignment' in grading:
        align_ok = img_after.get('alignment') == grading['expected_image_alignment']

    return bool(w_ok and ar_ok and align_ok)


def _check_paragraph_xml(p_node, grading, w_ns):
    t = _tools()
    p_pr = p_node.find(f'{{{w_ns}}}pPr')
    if p_pr is None:
        return False

    if 'expected_para_align' in grading:
        jc = p_pr.find(f'{{{w_ns}}}jc')
        if jc is None or jc.get(f'{{{w_ns}}}val') != grading['expected_para_align']:
            return False

    if 'expected_first_line_cm_range' in grading:
        low_cm, high_cm = grading['expected_first_line_cm_range']
        ind = p_pr.find(f'{{{w_ns}}}ind')
        if ind is None or ind.get(f'{{{w_ns}}}hanging') is not None:
            return False
        fl_dxa = t['map_docx']._int_attr(ind, 'firstLine')
        if fl_dxa is None:
            return False
        fl_cm = fl_dxa / t['map_docx'].DXA_PER_CM
        if not (low_cm <= fl_cm <= high_cm):
            return False

    sp = p_pr.find(f'{{{w_ns}}}spacing')
    if 'expected_line_rule' in grading or 'expected_line_pt_range' in grading or 'min_para_spacing_pt' in grading:
        if sp is None:
            return False
    if 'expected_line_rule' in grading:
        if sp.get(f'{{{w_ns}}}lineRule') != grading['expected_line_rule']:
            return False
    if 'expected_line_pt_range' in grading:
        low_pt, high_pt = grading['expected_line_pt_range']
        line_val = t['map_docx']._int_attr(sp, 'line')
        if line_val is None or not (low_pt <= (line_val / 20.0) <= high_pt):
            return False
    if 'min_para_spacing_pt' in grading:
        after_val = t['map_docx']._int_attr(sp, 'after', 0)
        if (after_val / 20.0) < (grading['min_para_spacing_pt'] - 0.05):
            return False

    if 'expected_keep_with_next' in grading:
        kn = p_pr.find(f'{{{w_ns}}}keepNext')
        is_kn = kn is not None and (kn.get(f'{{{w_ns}}}val') or '1').lower() not in ('0', 'false', 'off')
        if is_kn is not grading['expected_keep_with_next']:
            return False

    return True


def _non_target_ppr_signatures(doc_root, excluded_body_indices, w_ns):
    body = doc_root.find(f'{{{w_ns}}}body')
    sigs = []
    for b_idx, child in enumerate(body):
        if etree.QName(child).localname != 'p' or b_idx in excluded_body_indices:
            continue
        p_pr = child.find(f'{{{w_ns}}}pPr')
        sigs.append(etree.tostring(p_pr) if p_pr is not None else b'')
    return sigs


def grade_case(case, src_docx, submission_case_dir):
    """Arm-agnostic deterministic evaluation returning Binary Hit, Checkpoint Score, and telemetry."""
    t = _tools()
    sub_dir = Path(submission_case_dir)
    cid = case['id']
    res_file = sub_dir / f'{cid}.result.json'
    dst_docx = sub_dir / f'{cid}.docx'

    res_json = None
    if res_file.is_file():
        try:
            res_json = json.loads(res_file.read_text(encoding='utf-8-sig'))
        except (json.JSONDecodeError, OSError):
            res_json = None

    expected_outcome = case['expected_outcome']
    grading = case['grading']
    src_path = Path(src_docx)
    src_sha = _sha256(src_path)

    # Branch 1: Expected Refusal (e.g. T04 NĐ30 layout-only, T11 <0.51 cm col, T17 ambiguous crop)
    if expected_outcome == 'refuse':
        status_refused = isinstance(res_json, dict) and res_json.get('status') == 'refused' and bool(str(res_json.get('reason', '')).strip())
        no_mutated_output = (not dst_docx.is_file()) or (_sha256(dst_docx) == src_sha)
        checkpoints = {
            'refusal_status': status_refused,
            'no_mutated_output': no_mutated_output,
        }
        hit = 1 if all(checkpoints.values()) else 0
        return {
            'case_id': cid,
            'length_tier': case['length_tier'],
            'instruction_type': case['instruction_type'],
            'expected_outcome': expected_outcome,
            'hit': hit,
            'checkpoint_score': round(sum(1 for v in checkpoints.values() if v) / len(checkpoints), 4),
            'checkpoints': checkpoints,
            'telemetry': {
                'xml_ntpr': 1.0 if no_mutated_output else 0.0,
                'zip_parts_byte_identical': 1.0 if no_mutated_output else 0.0,
                'semantic_non_target_preserved': no_mutated_output,
            },
            'failure_reasons': [k for k, ok in checkpoints.items() if not ok],
        }

    # Branch 2: Expected Edit (supports table, image, paragraph, merge_runs, style, section, paginate, and multi targets)
    target_kind = grading.get('target_kind', 'table')
    has_tbl_style_checks = any(
        k in grading
        for k in (
            'expected_cell_margins_cm',
            'expected_data_font_pt',
            'expected_header_font_pt',
            'require_header_bold',
            'require_no_cell_margin_overrides',
            'expected_repeat_header',
            'expected_cant_split',
            'expected_border_style',
            'expected_border_color',
            'expected_header_shading_fill',
        )
    )
    checkpoints = {'valid_output_pkg': False}
    if target_kind in ('table', 'multi'):
        checkpoints['ooxml_target_width'] = False
        if has_tbl_style_checks:
            checkpoints['ooxml_table_margins_and_fonts'] = False
        checkpoints['pdf_geometry_match'] = False
    if target_kind in ('image', 'multi'):
        checkpoints['ooxml_target_image'] = False
        if target_kind == 'image':
            checkpoints['pdf_geometry_match'] = False
        else:
            checkpoints['pdf_image_geometry_match'] = False
    if target_kind in ('paragraph', 'multi'):
        checkpoints['ooxml_target_paragraph'] = False
    if target_kind == 'merge_runs':
        checkpoints['ooxml_runs_merged'] = False
    if target_kind == 'style':
        checkpoints['ooxml_target_style'] = False
        if grading.get('require_document_xml_unchanged'):
            checkpoints['document_xml_unchanged'] = False
    if target_kind == 'section':
        checkpoints['ooxml_target_section'] = False
    if target_kind == 'paginate':
        checkpoints['page_count_target_met'] = False
        if grading.get('require_no_orphan_signature'):
            checkpoints['no_orphan_signature'] = False
        if grading.get('validate_profile'):
            checkpoints['profile_validation_pass'] = False
    checkpoints['ooxml_order_clean'] = False
    checkpoints['semantic_non_target_preserved'] = False
    if grading.get('require_track_changes_tags'):
        checkpoints['ooxml_track_changes_present'] = False
    if grading.get('require_namespace_audit_pass') or grading.get('require_mc_alternate_preserved'):
        checkpoints['ooxml_namespaces_preserved'] = False
    if grading.get('require_header_footer_byte_identical'):
        checkpoints['header_footer_byte_identical'] = False
    if grading.get('expected_pagination_ripple') is not None:
        checkpoints['pagination_ripple_reported'] = False

    telemetry = {
        'xml_ntpr': None,
        'zip_parts_byte_identical': None,
        'render_vs_ooxml_error_cm': None,
        'pdf_width_before_cm': None,
        'pdf_width_after_cm': None,
        'ooxml_width_after_cm': None,
        'ooxml_image_width_after_cm': None,
        'pdf_image_width_after_cm': None,
        'overflow_right_mm_after': None,
        'alignment_after': None,
        'pagination_ripple_detected': None,
        'submitted_status': res_json.get('status') if isinstance(res_json, dict) else None,
    }

    if not dst_docx.is_file():
        return {
            'case_id': cid,
            'length_tier': case['length_tier'],
            'instruction_type': case['instruction_type'],
            'expected_outcome': expected_outcome,
            'hit': 0,
            'checkpoint_score': 0.0,
            'checkpoints': checkpoints,
            'telemetry': telemetry,
            'failure_reasons': ['missing_output_docx'],
        }

    # Package + ZIP sanity + XML parsing
    has_zip_dirs = False
    try:
        with zipfile.ZipFile(dst_docx, 'r') as zf:
            has_zip_dirs = any(info.is_dir() for info in zf.infolist())
        _, before_parts = t['docx_package'].read_package(src_path)
        _, after_parts = t['docx_package'].read_package(dst_docx)
        parser = etree.XMLParser(resolve_entities=False, no_network=True)
        before_root = etree.fromstring(before_parts['word/document.xml'], parser)
        after_root = etree.fromstring(after_parts['word/document.xml'], parser)
        before_map = t['map_docx'].map_document(src_path, include_row_texts=True)
        after_map = t['map_docx'].map_document(dst_docx, include_row_texts=True)
    except Exception as exc:
        return {
            'case_id': cid,
            'length_tier': case['length_tier'],
            'instruction_type': case['instruction_type'],
            'expected_outcome': expected_outcome,
            'hit': 0,
            'checkpoint_score': 0.0,
            'checkpoints': checkpoints,
            'telemetry': {**telemetry, 'error': str(exc)},
            'failure_reasons': ['invalid_docx_package'],
        }

    other_parts = [k for k in before_parts if k != 'word/document.xml']
    same_other = sum(1 for k in other_parts if before_parts.get(k) == after_parts.get(k))
    telemetry['zip_parts_byte_identical'] = round(same_other / len(other_parts), 4) if other_parts else 1.0

    w_ns = t['docx_package'].W
    target_tbl_ord = grading.get('target_ordinal')
    target_img_ord = grading.get('target_image_ordinal')

    if len(after_map['tables']) != len(before_map['tables']):
        return {
            'case_id': cid,
            'length_tier': case['length_tier'],
            'instruction_type': case['instruction_type'],
            'expected_outcome': expected_outcome,
            'hit': 0,
            'checkpoint_score': 0.0,
            'checkpoints': checkpoints,
            'telemetry': telemetry,
            'failure_reasons': ['table_count_mismatch'],
        }
    if target_tbl_ord is not None and len(before_map['tables']) < target_tbl_ord:
        return {
            'case_id': cid,
            'length_tier': case['length_tier'],
            'instruction_type': case['instruction_type'],
            'expected_outcome': expected_outcome,
            'hit': 0,
            'checkpoint_score': 0.0,
            'checkpoints': checkpoints,
            'telemetry': telemetry,
            'failure_reasons': ['table_count_mismatch'],
        }
    if len(after_map.get('images', [])) != len(before_map.get('images', [])):
        return {
            'case_id': cid,
            'length_tier': case['length_tier'],
            'instruction_type': case['instruction_type'],
            'expected_outcome': expected_outcome,
            'hit': 0,
            'checkpoint_score': 0.0,
            'checkpoints': checkpoints,
            'telemetry': telemetry,
            'failure_reasons': ['image_count_mismatch'],
        }
    if target_img_ord is not None and len(before_map.get('images', [])) < target_img_ord:
        return {
            'case_id': cid,
            'length_tier': case['length_tier'],
            'instruction_type': case['instruction_type'],
            'expected_outcome': expected_outcome,
            'hit': 0,
            'checkpoint_score': 0.0,
            'checkpoints': checkpoints,
            'telemetry': telemetry,
            'failure_reasons': ['image_count_mismatch'],
        }

    before_body_ps = [
        (b_idx, ch) for b_idx, ch in enumerate(before_root.find(f'{{{w_ns}}}body'))
        if etree.QName(ch).localname == 'p'
    ]
    after_body_ps = [
        (b_idx, ch) for b_idx, ch in enumerate(after_root.find(f'{{{w_ns}}}body'))
        if etree.QName(ch).localname == 'p'
    ]
    target_para_list_idxs = []
    if target_kind in ('paragraph', 'merge_runs', 'multi'):
        if len(before_body_ps) != len(after_body_ps):
            return {
                'case_id': cid,
                'length_tier': case['length_tier'],
                'instruction_type': case['instruction_type'],
                'expected_outcome': expected_outcome,
                'hit': 0,
                'checkpoint_score': 0.0,
                'checkpoints': checkpoints,
                'telemetry': telemetry,
                'failure_reasons': ['paragraph_count_mismatch'],
            }
        for ord_num in grading.get('target_paragraph_ordinals', []):
            if 1 <= ord_num <= len(before_body_ps):
                target_para_list_idxs.append(ord_num - 1)
        if grading.get('target_paragraph_text_match'):
            want_txt = t['map_docx'].normalize_nfc(grading['target_paragraph_text_match'])
            for idx_p, (_, p_el) in enumerate(before_body_ps):
                txt = t['map_docx'].normalize_nfc(t['docx_package'].text_of(p_el))
                if txt == want_txt and idx_p not in target_para_list_idxs:
                    target_para_list_idxs.append(idx_p)
        if not target_para_list_idxs:
            return {
                'case_id': cid,
                'length_tier': case['length_tier'],
                'instruction_type': case['instruction_type'],
                'expected_outcome': expected_outcome,
                'hit': 0,
                'checkpoint_score': 0.0,
                'checkpoints': checkpoints,
                'telemetry': telemetry,
                'failure_reasons': ['target_paragraph_not_found'],
            }

    allowed_body_indices = set()
    t_before = t_after = target_tbl_node = None
    if target_tbl_ord is not None:
        t_before = before_map['tables'][target_tbl_ord - 1]
        t_after = after_map['tables'][target_tbl_ord - 1]
        after_tbl_nodes = [
            ch for ch in after_root.find(f'{{{w_ns}}}body')
            if etree.QName(ch).localname == 'tbl'
        ]
        target_tbl_node = after_tbl_nodes[target_tbl_ord - 1]
        allowed_body_indices.add(t_before['body_index'])

    img_before = img_after = None
    if target_img_ord is not None:
        img_before = before_map['images'][target_img_ord - 1]
        img_after = after_map['images'][target_img_ord - 1]
        allowed_body_indices.add(img_before['body_index'])

    target_para_body_indices = {before_body_ps[i][0] for i in target_para_list_idxs}
    allowed_body_indices |= target_para_body_indices

    target_sect_after_node = None
    if target_kind == 'section':
        before_secs_xml = t['layout_edit']._locate_sections_xml(before_root.find(f'{{{w_ns}}}body'))
        after_secs_xml = t['layout_edit']._locate_sections_xml(after_root.find(f'{{{w_ns}}}body'))
        sec_idx = grading.get('target_section_index', 1)
        b_match = next((s for s in before_secs_xml if s[0] == sec_idx), None)
        a_match = next((s for s in after_secs_xml if s[0] == sec_idx), None)
        if b_match is not None:
            allowed_body_indices.add(b_match[2])
        if a_match is not None:
            target_sect_after_node = a_match[1]

    if target_kind == 'paginate':
        # Layout compaction may touch non-layout tables, images, non-heading paragraphs, and sectPr only.
        for bt in before_map['tables']:
            if not bt.get('is_layout_table'):
                allowed_body_indices.add(bt['body_index'])
        for bi in before_map.get('images', []):
            allowed_body_indices.add(bi['body_index'])
        before_body_children = list(before_root.find(f'{{{w_ns}}}body'))
        for bp in before_map['paragraphs']:
            p_node = before_body_children[bp['body_index']]
            if not t['auto_paginate']._is_heading_or_title_node(p_node, bp.get('style')):
                allowed_body_indices.add(bp['body_index'])
        for _, _, sec_b_idx, _ in t['layout_edit']._locate_sections_xml(before_root.find(f'{{{w_ns}}}body')):
            allowed_body_indices.add(sec_b_idx)

    # Non-target XML invariance telemetry
    try:
        ntpr = t['ooxml_order'].check_non_target_preservation(
            before_root, after_root, allowed_body_indices,
        )
        telemetry['xml_ntpr'] = ntpr['xml_ntpr']
    except Exception:
        telemetry['xml_ntpr'] = 0.0

    # Semantic non-target preservation (fair to both python-docx and vndocs)
    para_ok = _body_paragraphs_signature(before_root) == _body_paragraphs_signature(after_root)
    non_target_tables_ok = True
    for bt, at in zip(before_map['tables'], after_map['tables']):
        if bt['ordinal'] == target_tbl_ord:
            continue
        if target_kind == 'paginate' and not bt.get('is_layout_table'):
            if bt['row_texts_nfc'] != at['row_texts_nfc']:
                non_target_tables_ok = False
                break
            continue
        if (
            bt['rows'] != at['rows']
            or bt['grid_cols'] != at['grid_cols']
            or bt['alignment'] != at['alignment']
            or bt['row_texts_nfc'] != at['row_texts_nfc']
            or abs(bt['total_width_cm'] - at['total_width_cm']) > 0.1
            or any(abs(a - b) > 0.1 for a, b in zip(bt['grid_widths_cm'], at['grid_widths_cm']))
        ):
            non_target_tables_ok = False
            break
    if t_before is not None and t_before['row_texts_nfc'] != t_after['row_texts_nfc']:
        non_target_tables_ok = False

    non_target_images_ok = True
    for bi, ai in zip(before_map.get('images', []), after_map.get('images', [])):
        if bi['ordinal'] == target_img_ord or target_kind == 'paginate':
            continue
        if (
            abs(bi['width_cm'] - ai['width_cm']) > 0.05
            or abs(bi['height_cm'] - ai['height_cm']) > 0.05
            or bi['alignment'] != ai['alignment']
        ):
            non_target_images_ok = False
            break

    non_target_ppr_ok = True
    if target_kind in ('paragraph', 'merge_runs', 'multi'):
        excluded_p_indices = set(target_para_body_indices)
        if img_before is not None:
            excluded_p_indices.add(img_before['body_index'])
        non_target_ppr_ok = (
            _non_target_ppr_signatures(before_root, excluded_p_indices, w_ns)
            == _non_target_ppr_signatures(after_root, excluded_p_indices, w_ns)
        )

    checkpoints['semantic_non_target_preserved'] = bool(
        para_ok and non_target_tables_ok and non_target_images_ok and non_target_ppr_ok
    )

    # Target ECMA child ordering check
    order_violations = []
    if target_tbl_node is not None:
        order_violations.extend(_check_target_tbl_ecma_order(target_tbl_node, w_ns))
    if img_after is not None:
        img_p_node = list(after_root.find(f'{{{w_ns}}}body'))[img_after['body_index']]
        order_violations.extend(_check_ppr_ecma_order(img_p_node, w_ns))
    for i in target_para_list_idxs:
        order_violations.extend(_check_ppr_ecma_order(after_body_ps[i][1], w_ns))
    if target_sect_after_node is not None:
        order_violations.extend(v['message'] for v in t['ooxml_order'].check_element_order(target_sect_after_node))
    if target_kind == 'paginate':
        from collections import Counter
        after_body_children = list(after_root.find(f'{{{w_ns}}}body'))
        for b_idx in sorted(allowed_body_indices):
            if b_idx >= len(after_body_children):
                continue
            ch = after_body_children[b_idx]
            ln = etree.QName(ch).localname
            if ln == 'tbl':
                order_violations.extend(_check_target_tbl_ecma_order(ch, w_ns))
            elif ln == 'p':
                order_violations.extend(_check_ppr_ecma_order(ch, w_ns))
                p_sect = ch.find(f'{{{w_ns}}}pPr/{{{w_ns}}}sectPr')
                if p_sect is not None:
                    order_violations.extend(v['message'] for v in t['ooxml_order'].check_element_order(p_sect))
            elif ln == 'sectPr':
                order_violations.extend(v['message'] for v in t['ooxml_order'].check_element_order(ch))
        before_doc_cnt = Counter(
            (v['container'], v['rule'], v['message'])
            for v in t['ooxml_order'].check_element_order(before_root)
        )
        after_doc_cnt = Counter(
            (v['container'], v['rule'], v['message'])
            for v in t['ooxml_order'].check_element_order(after_root)
        )
        for key, cnt in after_doc_cnt.items():
            if cnt > before_doc_cnt.get(key, 0):
                order_violations.append(f'new order violation: {key[2]}')

    after_styles_root = (
        etree.fromstring(after_parts['word/styles.xml'], parser)
        if 'word/styles.xml' in after_parts
        else None
    )
    if target_kind == 'style':
        style_ok, style_viols = _check_style_xml(after_styles_root, grading, t, w_ns)
        order_violations.extend(style_viols)
        checkpoints['ooxml_target_style'] = bool(style_ok)
        if 'document_xml_unchanged' in checkpoints:
            checkpoints['document_xml_unchanged'] = bool(
                before_parts.get('word/document.xml') == after_parts.get('word/document.xml')
            )

    checkpoints['ooxml_order_clean'] = len(order_violations) == 0

    if target_kind == 'section':
        checkpoints['ooxml_target_section'] = _check_section_layout_meta(before_map, after_map, grading)

    if target_kind == 'merge_runs':
        checkpoints['ooxml_runs_merged'] = all(
            _check_merged_runs(before_body_ps[i][1], after_body_ps[i][1], grading, t, w_ns)
            for i in target_para_list_idxs
        )

    if 'ooxml_track_changes_present' in checkpoints:
        has_grid_chg = target_tbl_node is not None and (
            target_tbl_node.find(f'{{{w_ns}}}tblGrid/{{{w_ns}}}tblGridChange') is not None
        )
        has_tblpr_chg = target_tbl_node is not None and (
            target_tbl_node.find(f'{{{w_ns}}}tblPr/{{{w_ns}}}tblPrChange') is not None
        )
        checkpoints['ooxml_track_changes_present'] = bool(has_grid_chg and has_tblpr_chg)

    if 'ooxml_namespaces_preserved' in checkpoints:
        mc_ns = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
        ns_audit = t['docx_package'].audit_namespace_preservation(
            before_parts['word/document.xml'],
            after_parts['word/document.xml'],
        )
        alt_ok = True
        if grading.get('require_mc_alternate_preserved'):
            before_alts = [etree.tostring(el) for el in before_root.iter(f'{{{mc_ns}}}AlternateContent')]
            after_alts = [etree.tostring(el) for el in after_root.iter(f'{{{mc_ns}}}AlternateContent')]
            alt_ok = len(before_alts) >= 1 and before_alts == after_alts
        checkpoints['ooxml_namespaces_preserved'] = bool(ns_audit.get('preserved') and alt_ok)

    if 'header_footer_byte_identical' in checkpoints:
        hf_keys = [k for k in before_parts if k.startswith(('word/header', 'word/footer')) and k.endswith('.xml')]
        checkpoints['header_footer_byte_identical'] = bool(
            hf_keys and all(before_parts.get(k) == after_parts.get(k) for k in hf_keys)
        )

    # Target OOXML checks per target kind
    if target_kind in ('table', 'multi'):
        telemetry['ooxml_width_after_cm'] = t_after['total_width_cm']
        telemetry['alignment_after'] = t_after['alignment']
        if 'expected_width_min_cm' in grading and 'expected_width_max_cm' in grading:
            width_ok = grading['expected_width_min_cm'] <= t_after['total_width_cm'] <= grading['expected_width_max_cm']
        elif 'expected_total_width_cm' in grading:
            width_ok = abs(t_after['total_width_cm'] - grading['expected_total_width_cm']) <= grading['width_tolerance_cm']
        elif 'expected_width_max_drift_cm' in grading:
            width_ok = abs(t_after['total_width_cm'] - t_before['total_width_cm']) <= grading['expected_width_max_drift_cm']
        else:
            width_ok = True

        cols_ok = True
        if 'expected_col_widths_cm' in grading:
            tol = grading.get('col_tolerance_cm', 0.15)
            cols_ok = (
                len(t_after['grid_widths_cm']) == len(grading['expected_col_widths_cm'])
                and all(abs(a - b) <= tol for a, b in zip(t_after['grid_widths_cm'], grading['expected_col_widths_cm']))
            )
        if 'expected_col_widths_pct' in grading:
            ptol = grading.get('pct_tolerance', 2.5)
            cols_ok = cols_ok and (
                len(t_after['grid_widths_pct']) == len(grading['expected_col_widths_pct'])
                and all(abs(a - b) <= ptol for a, b in zip(t_after['grid_widths_pct'], grading['expected_col_widths_pct']))
            )
        if grading.get('check_grid_span_tcw'):
            cols_ok = cols_ok and _check_grid_span_tcw(target_tbl_node, w_ns)

        align_ooxml_ok = (
            True if 'expected_alignment' not in grading
            else (t_after['alignment'] == grading['expected_alignment'])
        )
        checkpoints['ooxml_target_width'] = bool(width_ok and cols_ok and align_ooxml_ok)
        if has_tbl_style_checks:
            checkpoints['ooxml_table_margins_and_fonts'] = _check_table_margins_and_fonts(target_tbl_node, grading, w_ns)

    if target_kind in ('image', 'multi'):
        telemetry['ooxml_image_width_after_cm'] = img_after['width_cm']
        if target_kind == 'image':
            telemetry['alignment_after'] = img_after['alignment']
        checkpoints['ooxml_target_image'] = _check_image_xml(after_root, img_before, img_after, grading, w_ns)

    if target_kind in ('paragraph', 'multi'):
        checkpoints['ooxml_target_paragraph'] = all(
            _check_paragraph_xml(after_body_ps[i][1], grading, w_ns)
            for i in target_para_list_idxs
        )

    # Render & diff via LibreOffice + PyMuPDF (single render pass reused when multi)
    primary_target = (
        t_before['anchor']
        if target_kind in ('table', 'multi')
        else (img_before['anchor'] if target_kind == 'image' else None)
    )
    with tempfile.TemporaryDirectory(prefix='vndocs-grade-') as grade_tmp:
        diff_res = t['diff_layout'].diff_layout(
            src_path, dst_docx, target=primary_target, output_dir=grade_tmp,
        )
        status_edited_ok = (res_json is None) or (isinstance(res_json, dict) and res_json.get('status') == 'edited')
        checkpoints['valid_output_pkg'] = bool(
            (not has_zip_dirs)
            and diff_res.get('status') == 'ok'
            and status_edited_ok
        )

        if diff_res.get('status') == 'ok':
            if target_kind in ('table', 'multi') and diff_res.get('target_metrics'):
                tm = diff_res['target_metrics']
                telemetry['pdf_width_before_cm'] = tm.get('width_before_cm')
                telemetry['pdf_width_after_cm'] = tm.get('width_after_cm')
                telemetry['render_vs_ooxml_error_cm'] = tm.get('render_vs_ooxml_error_cm')
                telemetry['overflow_right_mm_after'] = tm.get('overflow_right_mm_after')
                telemetry['pagination_ripple_detected'] = diff_res.get('pagination_ripple_detected')

                pdf_w = tm.get('width_after_cm')
                if pdf_w is None or not tm.get('matched_after'):
                    pdf_ok = False
                else:
                    if 'expected_width_min_cm' in grading and 'expected_width_max_cm' in grading:
                        pdf_w_ok = (grading['expected_width_min_cm'] - 0.15) <= pdf_w <= (grading['expected_width_max_cm'] + 0.15)
                    elif 'expected_total_width_cm' in grading:
                        pdf_w_ok = abs(pdf_w - grading['expected_total_width_cm']) <= (grading['width_tolerance_cm'] + 0.15)
                    elif 'expected_width_max_drift_cm' in grading:
                        pdf_w_ok = (
                            tm.get('width_before_cm') is not None
                            and abs(pdf_w - tm['width_before_cm']) <= (grading['expected_width_max_drift_cm'] + 0.15)
                        )
                    else:
                        pdf_w_ok = True
                    err_ok = tm.get('render_vs_ooxml_error_cm') is not None and tm['render_vs_ooxml_error_cm'] <= 0.5
                    ovf_ok = (tm.get('overflow_right_mm_after') or 0.0) <= grading.get('max_overflow_right_mm', 0.5)
                    pdf_align_ok = True
                    if grading.get('expected_alignment') == 'center':
                        after_sec = next(s for s in after_map['sections'] if s['section_index'] == t_after['section_index'])
                        expected_center_cm = after_sec['margins_cm']['left'] + after_sec['printable_width_cm'] / 2.0
                        x0_cm = tm.get('x0_after_cm')
                        pdf_align_ok = (
                            x0_cm is not None
                            and abs((x0_cm + pdf_w / 2.0) - expected_center_cm) <= 0.65
                        )
                    pdf_ok = bool(pdf_w_ok and err_ok and ovf_ok and pdf_align_ok)
                checkpoints['pdf_geometry_match'] = pdf_ok

            if target_kind in ('image', 'multi'):
                img_key = 'pdf_geometry_match' if target_kind == 'image' else 'pdf_image_geometry_match'
                after_pdf = diff_res.get('pdf_after')
                img_rec = next(
                    (rec for rec in diff_res.get('image_diffs', []) if rec['anchor'] == img_before['anchor']),
                    None,
                )
                if img_rec and after_pdf:
                    after_pdf_map = t['map_docx'].map_document(dst_docx, pdf_path=after_pdf)
                    after_img_pdf = after_pdf_map['images'][target_img_ord - 1].get('pdf') or {}
                    pdf_iw = img_rec.get('rendered_width_cm_after')
                    telemetry['pdf_image_width_after_cm'] = pdf_iw
                    if target_kind == 'image':
                        telemetry['pdf_width_before_cm'] = img_rec.get('rendered_width_cm_before')
                        telemetry['pdf_width_after_cm'] = pdf_iw
                        telemetry['overflow_right_mm_after'] = img_rec.get('overflow_right_mm_after')
                    tol_i = grading.get('image_width_tolerance_cm', 0.25) + 0.25
                    iw_ok = (
                        pdf_iw is not None
                        and abs(pdf_iw - grading['expected_image_width_cm']) <= tol_i
                    )
                    ovf_i_ok = (img_rec.get('overflow_right_mm_after') or 0.0) <= grading.get('max_overflow_right_mm', 0.5)
                    align_i_ok = True
                    if grading.get('expected_image_alignment') == 'center':
                        bbox = after_img_pdf.get('bbox_pt')
                        after_sec = next(s for s in after_map['sections'] if s['section_index'] == img_after['section_index'])
                        expected_center_cm = after_sec['margins_cm']['left'] + after_sec['printable_width_cm'] / 2.0
                        if bbox and len(bbox) == 4:
                            img_center_cm = ((bbox[0] + bbox[2]) / 2.0) / t['map_docx'].PT_PER_CM
                            align_i_ok = abs(img_center_cm - expected_center_cm) <= 0.7
                        else:
                            align_i_ok = False
                    checkpoints[img_key] = bool(iw_ok and ovf_i_ok and align_i_ok)

            if target_kind == 'paginate':
                after_pdf = diff_res.get('pdf_after')
                if after_pdf and Path(after_pdf).is_file():
                    import pymupdf
                    with pymupdf.open(after_pdf) as pdoc:
                        pages_after = len(pdoc)
                    telemetry['pages_after'] = pages_after
                    checkpoints['page_count_target_met'] = pages_after <= grading['expected_max_pages']
                    if 'no_orphan_signature' in checkpoints:
                        checkpoints['no_orphan_signature'] = not bool(
                            t['auto_paginate'].detect_orphan_signature(dst_docx, after_pdf)
                        )
                if 'profile_validation_pass' in checkpoints:
                    val_res = t['validate_docx'].inspect_doc(dst_docx, profile=grading['validate_profile'])
                    checkpoints['profile_validation_pass'] = val_res.get('status') == 'pass'

            if 'pagination_ripple_reported' in checkpoints:
                checkpoints['pagination_ripple_reported'] = bool(
                    isinstance(res_json, dict)
                    and res_json.get('pagination_ripple') is grading['expected_pagination_ripple']
                    and diff_res.get('pagination_ripple_detected') is True
                )

    hit = 1 if all(checkpoints.values()) else 0
    return {
        'case_id': cid,
        'length_tier': case['length_tier'],
        'instruction_type': case['instruction_type'],
        'expected_outcome': expected_outcome,
        'hit': hit,
        'checkpoint_score': round(sum(1 for v in checkpoints.values() if v) / len(checkpoints), 4),
        'checkpoints': checkpoints,
        'telemetry': telemetry,
        'failure_reasons': [k for k, ok in checkpoints.items() if not ok],
    }


def grade_directory(submission_dir, corpus_dir):
    cases_cfg = json.loads(CASES_PATH.read_text(encoding='utf-8'))
    cdir = Path(corpus_dir)
    sdir = Path(submission_dir)
    results = []
    for case in cases_cfg['cases']:
        cid = case['id']
        src_docx = cdir / f'{cid}.docx'
        sub_case_dir = sdir / cid
        results.append(grade_case(case, src_docx, sub_case_dir))
    return results


def parse_transcript(transcript_path):
    """Extract exact token counts, tool calls, elapsed latency, and skill/contamination markers."""
    tpath = Path(transcript_path)
    full_path = tpath.parent / 'transcript_full.jsonl'
    read_path = full_path if full_path.is_file() else tpath
    if not read_path.is_file():
        return {'missing': True}

    fresh_in = 0
    cache_read = 0
    out_tok = 0
    llm_turns = 0
    tool_calls_total = 0
    tools_by_name = {}
    timestamps = []
    models = set()
    vndocs_markers = (
        'vndocs', 'layout_edit', 'layout-edit', 'map_docx', 'diff_layout',
        'ooxml_order', 'docx_package', 'render_docx', 'annotate-marks',
    )
    saw_skill_marker = False
    matched_markers = set()

    for line in read_path.read_text(encoding='utf-8').splitlines():
        if not line.strip():
            continue
        step = json.loads(line)
        ts = step.get('created_at')
        if ts:
            try:
                timestamps.append(datetime.fromisoformat(ts.replace('Z', '+00:00')))
            except ValueError:
                pass
        if step.get('type') == 'PLANNER_RESPONSE':
            llm_turns += 1
            if step.get('model'):
                models.add(step['model'])
            fresh_in += int(step.get('input_tokens') or 0)
            cache_read += int(step.get('cache_read_tokens') or 0)
            out_tok += int(step.get('output_tokens') or 0)
            for tc in step.get('tool_calls') or []:
                tool_calls_total += 1
                nm = tc.get('name', 'unknown')
                tools_by_name[nm] = tools_by_name.get(nm, 0) + 1
                blob = json.dumps(tc.get('args', {}), ensure_ascii=False).lower()
                for mk in vndocs_markers:
                    if mk in blob:
                        saw_skill_marker = True
                        matched_markers.add(mk)

    elapsed = round((max(timestamps) - min(timestamps)).total_seconds(), 2) if len(timestamps) >= 2 else 0.0
    total_tok = fresh_in + cache_read + out_tok
    return {
        'missing': False,
        'transcript_file': str(read_path),
        'models': sorted(models),
        'llm_turns': llm_turns,
        'tool_calls_total': tool_calls_total,
        'tools_by_name': tools_by_name,
        'fresh_input_tokens': fresh_in,
        'cache_read_tokens': cache_read,
        'output_tokens': out_tok,
        'fresh_plus_output_tokens': fresh_in + out_tok,
        'total_tokens': total_tok,
        'elapsed_seconds': elapsed,
        'saw_vndocs_marker': saw_skill_marker,
        'matched_markers': sorted(matched_markers),
    }


def _stats(vals):
    clean = [float(v) for v in vals if v is not None]
    if not clean:
        return {'mean': 0.0, 'stddev': 0.0, 'min': 0.0, 'max': 0.0}
    mean = sum(clean) / len(clean)
    std = math.sqrt(sum((x - mean) ** 2 for x in clean) / (len(clean) - 1)) if len(clean) > 1 else 0.0
    return {
        'mean': round(mean, 4),
        'stddev': round(std, 4),
        'min': round(min(clean), 4),
        'max': round(max(clean), 4),
    }


def _summarize_arm(rows):
    hits = [r['hit'] for r in rows]
    cps = [r['checkpoint_score'] for r in rows]
    tot_tok = [r['transcript'].get('total_tokens', 0) for r in rows]
    fresh_out = [r['transcript'].get('fresh_plus_output_tokens', 0) for r in rows]
    fresh_in = [r['transcript'].get('fresh_input_tokens', 0) for r in rows]
    cache_r = [r['transcript'].get('cache_read_tokens', 0) for r in rows]
    out_t = [r['transcript'].get('output_tokens', 0) for r in rows]
    calls = [r['transcript'].get('tool_calls_total', 0) for r in rows]
    secs = [r['transcript'].get('elapsed_seconds', 0.0) for r in rows]
    invoked = [1 if r['transcript'].get('saw_vndocs_marker') else 0 for r in rows]
    sem_ntp = [1 if r['telemetry'].get('semantic_non_target_preserved', r['checkpoints'].get('semantic_non_target_preserved', r['checkpoints'].get('no_mutated_output'))) else 0 for r in rows]
    xml_ntprs = [r['telemetry'].get('xml_ntpr') for r in rows if r['telemetry'].get('xml_ntpr') is not None]
    zip_inv = [r['telemetry'].get('zip_parts_byte_identical') for r in rows if r['telemetry'].get('zip_parts_byte_identical') is not None]

    n_hits = sum(hits)
    return {
        'n': len(rows),
        'hits': n_hits,
        'hit_rate': _stats(hits),
        'checkpoint_score': _stats(cps),
        'total_tokens': _stats(tot_tok),
        'fresh_plus_output_tokens': _stats(fresh_out),
        'fresh_input_tokens': _stats(fresh_in),
        'cache_read_tokens': _stats(cache_r),
        'output_tokens': _stats(out_t),
        'tool_calls': _stats(calls),
        'time_seconds': _stats(secs),
        'tokens_per_hit_total': round(sum(tot_tok) / n_hits, 1) if n_hits else None,
        'tokens_per_hit_fresh_plus_out': round(sum(fresh_out) / n_hits, 1) if n_hits else None,
        'skill_invocation_rate': round(sum(invoked) / len(invoked), 4) if invoked else 0.0,
        'semantic_ntp_rate': round(sum(sem_ntp) / len(sem_ntp), 4) if sem_ntp else 0.0,
        'xml_ntpr_mean': round(sum(xml_ntprs) / len(xml_ntprs), 4) if xml_ntprs else None,
        'zip_byte_invariance_mean': round(sum(zip_inv) / len(zip_inv), 4) if zip_inv else None,
    }


def analyze_ab(ab_root_dir):
    root = Path(ab_root_dir)
    manifest = json.loads((root / 'runs_manifest.json').read_text(encoding='utf-8'))
    cases_by_id = {c['id']: c for c in json.loads(CASES_PATH.read_text(encoding='utf-8'))['cases']}

    graded_runs = []
    for entry in manifest['runs']:
        case = cases_by_id[entry['case_id']]
        grade = grade_case(case, entry['src_docx'], entry['work_dir'])
        tr = parse_transcript(entry['transcript_path'])
        row = {
            **grade,
            'model_tier': entry['model_tier'],
            'arm': entry['arm'],
            'conversation_id': entry['conversation_id'],
            'arm_a_contaminated': bool(entry['arm'] == 'without_skill' and tr.get('saw_vndocs_marker')),
            'transcript': tr,
        }
        graded_runs.append(row)

    # Group per model_tier (flash, inherit, etc.)
    model_tiers = sorted({r['model_tier'] for r in graded_runs})
    summary = {}
    for mt in model_tiers:
        sub_mt = [r for r in graded_runs if r['model_tier'] == mt]
        without_rows = [r for r in sub_mt if r['arm'] == 'without_skill']
        with_rows = [r for r in sub_mt if r['arm'] == 'with_skill']
        s_wo = _summarize_arm(without_rows)
        s_wi = _summarize_arm(with_rows)
        hit_wo = s_wo['hit_rate']['mean']
        hit_wi = s_wi['hit_rate']['mean']
        norm_g = round((hit_wi - hit_wo) / (1.0 - hit_wo), 4) if hit_wo < 1.0 else (0.0 if hit_wi == hit_wo else None)

        by_tier = {}
        for lt in ('short', 'medium', 'long'):
            lt_wo = [r for r in without_rows if r['length_tier'] == lt]
            lt_wi = [r for r in with_rows if r['length_tier'] == lt]
            if lt_wo or lt_wi:
                by_tier[lt] = {
                    'without_skill': _summarize_arm(lt_wo),
                    'with_skill': _summarize_arm(lt_wi),
                }

        by_itype = {}
        for itype in ('descriptive', 'inferential'):
            it_wo = [r for r in without_rows if r['instruction_type'] == itype]
            it_wi = [r for r in with_rows if r['instruction_type'] == itype]
            if it_wo or it_wi:
                by_itype[itype] = {
                    'without_skill': _summarize_arm(it_wo),
                    'with_skill': _summarize_arm(it_wi),
                }

        summary[mt] = {
            'without_skill': s_wo,
            'with_skill': s_wi,
            'delta': {
                'hit_rate_pp': round((hit_wi - hit_wo) * 100.0, 2),
                'normalized_gain_g': norm_g,
                'checkpoint_score_pp': round((s_wi['checkpoint_score']['mean'] - s_wo['checkpoint_score']['mean']) * 100.0, 2),
                'total_tokens_delta': round(s_wi['total_tokens']['mean'] - s_wo['total_tokens']['mean'], 1),
                'fresh_plus_out_tokens_delta': round(s_wi['fresh_plus_output_tokens']['mean'] - s_wo['fresh_plus_output_tokens']['mean'], 1),
                'output_tokens_delta': round(s_wi['output_tokens']['mean'] - s_wo['output_tokens']['mean'], 1),
                'tool_calls_delta': round(s_wi['tool_calls']['mean'] - s_wo['tool_calls']['mean'], 2),
                'time_seconds_delta': round(s_wi['time_seconds']['mean'] - s_wo['time_seconds']['mean'], 2),
            },
            'by_length_tier': by_tier,
            'by_instruction_type': by_itype,
        }

    payload_stats = measure_payload(root / 'corpus' if (root / 'corpus').is_dir() else None)
    benchmark_out = {
        'benchmark_name': 'vietnamese-docx-layout-edit-v1',
        'generated_at': datetime.now().astimezone().isoformat(),
        'run_summary': summary,
        'static_payload': payload_stats,
        'contaminated_without_skill_runs': [
            {'model_tier': r['model_tier'], 'case_id': r['case_id'], 'markers': r['transcript'].get('matched_markers')}
            for r in graded_runs if r['arm_a_contaminated']
        ],
        'runs': graded_runs,
    }
    (root / 'benchmark.json').write_text(json.dumps(benchmark_out, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    _write_live_ab_report(root / 'LIVE-AB-REPORT.md', benchmark_out)
    return benchmark_out


def _write_live_ab_report(report_path, data):
    lines = [
        '# Live A/B Benchmark Report: Layout Editing Skill (`with_skill` vs `without_skill`)',
        '',
        f"- Generated: `{data['generated_at']}`",
        f"- Arm A contamination count: **{len(data['contaminated_without_skill_runs'])}**",
        '',
        '## 1. Tổng hợp theo Model',
        '| Model | Arm | Hit Rate | Checkpoint Mean | Normalized Gain $g$ | Skill Invocation | Sem. NTP | XML NTPR | Mean Fresh+Out Tok | Mean Total Tok | Tok/Hit (Total) | Mean Calls | Mean Time (s) |',
        '|---|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|',
    ]
    for mt, info in data['run_summary'].items():
        g_val = info['delta']['normalized_gain_g']
        g_str = f'{g_val * 100:.1f}%' if g_val is not None else 'N/A'
        for arm in ('without_skill', 'with_skill'):
            s = info[arm]
            g_col = g_str if arm == 'with_skill' else '-'
            ntpr_s = f"{s['xml_ntpr_mean']:.3f}" if s['xml_ntpr_mean'] is not None else 'N/A'
            tph_s = f"{s['tokens_per_hit_total']:.0f}" if s['tokens_per_hit_total'] is not None else 'INF'
            lines.append(
                f"| `{mt}` | `{arm}` | {s['hits']}/{s['n']} ({s['hit_rate']['mean']*100:.1f}%) | "
                f"{s['checkpoint_score']['mean']*100:.1f}% | {g_col} | {s['skill_invocation_rate']*100:.1f}% | "
                f"{s['semantic_ntp_rate']*100:.1f}% | {ntpr_s} | "
                f"{s['fresh_plus_output_tokens']['mean']:.0f} | {s['total_tokens']['mean']:.0f} | "
                f"{tph_s} | {s['tool_calls']['mean']:.1f} | {s['time_seconds']['mean']:.1f} |"
            )

    lines.extend(['', '## 2. Phân tầng theo Độ dài Tài liệu & Loại Lệnh'])
    for mt, info in data['run_summary'].items():
        lines.append(f"### Model: `{mt}`")
        lines.append('| Phân nhóm | Without Hit | With Hit | Delta Hit | Without Mean Total Tok | With Mean Total Tok | Without Mean Fresh+Out | With Mean Fresh+Out |')
        lines.append('|---|---:|---:|---:|---:|---:|---:|---:|')
        for grp_label, grp in [('length:' + k, v) for k, v in info['by_length_tier'].items()] + [('type:' + k, v) for k, v in info['by_instruction_type'].items()]:
            wo, wi = grp['without_skill'], grp['with_skill']
            d_pp = (wi['hit_rate']['mean'] - wo['hit_rate']['mean']) * 100.0
            lines.append(
                f"| `{grp_label}` | {wo['hits']}/{wo['n']} ({wo['hit_rate']['mean']*100:.0f}%) | "
                f"{wi['hits']}/{wi['n']} ({wi['hit_rate']['mean']*100:.0f}%) | {d_pp:+.1f} pp | "
                f"{wo['total_tokens']['mean']:.0f} | {wi['total_tokens']['mean']:.0f} | "
                f"{wo['fresh_plus_output_tokens']['mean']:.0f} | {wi['fresh_plus_output_tokens']['mean']:.0f} |"
            )

    lines.extend([
        '',
        '## 3. Chi tiết từng ca (All Runs)',
        '| Model | Case | Tier | Type | Arm | Hit | CP | Invoked | Total Tok | Fresh+Out | Calls | Time (s) | Lý do trượt (nếu có) |',
        '|---|---|---|---|---|---:|---:|---:|---:|---:|---:|---:|---|',
    ])
    for r in data['runs']:
        tr = r['transcript']
        fails = ', '.join(r['failure_reasons']) or '-'
        lines.append(
            f"| `{r['model_tier']}` | `{r['case_id']}` | {r['length_tier']} | {r['instruction_type']} | "
            f"`{r['arm']}` | {r['hit']} | {r['checkpoint_score']*100:.0f}% | "
            f"{'Y' if tr.get('saw_vndocs_marker') else 'N'} | {tr.get('total_tokens', 0)} | "
            f"{tr.get('fresh_plus_output_tokens', 0)} | {tr.get('tool_calls_total', 0)} | "
            f"{tr.get('elapsed_seconds', 0.0):.1f} | {fails} |"
        )

    sp = data['static_payload']
    lines.extend([
        '',
        '## 4. Payload Tĩnh (`word/document.xml` vs `vndocs map`)',
        f"- Compact skill block (~550-650 tok target): **{sp['instruction_overhead']['compact_layout_skill_bytes']} bytes** (~{sp['instruction_overhead']['compact_layout_skill_est_tok']} est tok)",
        f"- Full `SKILL.md`: **{sp['instruction_overhead']['full_skill_md_bytes']} bytes** (~{sp['instruction_overhead']['full_skill_md_est_tok']} est tok)",
        f"- Full `SKILL.md` + `tooling.md`: **{sp['instruction_overhead']['full_skill_plus_tooling_bytes']} bytes** (~{sp['instruction_overhead']['full_skill_plus_tooling_est_tok']} est tok)",
        '',
        '| Tier | Cases | Mean `document.xml` | Mean `map --format markdown` | Giảm payload tĩnh |',
        '|---|---:|---:|---:|---:|',
    ])
    for tier, ts in sp['by_length_tier'].items():
        lines.append(
            f"| {tier} | {ts['count']} | {ts['mean_raw_doc_xml_bytes']} B (~{ts['mean_est_tok_raw_xml']} tok) | "
            f"{ts['mean_map_markdown_bytes']} B (~{ts['mean_est_tok_map_md']} tok) | **{ts['mean_reduction_pct']}%** |"
        )

    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text('\n'.join(lines) + '\n', encoding='utf-8')


def _execute_oracle_case(case, src_docx, sub_case_dir):
    """Run deterministic vndocs oracle action for a single benchmark case T01..T30."""
    t = _tools()
    cid = case['id']
    src_path = Path(src_docx)
    sub_dir = Path(sub_case_dir)
    sub_dir.mkdir(parents=True, exist_ok=True)
    dst_docx = sub_dir / f'{cid}.docx'
    res_file = sub_dir / f'{cid}.result.json'

    def _refuse(reason):
        res_file.write_text(
            json.dumps({'status': 'refused', 'reason': reason}, ensure_ascii=False, indent=2) + '\n',
            encoding='utf-8',
        )
        return {'mode': 'refuse', 'reason': reason}

    try:
        mapping = t['map_docx'].map_document(src_path, include_row_texts=True)
    except ValueError as exc:
        if 'signature' in str(exc).lower():
            return _refuse(str(exc))
        raise

    tbls = mapping['tables']
    data_tbls = [tb for tb in tbls if not tb['is_layout_table']]
    imgs = mapping.get('images', [])
    paras = mapping['paragraphs']
    kwargs = {}
    ops = None

    if cid == 'T01_starter_scale_center':
        a = tbls[0]['anchor']
        ops = [{'op': 'table.scale_width', 'target': a, 'factor': 0.8}, {'op': 'table.set_alignment', 'target': a, 'align': 'center'}]
    elif cid == 'T02_starter_minutes_cols':
        ops = [{'op': 'table.set_col_widths', 'target': tbls[0]['anchor'], 'widths': ['20%', '50%', '30%']}]
    elif cid == 'T03_nd30_budget_not_header':
        a = data_tbls[0]['anchor']
        ops = [{'op': 'table.scale_width', 'target': a, 'factor': 0.8}, {'op': 'table.set_alignment', 'target': a, 'align': 'center'}]
        kwargs['profile'] = 'administrative'
    elif cid == 'T04_nd30_layout_only_refusal':
        if not data_tbls:
            return _refuse('Chỉ có bảng bố cục NĐ30 không viền, không có bảng dữ liệu.')
    elif cid == 'T05_long_multi_similar_tables':
        a = tbls[2]['anchor']
        ops = [{'op': 'table.scale_width', 'target': a, 'factor': 0.75}, {'op': 'table.set_alignment', 'target': a, 'align': 'center'}]
    elif cid == 'T06_overflow_wide_table':
        a = tbls[0]['anchor']
        ops = [{'op': 'table.scale_width', 'target': a, 'target_width_pct': 95.0}, {'op': 'table.set_alignment', 'target': a, 'align': 'center'}]
    elif cid == 'T07_merged_gridspan_vmerge_med':
        a = tbls[1]['anchor']
        ops = [{'op': 'table.set_col_widths', 'target': a, 'widths': ['2.5cm', '6.0cm', '2.5cm', '3.0cm']}, {'op': 'table.set_alignment', 'target': a, 'align': 'center'}]
    elif cid == 'T08_long_landscape_multisection':
        land_secs = {s['section_index'] for s in mapping['sections'] if s['orientation'] == 'landscape'}
        a = next(tb['anchor'] for tb in tbls if tb['section_index'] in land_secs)
        ops = [{'op': 'table.scale_width', 'target': a, 'target_width_pct': 80.0}, {'op': 'table.set_alignment', 'target': a, 'align': 'center'}]
    elif cid == 'T09_long_multipage_table':
        a = tbls[1]['anchor']
        ops = [{'op': 'table.scale_width', 'target': a, 'factor': 0.8}, {'op': 'table.set_alignment', 'target': a, 'align': 'center'}]
    elif cid == 'T10_nfd_vietnamese_med':
        a = tbls[0]['anchor']
        ops = [{'op': 'table.scale_width', 'target': a, 'factor': 0.8}, {'op': 'table.set_alignment', 'target': a, 'align': 'center'}]
    elif cid == 'T11_impossible_narrow_col_refusal':
        try:
            t['layout_edit'].apply_layout_plan(
                src_path, dst_docx, [{'op': 'table.set_col_widths', 'target': tbls[0]['anchor'], 'widths': ['0.3cm', '5cm', '5cm', '3cm']}],
            )
        except t['layout_edit'].LayoutEditError as exc:
            return _refuse(exc.message)
    elif cid == 'T12_pagination_ripple_awareness':
        a = tbls[0]['anchor']
        ops = [{'op': 'table.scale_width', 'target': a, 'factor': 0.55}, {'op': 'table.set_alignment', 'target': a, 'align': 'center'}]
    elif cid == 'T13_screenshot_cropped_similar_tables':
        a = tbls[1]['anchor']
        ops = [{'op': 'table.scale_width', 'target': a, 'factor': 0.8}, {'op': 'table.set_alignment', 'target': a, 'align': 'center'}]
    elif cid == 'T14_table_cell_margins_and_font_nd30':
        a = data_tbls[0]['anchor']
        ops = [
            {'op': 'table.set_cell_margins', 'target': a, 'top_cm': 0.12, 'bottom_cm': 0.12, 'left_cm': 0.18, 'right_cm': 0.18, 'clear_cell_overrides': True},
            {'op': 'table.set_font_size', 'target': a, 'size_pt': 12.0, 'preserve_header_size': True},
        ]
        kwargs['profile'] = 'administrative'
    elif cid == 'T15_image_scale_aspect_ratio':
        ops = [{'op': 'image.scale', 'target': imgs[0]['anchor'], 'width_cm': 12.0, 'align': 'center'}]
    elif cid == 'T16_party_hd05_paragraph_format':
        ops = [{
            'op': 'paragraph.set_format',
            'targets': [paras[3]['anchor'], paras[4]['anchor'], paras[5]['anchor']],
            'align': 'both',
            'first_line_indent_cm': 1.0,
            'line_spacing_exact_pt': 20.0,
            'space_after_pt': 6.0,
        }]
        kwargs['profile'] = 'party'
    elif cid == 'T17_screenshot_ambiguous_identical_tables_refusal':
        if len(tbls) >= 2 and tbls[0]['row_texts_nfc'] == tbls[1]['row_texts_nfc']:
            return _refuse('Hai bảng có cấu trúc và nội dung giống hệt nhau; ảnh crop không phân biệt được.')
    elif cid == 'T18_combined_multi_op_long_doc':
        a = tbls[1]['anchor']
        cap_txt = t['map_docx'].normalize_nfc(case['grading']['target_paragraph_text_match'])
        cap_a = next(p['anchor'] for p in paras if t['map_docx'].normalize_nfc(p['text_preview']) == cap_txt[:90])
        ops = [
            {'op': 'table.scale_width', 'target': a, 'factor': 0.8},
            {'op': 'table.set_alignment', 'target': a, 'align': 'center'},
            {'op': 'table.set_font_size', 'target': a, 'size_pt': 12.0, 'preserve_header_size': True},
            {'op': 'image.scale', 'target': imgs[0]['anchor'], 'width_cm': 11.5, 'align': 'center'},
            {'op': 'paragraph.set_format', 'target': cap_a, 'align': 'center', 'keep_with_next': True},
        ]
    elif cid == 'T19_table_pagination_repeat_header_cantsplit':
        ops = [{'op': 'table.set_pagination', 'target': tbls[1]['anchor'], 'repeat_header': True, 'header_rows': 1, 'cant_split': True}]
    elif cid == 'T20_table_borders_and_header_shading':
        a = tbls[0]['anchor']
        ops = [
            {'op': 'table.set_borders', 'target': a, 'style': 'single', 'size_eighth_pt': 6, 'color': '1F4E79'},
            {'op': 'table.set_shading', 'target': a, 'fill': 'D9E2F3', 'scope': 'header'},
        ]
    elif cid == 'T21_style_update_normal_global':
        ops = [{
            'op': 'style.update',
            'style_name': 'Normal',
            'font': 'Times New Roman',
            'size_pt': 14.0,
            'color': '112233',
            'line_spacing_exact_pt': 20.0,
            'space_after_pt': 6.0,
        }]
    elif cid == 'T22_section_landscape_appendix_layout':
        ops = [{
            'op': 'section.set_layout',
            'section_index': 2,
            'orientation': 'landscape',
            'page_size': 'A4',
            'margins_cm': {'top': 2.0, 'bottom': 2.0, 'left': 2.5, 'right': 2.0},
        }]
    elif cid == 'T23_paragraph_merge_fragmented_runs':
        ops = [{'op': 'paragraph.merge_runs', 'target': paras[0]['anchor']}]
    elif cid == 'T24_track_changes_table_scale':
        a = tbls[0]['anchor']
        ops = [{'op': 'table.scale_width', 'target': a, 'factor': 0.8}, {'op': 'table.set_alignment', 'target': a, 'align': 'center'}]
        kwargs['track_changes'] = True
        kwargs['author'] = 'vndocs'
    elif cid == 'T25_auto_paginate_orphan_nd30_signature':
        pag_res = t['auto_paginate'].auto_paginate(
            src_path, dst_docx, fix_orphan_signature=True, profile='administrative',
        )
        res_file.write_text(json.dumps({'status': 'edited', 'paginate': pag_res}, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
        return {'mode': 'paginate', 'report': pag_res}
    elif cid == 'T26_auto_paginate_target_2pages_report':
        pag_res = t['auto_paginate'].auto_paginate(
            src_path, dst_docx, target_pages=2, profile='academic',
        )
        res_file.write_text(json.dumps({'status': 'edited', 'paginate': pag_res}, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
        return {'mode': 'paginate', 'report': pag_res}
    elif cid == 'T27_complex_word365_namespaces_mc_alternate':
        ops = [{'op': 'table.scale_width', 'target': tbls[0]['anchor'], 'factor': 0.85}]
    elif cid == 'T29_nested_header_footer_preservation':
        a = tbls[0]['anchor']
        ops = [{'op': 'table.scale_width', 'target': a, 'factor': 0.85}, {'op': 'table.set_alignment', 'target': a, 'align': 'center'}]
    elif cid == 'T30_long_full_pipeline_nd30_report':
        a = data_tbls[0]['anchor']
        cap_txt = t['map_docx'].normalize_nfc(case['grading']['target_paragraph_text_match'])
        cap_a = next(p['anchor'] for p in paras if t['map_docx'].normalize_nfc(p['text_preview']) == cap_txt[:90])
        ops = [
            {'op': 'table.scale_width', 'target': a, 'factor': 0.85},
            {'op': 'table.set_alignment', 'target': a, 'align': 'center'},
            {'op': 'table.set_pagination', 'target': a, 'repeat_header': True, 'header_rows': 1},
            {'op': 'table.set_font_size', 'target': a, 'size_pt': 12.0, 'preserve_header_size': True},
            {'op': 'image.scale', 'target': imgs[0]['anchor'], 'width_cm': 11.5, 'align': 'center'},
            {'op': 'paragraph.set_format', 'target': cap_a, 'align': 'center', 'keep_with_next': True},
        ]
        kwargs['profile'] = 'administrative'
    else:
        raise ValueError(f'No oracle action defined for case {cid}')

    if ops is None:
        raise ValueError(f'Oracle fell through without action or refusal for {cid}')

    edit_res = t['layout_edit'].apply_layout_plan(src_path, dst_docx, ops, **kwargs)
    res_payload = {'status': 'edited', 'layout_edit': {'status': edit_res.get('status'), 'xml_ntpr': edit_res.get('non_target_preservation', {}).get('xml_ntpr')}}
    if cid == 'T12_pagination_ripple_awareness':
        with tempfile.TemporaryDirectory(prefix='vndocs-oracle-ripple-') as rtmp:
            diff_out = t['diff_layout'].diff_layout(src_path, dst_docx, target=tbls[0]['anchor'], output_dir=rtmp)
        res_payload['pagination_ripple'] = bool(diff_out.get('pagination_ripple_detected'))
    res_file.write_text(json.dumps(res_payload, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    return {'mode': 'layout_edit', 'report': edit_res}


def _write_layout_benchmark_doc(doc_path, oracle_report):
    ov = oracle_report['summary']
    sp = oracle_report['static_payload']
    lines = [
        '# Layout Editing Benchmark Report (`evals/layout_edit_cases.json` v2.0)',
        '',
        f"- Generated: `{oracle_report['generated_at']}`",
        f"- Corpus cases: **{ov['total_cases']}** (`short`: {oracle_report['by_length_tier']['short']['n']}, `medium`: {oracle_report['by_length_tier']['medium']['n']}, `long`: {oracle_report['by_length_tier']['long']['n']})",
        f"- Deterministic Oracle Hit Rate: **{ov['hits']}/{ov['total_cases']} ({ov['hit_rate'] * 100:.1f}%)**",
        f"- Mean Checkpoint Score: **{ov['checkpoint_score_mean'] * 100:.1f}%**",
        f"- Mean `xml_ntpr`: **{ov['xml_ntpr_mean']:.4f}** | Mean ZIP non-target byte identity: **{ov['zip_byte_invariance_mean']:.4f}**",
        '',
        '## 1. Stratification Summary',
        '',
        '| Slice | Cases | Hits | Hit Rate | Mean Checkpoint | Mean `xml_ntpr` |',
        '|---|---:|---:|---:|---:|---:|',
    ]
    for label, entry in (
        [(f'length:{k}', v) for k, v in oracle_report['by_length_tier'].items()]
        + [(f'type:{k}', v) for k, v in oracle_report['by_instruction_type'].items()]
        + [(f'outcome:{k}', v) for k, v in oracle_report['by_expected_outcome'].items()]
    ):
        ntpr_s = f"{entry['xml_ntpr_mean']:.4f}" if entry['xml_ntpr_mean'] is not None else 'N/A'
        lines.append(
            f"| `{label}` | {entry['n']} | {entry['hits']} | {entry['hit_rate'] * 100:.1f}% | "
            f"{entry['checkpoint_score_mean'] * 100:.1f}% | {ntpr_s} |"
        )

    lines.extend([
        '',
        '## 2. Static Context Payload (`word/document.xml` vs `vndocs map`)',
        '',
        f"- `SKILL.md`: **{sp['instruction_overhead']['full_skill_md_bytes']} bytes** (~{sp['instruction_overhead']['full_skill_md_est_tok']} est tok)",
        f"- `SKILL.md` + `references/tooling.md`: **{sp['instruction_overhead']['full_skill_plus_tooling_bytes']} bytes** (~{sp['instruction_overhead']['full_skill_plus_tooling_est_tok']} est tok)",
        '',
        '| Tier | Cases | Mean `document.xml` | Mean `map --format markdown` | Reduction |',
        '|---|---:|---:|---:|---:|',
    ])
    for tier, ts in sp['by_length_tier'].items():
        lines.append(
            f"| `{tier}` | {ts['count']} | {ts['mean_raw_doc_xml_bytes']} B (~{ts['mean_est_tok_raw_xml']} tok) | "
            f"{ts['mean_map_markdown_bytes']} B (~{ts['mean_est_tok_map_md']} tok) | **{ts['mean_reduction_pct']}%** |"
        )

    lines.extend([
        '',
        '## 3. Per-Case Oracle Results (`T01`–`T30`)',
        '',
        '| Case ID | Tier | Type | Outcome | Hit | CP | `xml_ntpr` | Failures |',
        '|---|---|---|---|---:|---:|---:|---|',
    ])
    for r in oracle_report['cases']:
        ntpr_s = f"{r['telemetry']['xml_ntpr']:.4f}" if r['telemetry'].get('xml_ntpr') is not None else '-'
        fails = ', '.join(r['failure_reasons']) or '-'
        lines.append(
            f"| `{r['case_id']}` | {r['length_tier']} | {r['instruction_type']} | {r['expected_outcome']} | "
            f"{r['hit']} | {r['checkpoint_score'] * 100:.0f}% | {ntpr_s} | {fails} |"
        )

    lines.extend([
        '',
        '## 4. Reproduce',
        '',
        '```bash',
        'python evals/run_layout_benchmark.py --run-oracle --write-doc evals/LAYOUT-BENCHMARK.md',
        'python evals/run_layout_benchmark.py --measure-payload',
        '```',
    ])
    out = Path(doc_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text('\n'.join(lines) + '\n', encoding='utf-8')


def run_oracle_suite(write_doc=None):
    """Generate the 30-case corpus, execute deterministic vndocs oracle plans, and grade every case."""
    cases_cfg = json.loads(CASES_PATH.read_text(encoding='utf-8'))
    with tempfile.TemporaryDirectory(prefix='vndocs-oracle-') as tmp_root:
        root = Path(tmp_root)
        corpus_dir = root / 'corpus'
        sub_dir = root / 'submissions'
        generate_corpus(corpus_dir)

        results = []
        for case in cases_cfg['cases']:
            cid = case['id']
            src_docx = corpus_dir / f'{cid}.docx'
            sub_case_dir = sub_dir / cid
            _execute_oracle_case(case, src_docx, sub_case_dir)
            results.append(grade_case(case, src_docx, sub_case_dir))

        payload_stats = measure_payload(corpus_dir)

    def _slice_summary(rows):
        n = len(rows)
        hits = sum(r['hit'] for r in rows)
        ntprs = [r['telemetry']['xml_ntpr'] for r in rows if r['telemetry'].get('xml_ntpr') is not None]
        zips = [r['telemetry']['zip_parts_byte_identical'] for r in rows if r['telemetry'].get('zip_parts_byte_identical') is not None]
        return {
            'n': n,
            'hits': hits,
            'hit_rate': round(hits / n, 4) if n else 0.0,
            'checkpoint_score_mean': round(sum(r['checkpoint_score'] for r in rows) / n, 4) if n else 0.0,
            'xml_ntpr_mean': round(sum(ntprs) / len(ntprs), 4) if ntprs else None,
            'zip_byte_invariance_mean': round(sum(zips) / len(zips), 4) if zips else None,
        }

    overall = _slice_summary(results)
    report = {
        'benchmark_name': cases_cfg['benchmark_name'],
        'version': cases_cfg['version'],
        'generated_at': datetime.now().astimezone().isoformat(),
        'summary': {
            'total_cases': overall['n'],
            'hits': overall['hits'],
            'hit_rate': overall['hit_rate'],
            'checkpoint_score_mean': overall['checkpoint_score_mean'],
            'xml_ntpr_mean': overall['xml_ntpr_mean'],
            'zip_byte_invariance_mean': overall['zip_byte_invariance_mean'],
        },
        'by_length_tier': {
            lt: _slice_summary([r for r in results if r['length_tier'] == lt])
            for lt in ('short', 'medium', 'long')
        },
        'by_instruction_type': {
            it: _slice_summary([r for r in results if r['instruction_type'] == it])
            for it in ('descriptive', 'inferential')
        },
        'by_expected_outcome': {
            eo: _slice_summary([r for r in results if r['expected_outcome'] == eo])
            for eo in ('edit', 'refuse')
        },
        'static_payload': payload_stats,
        'cases': results,
    }
    if write_doc:
        _write_layout_benchmark_doc(write_doc, report)
    return report


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument('--generate-corpus', metavar='DIR', help='Generate the 30 stratified source DOCX files')
    mode.add_argument('--measure-payload', action='store_true', help='Measure static XML vs map payload sizes')
    mode.add_argument('--grade-dir', metavar='SUBMISSION_DIR', help='Grade a directory of <case_id>/<case_id>.docx + .result.json')
    mode.add_argument('--analyze-ab', metavar='AB_ROOT', help='Grade and analyze Live A/B transcripts under AB_ROOT')
    mode.add_argument('--run-oracle', action='store_true', help='Run deterministic 30-case vndocs oracle suite')

    parser.add_argument('--corpus-dir', help='Corpus directory (used by --measure-payload or --grade-dir)')
    parser.add_argument('--report', help='Optional JSON output file')
    parser.add_argument('--write-doc', help='Optional Markdown benchmark summary path (with --run-oracle)')
    args = parser.parse_args(argv)

    if args.generate_corpus:
        out = generate_corpus(args.generate_corpus)
    elif args.measure_payload:
        out = measure_payload(args.corpus_dir)
    elif args.grade_dir:
        if not args.corpus_dir:
            parser.error('--grade-dir requires --corpus-dir')
        out = grade_directory(args.grade_dir, args.corpus_dir)
    elif args.run_oracle:
        out = run_oracle_suite(write_doc=args.write_doc)
    else:
        out = analyze_ab(args.analyze_ab)

    text = json.dumps(out, ensure_ascii=False, indent=2)
    print(text)
    if args.report:
        rp = Path(args.report)
        rp.parent.mkdir(parents=True, exist_ok=True)
        rp.write_text(text + '\n', encoding='utf-8')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
