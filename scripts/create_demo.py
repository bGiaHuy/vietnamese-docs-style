"""Create a reproducible office template and filled verification document."""
import json
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION_START, WD_ORIENT
from docx.shared import Cm

try:
    from .build_docx import setup_document, add_body, add_table, add_numbered
    from .docx_package import ensure_utf8_stdio
    from .fill_template import fill_template
except ImportError:
    from build_docx import setup_document, add_body, add_table, add_numbered
    from docx_package import ensure_utf8_stdio
    from fill_template import fill_template


def create_demo(output_dir):
    directory = Path(output_dir)
    directory.mkdir(parents=True, exist_ok=True)
    doc = Document()
    setup_document(doc, 'minutes-general')
    doc.sections[0].header.paragraphs[0].text = 'TÀI LIỆU THỬ NGHIỆM'
    doc.add_paragraph('Biên bản họp thử nghiệm', style='Title')
    add_body(doc, 'Mẫu do dự án xây dựng để kiểm tra công cụ; mọi dữ liệu là giả định, không phải biên bản đã ban hành.')
    add_body(doc, 'Cuộc họp: {{meeting}}')
    add_body(doc, 'Thời gian: {{time}}')
    p = doc.add_paragraph('Người ghi nhận: ')
    p.add_run('{{secre').bold = True
    p.add_run('tary}}').bold = True
    doc.add_heading('Nội dung', level=1)
    add_body(doc, '{{discussion}}')
    doc.add_heading('Kết luận', level=1)
    add_numbered(doc, ['{{decision1}}', '{{decision2}}'])
    doc.add_heading('Phân công', level=1)
    add_table(doc, ['Công việc', 'Phụ trách', 'Hạn dự kiến'],
              [['{{task}}', '{{owner}}', '{{deadline}}']], 'minutes-general')
    second = doc.add_section(WD_SECTION_START.NEW_PAGE)
    second.page_width, second.page_height = Cm(29.7), Cm(21)
    second.orientation = WD_ORIENT.LANDSCAPE
    doc.add_heading('Phụ lục theo dõi thử nghiệm', level=1)
    add_body(doc, 'Trang ngang kiểm tra việc giữ section khi điền mẫu. Nội dung phụ lục ngoài token cần được bảo toàn.')
    add_table(doc, ['STT', 'Hạng mục', 'Điều kiện hoàn thành', 'Ghi chú'],
              [['1', 'Bản nháp', 'Đối chiếu yêu cầu đầu vào', 'Dữ liệu giả định'],
               ['2', 'Bàn giao', 'Kiểm tra nội dung và bố cục', 'Không phải quyết định thực tế']],
              'minutes-general')
    source = directory / 'office-template.docx'
    doc.save(source)
    values = {
        'meeting': 'Rà soát kế hoạch tài liệu mẫu', 'time': '09 giờ ngày 13/10/2026 (giả định)',
        'secretary': 'Thành viên A (giả định)',
        'discussion': 'Nhóm xem xét yêu cầu tạo tài liệu theo mẫu, giữ dữ kiện và kiểm tra định dạng sau khi lưu.',
        'decision1': 'Hoàn thành bản nháp theo cấu trúc đã thống nhất trong tình huống thử nghiệm.',
        'decision2': 'Rà soát dữ liệu trước khi sử dụng tài liệu cho công việc thực tế.',
        'task': 'Chuẩn bị bản nháp', 'owner': 'Thành viên B (giả định)',
        'deadline': '15/10/2026 (giả định)',
    }
    (directory / 'values.json').write_text(json.dumps(values, ensure_ascii=False, indent=2), encoding='utf-8')
    return fill_template(source, directory / 'office-filled.docx', values)


def main(argv=None):
    import argparse
    ensure_utf8_stdio()
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output-dir', default='artifacts/demo')
    args = parser.parse_args(argv)
    print(json.dumps(create_demo(args.output_dir), ensure_ascii=False, indent=2))
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
