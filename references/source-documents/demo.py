import sys
import os
from docx import Document
from docx.shared import Inches, Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

# Ensure vn-report-pro-v3 is in path
sys.path.append(os.path.join(os.path.dirname(__file__), 'vn-report-pro-v3'))
from references.style_spec import (
    BLACK, add_centered, add_blank, add_bullet, add_bullets,
    add_table, add_team_table, add_right, add_body, add_h1, add_h2
)

def configure_styles(doc):
    # Configure Normal style
    style_normal = doc.styles['Normal']
    style_normal.font.name = 'Times New Roman'
    style_normal.font.size = Pt(12)
    style_normal.font.color.rgb = BLACK
    style_normal.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    
    # Configure Headings style to be black and Times New Roman
    for level, size in [(1, 16), (2, 13), (3, 12)]:
        style_name = f'Heading {level}'
        h_style = doc.styles[style_name]
        h_style.font.name = 'Times New Roman'
        h_style.font.size = Pt(size)
        h_style.font.bold = True
        h_style.font.color.rgb = BLACK
        h_style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        h_style.paragraph_format.keep_with_next = True

def build_meeting_minutes():
    doc = Document()
    
    # Apply document-level styles
    configure_styles(doc)
    
    # 1. Page Setup (Letter size, standard margins)
    section = doc.sections[0]
    section.page_width = Inches(8.5)
    section.page_height = Inches(11.0)
    section.left_margin = Cm(3.0)
    section.right_margin = Cm(3.0)
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    
    # 2. National Header & Org Header (Decree 30 style via borderless table)
    # To place Left Org Name and Right National Epithet on the same top level
    header_table = doc.add_table(rows=1, cols=2)
    header_table.style = 'Normal Table'
    header_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # Left Cell: Org / Unit Info
    cell_left = header_table.cell(0, 0)
    p_left = cell_left.paragraphs[0]
    p_left.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_left.paragraph_format.space_after = Pt(0)
    
    r_org = p_left.add_run("TRƯỜNG ĐẠI HỌC FPT\n")
    r_org.font.name = 'Times New Roman'
    r_org.font.size = Pt(12)
    r_org.font.color.rgb = BLACK
    
    r_unit = p_left.add_run("HỘI ĐỒNG THI ĐUA")
    r_unit.bold = True
    r_unit.font.name = 'Times New Roman'
    r_unit.font.size = Pt(12)
    r_unit.font.color.rgb = BLACK
    
    # Simple line under organization name
    r_line_l = p_left.add_run("\n----------")
    r_line_l.font.name = 'Times New Roman'
    r_line_l.font.size = Pt(10)
    r_line_l.font.color.rgb = BLACK
    
    p_no = cell_left.add_paragraph()
    p_no.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_no.paragraph_format.space_before = Pt(6)
    r_no = p_no.add_run("Số: 05/BB-HĐTĐ")
    r_no.font.name = 'Times New Roman'
    r_no.font.size = Pt(13)
    r_no.font.color.rgb = BLACK
    
    # Right Cell: National Info & Location
    cell_right = header_table.cell(0, 1)
    p_right = cell_right.paragraphs[0]
    p_right.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_right.paragraph_format.space_after = Pt(0)
    
    r_nat1 = p_right.add_run("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\n")
    r_nat1.bold = True
    r_nat1.font.name = 'Times New Roman'
    r_nat1.font.size = Pt(12)
    r_nat1.font.color.rgb = BLACK
    
    r_nat2 = p_right.add_run("Độc lập - Tự do - Hạnh phúc")
    r_nat2.bold = True
    r_nat2.font.name = 'Times New Roman'
    r_nat2.font.size = Pt(13)
    r_nat2.font.color.rgb = BLACK
    
    # Line under national epithet
    r_line_r = p_right.add_run("\n____________________")
    r_line_r.font.name = 'Times New Roman'
    r_line_r.font.size = Pt(10)
    r_line_r.font.color.rgb = BLACK
    
    p_loc = cell_right.add_paragraph()
    p_loc.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_loc.paragraph_format.space_before = Pt(6)
    r_loc = p_loc.add_run("Hà Nội, ngày 21 tháng 03 năm 2026")
    r_loc.italic = True
    r_loc.font.name = 'Times New Roman'
    r_loc.font.size = Pt(13)
    r_loc.font.color.rgb = BLACK
    
    # Prevent table width deformation
    for cell in header_table.columns[0].cells:
        cell.width = Inches(3.2)
    for cell in header_table.columns[1].cells:
        cell.width = Inches(4.3)
        
    add_blank(doc)
    add_blank(doc)
    
    # 3. Document Title
    add_centered(doc, "BIÊN BẢN HỌP HỘI ĐỒNG THI ĐUA", 16, True)
    add_centered(doc, "V/v Đánh giá và khen thưởng dự án AI học máy năm 2026", 13, True, True)
    add_blank(doc)
    
    # 4. Context/Time Intro
    add_body(doc, "Hôm nay, vào lúc 09 giờ 30 phút, ngày 21 tháng 03 năm 2026, tại Văn phòng Hội đồng trường Đại học FPT, tiến hành họp đánh giá thi đua và đề xuất khen thưởng cho các dự án nghiên cứu phát triển công nghệ học máy tiêu biểu.")
    add_blank(doc)
    
    # 5. Section I: Thành phần tham dự
    add_h1(doc, "I. THÀNH PHẦN THAM DỰ")
    add_body(doc, "Cuộc họp diễn ra dưới sự điều hành của Chủ trì và Thư ký với danh sách thành viên tham gia chi tiết dưới đây:")
    add_blank(doc)
    
    members = [
        ("Nguyễn Văn A", "Chủ trì cuộc họp", "Ban Giám hiệu"),
        ("Trần Thị B", "Thư ký cuộc họp", "Phòng Đào tạo"),
        ("Phạm Minh C", "Thành viên Hội đồng", "Khoa Kỹ thuật máy tính"),
        ("Lê Hoàng D", "Thành viên Hội đồng", "Ban Khảo thí & Đảm bảo chất lượng")
    ]
    add_team_table(doc, members)
    add_blank(doc)
    
    # 6. Section II: Nội dung cuộc họp
    add_h1(doc, "II. NỘI DUNG CUỘC HỌP")
    
    add_h2(doc, "1. Báo cáo tiến độ và kết quả các dự án AI")
    add_body(doc, "Đại diện ban chuyên môn đã trình bày báo cáo tóm tắt các dự án AI đang được áp dụng trực tiếp vào quản lý đào tạo:")
    add_bullets(doc, [
        "Dự án tự động hóa chấm điểm và phân tích kết quả học tập đạt độ chính xác 98% đối với lớp SSA101.",
        "Dự án nhận diện hành vi trong phòng thi đã thử nghiệm thành công tại 3 phòng máy chuyên dụng.",
        "Đã giảm thiểu 45% thời gian tổng hợp kết quả học tập định kỳ của sinh viên khoa Công nghệ thông tin."
    ], level=1)
    
    add_h2(doc, "2. Ý kiến thảo luận của các thành viên Hội đồng")
    add_body(doc, "Các thành viên Hội đồng thảo luận chi tiết về các giải pháp cải tiến hiệu suất hệ thống:")
    add_bullet(doc, "Đồng chí Phạm Minh C đề xuất tiếp tục mở rộng quy mô dự án nhận diện hành vi cho toàn bộ kỳ thi học kỳ tới.", level=1)
    add_bullet(doc, "Đồng chí Lê Hoàng D đề nghị tăng cường công tác kiểm định dữ liệu và bảo mật thông tin cá nhân của sinh viên.", level=1)
    add_bullet(doc, "Ý kiến phản hồi từ Chủ trì cuộc họp:", level=1)
    add_bullet(doc, "Nhất trí cao với các đề xuất cải tiến kỹ thuật.", level=2)
    add_bullet(doc, "Yêu cầu xây dựng khung đánh giá chất lượng dự án trước ngày 15 tháng 04 năm 2026.", level=2)
    add_blank(doc)
    
    # 7. Section III: Kết luận chung
    add_h1(doc, "III. KẾT LUẬN CHUNG VÀ PHÂN CÔNG CÔNG VIỆC")
    add_body(doc, "Sau khi thảo luận và biểu quyết, Hội đồng thi đua đi đến thống nhất các quyết nghị sau:")
    add_bullets(doc, [
        "Thông qua tờ trình đề xuất khen thưởng tập thể ML-Zero vì có thành tích xuất sắc trong đổi mới sáng tạo.",
        "Chuyển giao hệ thống quản lý học tập tự động cho phòng Đào tạo vận hành từ học kỳ Fall 2026.",
        "Biên bản này được lập thành 03 bản, có giá trị pháp lý như nhau."
    ], level=1)
    add_blank(doc)
    
    # Task assignments table
    add_h2(doc, "Bảng phân công trách nhiệm thực hiện:")
    assignments = [
        ("Hoàn thiện hồ sơ khen thưởng", "Trần Thị B", "05/04/2026"),
        ("Xây dựng khung đánh giá dự án", "Phạm Minh C", "15/04/2026"),
        ("Báo cáo tiến độ triển khai kỹ thuật", "Lê Hoàng D", "20/04/2026")
    ]
    add_table(doc, ["STT", "Nội dung công việc", "Người chịu trách nhiệm", "Hạn hoàn thành"], [
        [str(i+1), task, owner, deadline] for i, (task, owner, deadline) in enumerate(assignments)
    ])
    add_blank(doc)
    add_blank(doc)
    
    # 8. Signatures Block (Decree 30 double columns style via borderless table)
    sig_table = doc.add_table(rows=1, cols=2)
    sig_table.style = 'Normal Table'
    sig_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    cell_sec = sig_table.cell(0, 0)
    p_sec = cell_sec.paragraphs[0]
    p_sec.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sec.paragraph_format.space_after = Pt(0)
    r_sec_title = p_sec.add_run("THƯ KÝ HỘI ĐỒNG\n")
    r_sec_title.bold = True
    r_sec_title.font.name = 'Times New Roman'
    r_sec_title.font.size = Pt(12)
    r_sec_title.font.color.rgb = BLACK
    p_sec.add_run("(Ký và ghi rõ họ tên)\n\n\n\n\n")
    r_sec_name = p_sec.add_run("Trần Thị B")
    r_sec_name.bold = True
    r_sec_name.font.name = 'Times New Roman'
    r_sec_name.font.size = Pt(12)
    r_sec_name.font.color.rgb = BLACK
    
    cell_chair = sig_table.cell(0, 1)
    p_chair = cell_chair.paragraphs[0]
    p_chair.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_chair.paragraph_format.space_after = Pt(0)
    r_chair_title = p_chair.add_run("CHỦ TRÌ HỘI ĐỒNG\n")
    r_chair_title.bold = True
    r_chair_title.font.name = 'Times New Roman'
    r_chair_title.font.size = Pt(12)
    r_chair_title.font.color.rgb = BLACK
    p_chair.add_run("(Ký và ghi rõ họ tên)\n\n\n\n\n")
    r_chair_name = p_chair.add_run("Nguyễn Văn A")
    r_chair_name.bold = True
    r_chair_name.font.name = 'Times New Roman'
    r_chair_name.font.size = Pt(12)
    r_chair_name.font.color.rgb = BLACK
    
    for cell in sig_table.columns[0].cells:
        cell.width = Inches(3.7)
    for cell in sig_table.columns[1].cells:
        cell.width = Inches(3.8)

    # Save to file
    doc.save("bien_ban_hop.docx")
    print("Document successfully created and saved as: bien_ban_hop.docx")
    return doc

def validate_generated_document(doc):
    print("Running Decree 30 validation suite...")
    errors = []
    full_text = "\n".join([p.text for p in doc.paragraphs])

    # 1. No bullet dot (•)
    if "•" in full_text:
        errors.append("ERROR: Found standard bullet dot (•) — must use manual dashes/plusses.")

    # 2. All headings are black
    for level in [1, 2, 3]:
        style_name = f'Heading {level}'
        if style_name in [s.name for s in doc.styles]:
            color = doc.styles[style_name].font.color.rgb
            if color and color != RGBColor(0, 0, 0):
                errors.append(f"ERROR: {style_name} font color is {color}, must be BLACK (#000000).")

    # 3. Margins check
    section = doc.sections[0]
    left_margin_cm = round(section.left_margin.cm, 1)
    right_margin_cm = round(section.right_margin.cm, 1)
    top_margin_cm = round(section.top_margin.cm, 1)
    bottom_margin_cm = round(section.bottom_margin.cm, 1)
    
    if left_margin_cm != 3.0:
        errors.append(f"ERROR: Left margin is {left_margin_cm}cm, must be 3.0cm.")
    if top_margin_cm != 2.0:
        errors.append(f"ERROR: Top margin is {top_margin_cm}cm, must be 2.0cm.")
    if bottom_margin_cm != 2.0:
        errors.append(f"ERROR: Bottom margin is {bottom_margin_cm}cm, must be 2.0cm.")

    if not errors:
        print("SUCCESS: Document validation passed! Strict Decree 30 / FPT requirements met.")
    else:
        print("VALIDATION ERRORS DETECTED:")
        for err in errors:
            print(f"- {err}")
            
if __name__ == "__main__":
    doc = build_meeting_minutes()
    validate_generated_document(doc)
