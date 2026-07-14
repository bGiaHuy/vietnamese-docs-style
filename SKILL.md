---
name: vietnamese-docx
description: >
  Tạo, định dạng và kiểm tra tài liệu DOCX tiếng Việt, gồm văn bản hành chính,
  báo cáo học thuật, đề xuất dự án và biên bản. Sử dụng khi người dùng yêu cầu
  tạo hoặc chỉnh sửa file Word tiếng Việt. Chỉ áp dụng thể thức Nghị định
  30/2020/NĐ-CP khi tài liệu thực sự là văn bản hành chính hoặc người dùng
  yêu cầu rõ ràng; không tự áp dụng NĐ30 cho mọi báo cáo học thuật.
---

# Vietnamese DOCX

Tạo tài liệu Word tiếng Việt đúng loại văn bản, đúng nguồn chuẩn và có thể
kiểm chứng bằng cấu trúc lẫn kết quả render.

## Nguyên tắc ưu tiên

Khi các yêu cầu xung đột, áp dụng thứ tự sau:

1. Yêu cầu trực tiếp của người dùng.
2. Template hoặc quy định của trường/cơ quan do người dùng cung cấp.
3. Quy định chuyên ngành tương ứng.
4. Nghị định 30/2020/NĐ-CP đối với văn bản hành chính.
5. Mặc định của skill.

Không trình bày sở thích mặc định của skill như một yêu cầu pháp lý.

## Quy trình

### 1. Xác định profile tài liệu

Chọn một trong các profile:

- `administrative`: công văn, quyết định, thông báo, tờ trình, giấy mời.
- `academic`: báo cáo nghiên cứu, tiểu luận, bài tập môn học.
- `proposal`: đề xuất dự án hoặc kế hoạch.
- `minutes-administrative`: biên bản hành chính.
- `minutes-general`: biên bản họp nhóm hoặc doanh nghiệp.
- `custom`: có template riêng do người dùng cung cấp.

Đọc `references/document-profiles.md` để biết thêm chi tiết về từng profile.

Nếu loại tài liệu ảnh hưởng đáng kể đến cấu trúc mà chưa xác định được,
hỏi người dùng một câu ngắn. Không hỏi lại thông tin đã có.

### 2. Đọc reference phù hợp

- Với `administrative` và `minutes-administrative`, đọc `references/administrative-format-nd30.md`.
- Với `academic`, đọc `references/academic-report.md`.
- Với biên bản, đọc `references/meeting-minutes.md`.
- Với mọi tài liệu tiếng Việt, đọc `references/editorial-quality-vi.md`.
- Đọc `references/validation-checklist.md` trước khi kiểm tra output.
- Đọc `references/style-spec.md` để biết thông số kỹ thuật cấu trúc tài liệu.

Không tải reference không liên quan.

### 3. Xử lý thông tin thiếu

Không tự bịa:

- Tên cơ quan.
- Số và ký hiệu văn bản.
- Căn cứ pháp lý.
- Người ký.
- Số liệu nghiên cứu.
- Nguồn tham khảo.
- Ngày, địa điểm hoặc thành viên.

Dùng placeholder có nhãn `[CẦN BỔ SUNG: ...]` nếu người dùng cho phép tạo
bản nháp. Hỏi lại nếu thiếu dữ liệu làm thay đổi tính hợp lệ của tài liệu.

### 4. Tạo DOCX

Ưu tiên dùng template trong `assets/` (nếu có) và script `scripts/build_docx.py`.

Mặc định cho văn bản hành chính khi không có template khác:

- Khổ A4 (210 x 297 mm).
- Font Times New Roman, Unicode.
- Màu chữ đen (#000000).
- Lề trong phạm vi NĐ30 (trái 30-35mm, phải 15-20mm, trên/dưới 20-25mm).
- Body căn đều hai lề (Justify).
- Cỡ chữ, vị trí và kiểu chữ theo từng thành phần thể thức.

Không áp dụng tự động các quy tắc sau cho mọi profile:

- Quốc hiệu và tiêu ngữ.
- Trang bìa.
- Nơi nhận.
- Thư ngỏ.
- Ngôi thứ ba.
- Mục tiêu bắt buộc phải có số.
- Cấm bullet hoặc auto-numbering.
- Cấm shading của bảng.

Chỉ áp dụng khi profile hoặc template yêu cầu.

### 5. Biên tập ngôn ngữ

Bảo đảm (theo `references/editorial-quality-vi.md`):

- Câu rõ chủ thể và hành động.
- Không bịa dữ kiện để làm văn bản có vẻ cụ thể.
- Hạn chế sáo rỗng và từ nối dư thừa.
- Không lạm dụng bị động.
- Thuật ngữ và cách xưng hô nhất quán.
- Phân biệt dữ kiện, nhận định và kiến nghị.
- Không tuyên bố có nguồn nếu chưa kiểm chứng.

### 6. Kiểm tra

Chạy:

```bash
python scripts/validate_docx.py <file.docx> --profile <profile>
python scripts/render_docx.py <file.docx>
```

Validation phải kiểm tra:

* Tất cả section và kích thước trang (A4).
* Toàn bộ lề.
* Font và màu trong paragraph, bảng, header và footer.
* Các trường bắt buộc theo profile.
* Heading, numbering và page break.
* Bảng tràn lề.
* Placeholder chưa được xử lý.
* Render lỗi, trang trắng bất thường hoặc nội dung bị cắt.

Không bàn giao nếu còn lỗi nghiêm trọng. Nếu không thể sửa vì thiếu dữ liệu,
nêu rõ trường còn thiếu.

### 7. Bàn giao

Trả về:

* File DOCX hoàn chỉnh.
* Profile đã áp dụng.
* Chuẩn hoặc template đã sử dụng.
* Những placeholder còn lại, nếu có.

Không gọi tài liệu là “chuẩn NĐ30” nếu chỉ áp dụng một phần của NĐ30.

