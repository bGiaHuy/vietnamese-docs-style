# Nghiệm thu vietnamese-docx

Ngày: 2026-10-13. Giữ nguyên tên skill. Đây là skill cho agent, không phải dịch
vụ chứng nhận văn bản hoặc hệ thống tự hiểu mọi template không cần kiểm tra.

| Yêu cầu | Bằng chứng hiện tại |
| --- | --- |
| Text ngắn, đề mục và đánh số không cần Word | SKILL.md, numbering-and-structure.md; kết quả và rà soát trong evals/manual-results.md |
| Chọn mẫu theo query/mục đích/nơi nhận | document-profiles.md, template-selection.md, starter-templates.md; 10 case thủ công gồm thiếu trường và thông báo văn phòng |
| Văn phòng chung, học thuật, hành chính | Profile office và các profile chuyên biệt; office-general, template-suite và demo trong artifacts đã tạo/kiểm tra |
| Quy định khác mặc định | Reference NĐ30 đối chiếu Điều 1–2, Phụ lục I và mẫu 1.4/1.5 trong bản ký Chính phủ; academic và custom không bị ép chuẩn hành chính |
| Nguồn và mẫu thực tế | template-sources.md; ba file UIT đã đọc; source/hash/contract/fidelity/render của mẫu biên bản hội đồng lưu ở artifacts/research/uit-trial |
| Template/config và giữ cấu trúc | document_rules.py, build_docx.py, docx_package.py, fill_template.py, edit_docx.py; 17 ZIP member và 96 paragraph ngoài 21 slot UIT giữ nguyên; thuộc tính section/table/paragraph/run không đổi |
| Validator có phạm vi trung thực | Khổ/lề từng section, style kế thừa ASCII/hAnsi, roles font-size/emphasis, bảng lồng, header/footer, placeholder, required literal, numbering reference integrity; unchecked luôn công bố |
| Render thật và lỗi thật | render_docx.py; bản LibreOffice giải nén chạy được; PDF/PNG thực được tạo và xem; test thiếu renderer và stale PDF không báo thành công |
| Kiểm thử, dependency, CI | 39 pytest pass trên Python 3.12 với LibreOffice thật; Ruff/compileall/quick_validate pass; requirements và workflow source repo có test/static checks |
| Gói cài sạch | scripts/package_skill.py; ZIP có skill/root đúng tên, không có dataset, file mẫu bên thứ ba, .git, .codex, runtime, cache hoặc tests; smoke test trên bản giải nén mới |

## Artifact đã xem

- artifacts/demo/office-filled.docx: 2 trang dọc/ngang, token chia run, bảng/numbering.
- artifacts/template-suite/: ba bản filled, mỗi bản 1 trang; config validation và ảnh QA đã xem.
- artifacts/office-general/office-notice.docx: 1 trang thông báo nội bộ, không tự thêm quốc hiệu/bìa học thuật.
- artifacts/research/uit-trial/uit-metadata-trial.docx: 2 trang; dữ liệu giả định rõ ràng, giữ bảng/ô gộp/numbering/header và vùng ngoài sửa.

File ZIP và báo cáo smoke cuối nằm ở dist/vietnamese-docx.zip và
artifacts/package-smoke-report.json. Smoke test dùng các helper từ bản giải nén
mới, không gọi helper trong source repo. Các starter được đóng gói có manifest
nguồn, hash, version, scope và trial review; mẫu UIT không đóng gói khi chưa rõ quyền.

## Giới hạn được công bố

- Mỗi nội dung mới cần validation và xem render lại; trial một trang không bảo đảm mọi độ dài đều vừa mẫu.
- Validator không chứng nhận pháp lý/sự thật và không bao phủ mọi conditional style, theme, complex script, field, text box, chart, footnote, revision hoặc content control. Những thao tác không hỗ trợ được từ chối hoặc ghi unchecked.
- Config required_text kiểm tra literal, không hiểu thẩm quyền, ý nghĩa hoặc vị trí. Generator không tự điền người ký/căn cứ/ngày khi chưa biết.
- NĐ30 đã đối chiếu bản ký; trạng thái CSDL ghi theo kết quả index và truy cập trực tiếp bị 403/502, không nhận là đã kiểm tra live toàn bộ sửa đổi. Tài liệu chính thức phải xác minh phạm vi/hiệu lực/mẫu chuyên ngành khi dùng.
- Starter hành chính tham chiếu mẫu 1.4, không phải mẫu được cơ quan chứng nhận. Mẫu học thuật không thay quy định riêng của trường; mẫu UIT nghiên cứu có version cũ nên không mặc định áp cho khóa hiện tại.
- Text/query được đánh giá thủ công bởi agent phát triển, chưa có benchmark độc lập hay thử nghiệm nhiều model/người dùng.
- CI đã cấu hình nhưng chưa chạy trên GitHub vì không push. Kiểm tra local thực hiện trên Python 3.12; matrix 3.10/3.12 sẽ chạy khi người dùng push.
- LibreOffice hệ thống trong phiên này thiếu DLL. Bản chính thức giải nén ở .tools hoạt động; VNDOCS_RENDERER trỏ tới .tools/lo-extracted/program/soffice.com. Runtime đó không đi vào ZIP.

Không commit, push, cài skill toàn máy hoặc phát hành công khai trong nhiệm vụ này.
