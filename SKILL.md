---
name: vietnamese-docx
description: >
  Soạn, biên tập hoặc kiểm tra văn bản tiếng Việt và tạo/sửa file DOCX theo đúng
  mẫu và nơi tiếp nhận, kể cả đoạn văn, đề mục, danh sách đánh số trong chat:
  tài liệu văn phòng, đề xuất, biên bản họp nội bộ; học thuật và giáo dục (báo cáo,
  tiểu luận, nghiên cứu, luận văn, đề thi); hành chính theo Nghị định 30/2020/NĐ-CP
  (công văn, quyết định, tờ trình, thông báo, giấy mời, biên bản, văn phong hành chính);
  văn bản của Đảng (QĐ 399-QĐ/TW, HD 05-HD/VPTW); văn bản quy phạm pháp luật (VBQPPL).
  Mỗi nhóm áp thể thức và mẫu riêng, không áp chéo NĐ30 lên văn bản Đảng, học thuật
  hay văn phòng.
---

# Vietnamese DOCX

Tạo văn bản đúng mục đích, đúng mẫu, giữ dữ kiện và có thể chỉnh sửa tiếp.

## Chọn cách xử lý

Xác định độc lập: đầu ra (text/DOCX), loại tài liệu và nơi tiếp nhận.

- Đoạn văn/danh sách trong chat: trả text trực tiếp; không tạo file hoặc bìa.
- Tạo Word: chọn profile theo `references/document-profiles.md`.
- Có mẫu hoặc file cần sửa: đọc `references/template-selection.md`; giữ cấu trúc ngoài phạm vi chỉnh sửa.
- Đề mục/đánh số: đọc `references/numbering-and-structure.md`.

Suy ra từ ngữ cảnh, không chỉ từ từ khóa “báo cáo”. Hỏi khi khác biệt về loại
hoặc nơi tiếp nhận làm thay đổi đáng kể đầu ra. Không bắt chọn tên profile.

## Ưu tiên

1. Yêu cầu trực tiếp của người dùng.
2. Mẫu và hướng dẫn của nơi tiếp nhận đúng loại, đúng phiên bản.
3. Quy định chuyên ngành/pháp lý áp dụng.
4. Mặc định của skill.

Nếu yêu cầu/mẫu khác quy định áp dụng, chỉ rõ khác biệt trước khi tuyên bố
kết quả tuân thủ. Không gọi sở thích định dạng là yêu cầu pháp lý.

## Reference theo nhu cầu

- Loại tài liệu: `references/document-profiles.md`.
- Chọn/tìm/sử dụng mẫu: `references/template-selection.md`.
- Starter đi kèm: `references/starter-templates.md`.
- Nguồn mẫu và văn bản thực tế: `references/template-sources.md`.
- Hành chính, thông số trang và cỡ chữ: `references/administrative-format-nd30.md`; đối chiếu nguồn gốc và hiệu lực khi kết quả phụ thuộc quy định hiện hành.
- Hành chính, loại văn bản, số/ký hiệu, thẩm quyền ký (TM., KT., Q., TL., TUQ.), nơi nhận, phụ lục, độ mật/khẩn, bản sao, viết hoa, khung nội dung và lỗi hay gặp: `references/administrative-doc-types-nd30.md`.
- Văn bản của Đảng (cấp ủy, chi bộ, ban đảng): `references/party-documents.md`. Không dùng thể thức NĐ30 cho văn bản Đảng.
- Văn bản quy phạm pháp luật: `references/legal-normative-documents.md`. Phân biệt với quyết định, nghị quyết cá biệt theo NĐ30.
- Văn phong hành chính – công vụ, căn cứ, viện dẫn, cụm từ khuôn mẫu: `references/administrative-writing-style.md`.
- Học thuật: `references/academic-report.md`.
- Biên bản: `references/meeting-minutes.md`.
- Biên tập: `references/editorial-quality-vi.md`.
- Đánh số: `references/numbering-and-structure.md`.
- Triển khai Word: `references/style-spec.md`; lệnh, cấu hình và giới hạn ở `references/tooling.md`.
- Trước bàn giao: `references/validation-checklist.md`.

Chỉ đọc phần liên quan trong `references/` theo bảng routing ở trên; ưu tiên các tài liệu quy chuẩn thay vì nạp toàn bộ thư mục tham khảo.

## Nội dung

Không bịa dữ kiện, nguồn, căn cứ, cơ quan, số văn bản, người ký hoặc quyết định
cuộc họp. Mẫu thực tế chỉ dùng học cấu trúc và cách diễn đạt; không mang thông
tin của văn bản cũ sang tài liệu mới. Giữ khác biệt giữa ý kiến, đề xuất,
quyết định và nghĩa vụ khi biên tập.

Bản nháp có thể dùng `[CẦN BỔ SUNG: ...]` và công bố trường thiếu. Hỏi khi thiếu
thông tin làm thay đổi bản chất tài liệu. Không thêm số liệu để câu có vẻ cụ thể.
Số hiệu văn bản nhớ được nhưng chưa có trong đầu vào hoặc chưa tra trên nguồn
chính thức là chưa xác minh: ghi `[CẦN BỔ SUNG: căn cứ ...]` thay vì điền.

## DOCX

Ưu tiên mẫu người dùng hoặc mẫu chính thức đã xác minh phù hợp. Không có mẫu:
dùng cấu trúc tối thiểu theo profile và công bố mặc định quan trọng.
`scripts/build_docx.py` là thư viện helper, không phải generator đầy đủ.
Không chạy `setup_document` lên mẫu sẵn để ghi đè style.
Mẫu token dùng `fill_template.py`; mẫu không token có thể dùng `edit_docx.py`
với hash nguồn và slot map đã xác minh. Xem `references/tooling.md`.

Giữ section, heading, numbering, header/footer, bảng, mục lục và trường số trang
của mẫu. Không áp quốc hiệu, bìa, thư ngỏ, nơi nhận hoặc cấm bullet cho mọi tài liệu.

## Kiểm tra và bàn giao

Text: rà dữ kiện, cấp đề mục, thứ tự và phạm vi yêu cầu.
DOCX: theo checklist, dùng helper khi phù hợp và kiểm tra hiển thị thật:

```bash
python scripts/validate_docx.py output.docx --profile academic
```

Validator kiểm tra thuộc tính đã hỗ trợ trong từng section, style kế thừa,
body/bảng/header/footer; nhận `--config` cho mẫu và `--draft` cho bản nháp.
Đọc `references/tooling.md` để hiểu pass/fail/partial và danh sách unchecked.
Các thuộc tính chưa kiểm tra không được gọi là đã đạt.
Render thật bằng `scripts/render_docx.py output.docx --png`; thiếu renderer
phải báo chưa kiểm tra. Mở xem ảnh mọi trang trước khi kết luận bố cục đạt.

Sửa lỗi làm sai nội dung/mất bố cục trước bàn giao. Trả file khi đã tạo file;
nêu mẫu/chuẩn, trường thiếu và kiểm tra chưa thực hiện. Không gọi tài liệu
“chuẩn NĐ30” chỉ vì đúng font và lề.
