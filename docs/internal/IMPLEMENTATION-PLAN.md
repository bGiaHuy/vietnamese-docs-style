# Trạng thái triển khai

Định hướng giữ nguyên: skill vietnamese-docx cho text/đánh số và DOCX văn phòng,
học thuật, hành chính; chọn mẫu theo yêu cầu và nơi tiếp nhận.

## Các hạng mục đã thực hiện

- Routing tách đầu ra/loại tài liệu/nơi nhận, có office profile chung; text ngắn không bị ép tạo Word.
- Reference thống nhất, bỏ quy tắc áp đặt và ví dụ bịa dữ kiện; NĐ30 đối chiếu bản ký, giới hạn kiểm tra hiệu lực ghi rõ.
- Cấu hình chung profile/template, per-section, named style roles; custom giữ mẫu.
- Generator kế thừa style, numbering Word và PAGE thật cho hành chính.
- Validator kiểm tra các thuộc tính được hỗ trợ, report pass/fail/partial và unchecked.
- Fill token qua run và edit slot đã xác minh; patch ZIP giữ các member ngoài vùng sửa, không tái dựng toàn bộ mẫu.
- Render thật, chống nhận nhầm PDF cũ, ảnh trang phục vụ QA.
- Ba starter dự án viết đã điền/render/xem; một mẫu UIT thực tế thử 21 vị trí và kiểm chứng fidelity.
- 39 test pass, gồm end-to-end nhiều section, custom, bản nháp, renderer thiếu và số trang thật; lint và skill-creator validation pass.
- CI cấu hình cho source repo; chưa push/chạy trên GitHub theo phạm vi đã giao.
- Các case text/query được thực hiện và ghi nhận ở evals; không gọi là benchmark model độc lập.
- Gói cài đặt chỉ gồm skill/reference/helper/starter; loại research, runtime và cache.

Chi tiết đối chiếu yêu cầu và giới hạn: VERIFICATION.md. Không có hạng mục triển
khai bắt buộc đang để dở trong phạm vi đã nêu; vẫn cần QA lại mỗi tài liệu mới,
xác minh mẫu của nơi tiếp nhận và quy định tại thời điểm sử dụng.
