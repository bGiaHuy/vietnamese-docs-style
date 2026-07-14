# Bảng định tuyến Hồ sơ Tài liệu (Document Profiles Routing)

> **Mục đích**: Bảng này hướng dẫn AI agent cách lựa chọn các quy tắc định dạng và thể thức phù hợp dựa trên loại tài liệu người dùng yêu cầu.

AI agent phải xác định loại tài liệu (profile) trước khi bắt đầu tạo file. Nếu không chắc chắn, hãy sử dụng `custom` hoặc hỏi lại người dùng.

## Bảng định tuyến

| Profile | Tiêu chuẩn / Chuẩn mực | Quốc hiệu & Tiêu ngữ | Trang Bìa (Cover Page) | Kính gửi / Nơi nhận | Ngôi kể | Dùng Bullet (`•`) | Auto Numbering | Shading Bảng |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| **`administrative`** (Công văn, Quyết định, Tờ trình...) | **Nghị định 30/2020/NĐ-CP** | **Bắt buộc** | Không | **Bắt buộc** | Ngôi thứ 3 (Khách quan) | Cấm | Cấm (Đánh số tay) | Cấm |
| **`academic`** (Báo cáo, Tiểu luận, Khóa luận...) | Mặc định (A4, TNR) hoặc Template của Trường | Không (trừ khi template yêu cầu) | **Có** (Thường xuyên) | Không | Linh hoạt (Ngôi thứ 3 hoặc "Chúng tôi") | Được phép | Được phép | Được phép (Tùy chọn) |
| **`proposal`** (Đề xuất dự án, Kế hoạch kinh doanh) | Mặc định (A4, TNR) hoặc Template của Tổ chức | Không | **Có** | Không (Thường có Thư ngỏ thay thế) | Linh hoạt | Được phép | Được phép | Được phép |
| **`minutes-administrative`** (Biên bản hành chính nhà nước) | **Nghị định 30/2020/NĐ-CP** + TT01 | **Bắt buộc** | Không | Không | Ngôi thứ 3 | Cấm | Cấm (Đánh số tay) | Cấm |
| **`minutes-general`** (Biên bản họp nhóm, nội bộ) | Cấu trúc Meeting Minutes chung (A4, TNR) | Không | Không | Không | Linh hoạt | Được phép | Được phép | Được phép |
| **`custom`** (Mẫu riêng) | Tuân thủ tuyệt đối Template/Yêu cầu của User | Tùy template | Tùy template | Tùy template | Tùy template | Tùy template | Tùy template | Tùy template |

## Nguyên tắc áp dụng (Priority Chain)

Trong trường hợp có mâu thuẫn giữa các quy định, AI agent phải ưu tiên theo thứ tự sau (từ cao xuống thấp):

1.  **Yêu cầu trực tiếp của người dùng**: VD: "Hãy dùng bullet `•` cho quyết định này". (Dù sai NĐ30 nhưng user yêu cầu thì phải làm theo).
2.  **Template của tổ chức/trường học**: Nếu user tải lên một template có sẵn.
3.  **Quy định chuyên ngành**: Ví dụ, Giáo án sẽ có cấu trúc riêng không theo NĐ30.
4.  **Nghị định 30/2020/NĐ-CP**: Áp dụng chặt chẽ cho profile `administrative` và `minutes-administrative`.
5.  **Mặc định của Skill**: Các quy tắc chung (A4, Times New Roman, v.v.) được định nghĩa trong `style-spec.md`.

## Ghi chú về Style
*   Tất cả các profile mặc định sử dụng khổ giấy **A4 (210 x 297mm)** trừ khi có yêu cầu khác.
*   Màu chữ mặc định cho phần nội dung (Body text) luôn là **Đen (#000000)**.
*   Font chữ ưu tiên là **Times New Roman**.

