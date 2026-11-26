# Hướng dẫn Quy chuẩn Thể thức & An toàn Chỉnh sửa Tài liệu `.docx` Tiếng Việt

Tài liệu này tổng hợp các quy chuẩn thể thức, định dạng và quy tắc an toàn khi chỉnh sửa tệp Word (`.docx`) tiếng Việt. Chỉ áp dụng chuẩn tương ứng với đúng **loại văn bản** và **cơ quan/nơi tiếp nhận** của tài liệu nguồn; không áp chéo quy chuẩn hành chính lên văn bản Đảng hay học thuật.

---

## 1. Phân loại Tài liệu & Không Áp Chéo Quy chuẩn

### 1.1 Văn bản Hành chính Nhà nước (Nghị định 30/2020/NĐ-CP)
- **Khổ giấy & lề trang:** Khổ A4 (`21,0 x 29,7 cm`). Lề chuẩn: trên `2,0–2,5 cm`, dưới `2,0–2,5 cm`, trái `3,0–3,5 cm`, phải `1,5–2,0 cm`.
- **Phông & cỡ chữ:** Phông `Times New Roman` (mã Unicode), chữ nội dung chính cỡ `13–14pt`; cỡ chữ trong bảng số liệu tối thiểu `12pt` (không thu nhỏ xuống `< 12pt` trừ khi người dùng yêu cầu riêng cho ghi chú phụ).
- **Bảo vệ thành phần thể thức bắt buộc:**
  - Tuyệt đối không xóa, gộp nhầm hay làm biến dạng bảng/khối **Quốc hiệu & Tiêu ngữ**, **Tên cơ quan ban hành**, **Nơi nhận**, và **Khối ký quyền hạn (`TM.`, `KT.`, `Q.`, `TL.`, `TUQ.`)**.
  - Các bảng bố cục không viền phục vụ Quốc hiệu/Chữ ký phải giữ nguyên vị trí và nội dung trừ khi người dùng yêu cầu rõ ràng bằng văn bản và yêu cầu đó hợp lệ.

### 1.2 Văn bản của Đảng (QĐ 399-QĐ/TW & HD 05-HD/VPTW)
- **Tuyệt đối không áp dụng Nghị định 30/2020/NĐ-CP** hoặc thay tiêu ngữ *ĐẢNG CỘNG SẢN VIỆT NAM* thành Quốc hiệu nhà nước.
- **Lề & định dạng văn bản chuẩn HD 05-HD/VPTW:**
  - Khổ A4, lề chuẩn trên `2,0 cm`, dưới `2,0 cm`, trái `3,0 cm`, phải `1,5 cm` (chú ý tôn trọng lề sẵn có của mẫu nếu yêu cầu chỉ giới hạn ở một số đoạn văn).
  - Đoạn văn nội dung chính: căn đều hai bên, giãn dòng cố định (**Exact**) trong khoảng **`18–22pt`** (thường dùng `18–20pt`), khoảng cách sau đoạn tối thiểu **`6pt`**.
  - Phông `Times New Roman` cỡ `14pt` cho thân văn bản; giữ nguyên khối Nơi nhận và Tiêu đề Đảng.

### 1.3 Văn bản Quy phạm Pháp luật (Nghị định 187/2025/NĐ-CP)
- Khổ A4; lề trên `2,0 cm`, lề dưới `2,0 cm`, lề trái `3,0–3,5 cm`, lề phải `1,5–2,0 cm`.
- Giữ nguyên cấu trúc Điều / Khoản / Điểm, phụ lục đính kèm và phân trang/khổ ngang của các biểu mẫu kèm theo thông tư/nghị định.

### 1.4 Tài liệu Học thuật, Đề cương, Luận văn & Biểu mẫu Cơ quan Tiếp nhận
- **Ưu tiên tối đa mẫu sẵn có của nơi tiếp nhận** (khoa/trường đại học, hội đồng bảo vệ, mẫu báo cáo thực tập). Không tự ý áp lề NĐ30, Quốc hiệu hay thay phông/cỡ chữ khác khi không được yêu cầu.
- Khi tài liệu đã định sẵn cỡ chữ tiêu đề/dữ liệu của trường (ví dụ `13pt` in đậm hoặc `14pt`), giữ nguyên thông số chữ đó nếu chỉ được yêu cầu chỉnh độ rộng, căn giữa hoặc lề trong ô.

---

## 2. Quy tắc An toàn Biên (*Fail-Closed*) vs. Trường hợp Hợp lệ Gần Ranh giới

### 2.1 Khi nào BẮT BUỘC Từ chối (Không sửa tệp gốc)
Trả về trạng thái `"refused"` (kèm lý do cụ thể) và **không xuất hay ghi đè tệp đã sửa** nếu gặp bất kỳ điều kiện nào sau đây:
1. **Tài liệu có chữ ký số OOXML:** Bên trong gói `.docx` tồn tại thành phần chữ ký số trong `_xmlsignatures/` (chỉnh sửa nội dung XML sẽ làm vô hiệu chữ ký xác thực).
2. **Xung đột thể thức bắt buộc:** Yêu cầu ép văn bản Đảng sang thể thức Nghị định 30/2020/NĐ-CP, xóa tiêu ngữ Đảng, hoặc thay đổi sai bản chất pháp lý của văn bản.
3. **Yêu cầu phá hủy khối thể thức bắt buộc:** Yêu cầu xóa bỏ hoàn toàn bảng Quốc hiệu/Tiêu ngữ hoặc khối Nơi nhận/Chữ ký ở văn bản hành chính chính thức.
4. **Yêu cầu không an toàn / chèn mã thực thi:** Yêu cầu nhúng macro VBA, kịch bản tự thực thi khi mở tệp, hoặc biến dạng cấu trúc gói OOXML trái chuẩn.
5. **Ảnh chụp cắt quá mờ/mơ hồ:** Có ảnh chụp màn hình đính kèm để định vị nhưng trong tài liệu có `≥ 2` bảng/ảnh giống hệt nhau về cấu trúc và nội dung mà ảnh bị cắt mất tiêu đề, cột phân biệt hoặc số trang nên không thể xác định duy nhất đối tượng đích.

### 2.2 Khi nào KHÔNG Được Từ chối Oan (*Near-Boundary Valid*)
- Nếu tài liệu có bảng bố cục (Quốc hiệu, Nơi nhận, Chữ ký) nằm cạnh **một bảng số liệu/chấm điểm/mức thu rõ ràng**, và người dùng chỉ yêu cầu chỉnh riêng bảng số liệu đó (hoặc đính kèm ảnh cắt có tiêu đề cột/nội dung phân biệt được duy nhất bảng đích): **hãy thực hiện chỉnh sửa đúng bảng đích** và giữ nguyên hoàn toàn các bảng bố cục còn lại.
- Nếu tệp nguồn có sẵn cảnh báo tương thích từ phần mềm khác (Word for Mac, ONLYOFFICE, Google Docs, LibreOffice) nhưng **không** có chữ ký số và yêu cầu hợp lệ: vẫn xử lý bình thường và bảo toàn tối đa các không gian tên/khối tương thích mở rộng có sẵn.

---

## 3. Checklist Bảo toàn OOXML & Bố cục khi Chỉnh sửa

1. **Vừa khít vùng in thực tế của từng phân đoạn (section):**
   - Luôn tính bề rộng khả dụng = *chiều rộng trang - lề trái - lề phải* của đúng section chứa bảng/ảnh đích (đặc biệt lưu ý tài liệu pha trộn trang dọc và trang ngang).
   - Không để bảng hoặc ảnh tràn vượt mép phải vùng in; với bảng căn giữa, xóa/đưa thụt lề trái bảng về `0` để bảng không bị lệch tâm.
2. **Đồng bộ đầy đủ kích thước bảng và ô gộp:**
   - Khi thay đổi độ rộng bảng hoặc từng cột, cập nhật đồng bộ cả độ rộng tổng bảng, lưới cột và độ rộng từng ô (tổng độ rộng cột con phải khớp chính xác với độ rộng ô gộp ngang), giữ nguyên tổng chiều ngang nhất quán giữa các hàng.
   - Khi đặt lề trong ô (cell margins) đồng nhất cho bảng, nhớ xóa các giá trị lề trong ô ghi đè cục bộ ở cấp từng ô nếu được yêu cầu đồng nhất toàn bảng.
   - Không chỉnh trực tiếp mã XML bằng biểu thức chính quy (regex) dễ làm gãy thứ tự phần tử con theo chuẩn WordprocessingML; luôn giữ đúng thứ tự phần tử thuộc tính bảng, hàng, ô, đoạn và chữ.
3. **Ảnh & Đối tượng đồ họa:**
   - Khi thu phóng ảnh hoặc khung hình minh họa, tính lại chiều cao theo đúng tỷ lệ `height / width` gốc để không làm méo hình; giữ nguyên các ảnh và khối đồ họa tương thích khác.
4. **Bảo toàn tuyệt đối phần ngoài mục tiêu:**
   - Không thêm, xóa hoặc làm biến đổi chữ nội dung tiếng Việt (giữ chuẩn Unicode NFC).
   - Không thay đổi các bảng, ảnh, đoạn văn, phân đoạn (section dọc/ngang), đầu trang/chân trang (header/footer), chú thích chân trang (footnote/endnote), mục lục/trường tự động hay các khai báo không gian tên ngoài phạm vi người dùng chỉ định.
5. **Kiểm tra trực quan sau chỉnh sửa:**
   - Xuất sang PDF/PNG (ví dụ qua LibreOffice headless) và kiểm tra bố cục thực tế (không tràn lề, không rớt chữ ký lẻ trang, đúng vị trí mục tiêu) trước khi bàn giao.
