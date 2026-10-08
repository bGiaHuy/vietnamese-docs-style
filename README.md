# Vietnamese Docs Style

**AI có thể tạo tài liệu Word. Nhưng liệu AI có thực sự hiểu cách người Việt Nam trình bày văn bản?**

`vietnamese-docs-style` là một AI Agent Skill hỗ trợ tạo, định dạng và biên tập tài liệu Word (`.docx`) theo các quy ước trình bày, cấu trúc và văn phong phù hợp với bối cảnh Việt Nam.

Dự án tập trung vào việc bản địa hóa quá trình tạo tài liệu bằng AI, thay vì chỉ dịch nội dung sang tiếng Việt.

## 1. Vấn đề đặt ra

Ngày nay, các AI Agent đã có khả năng tạo báo cáo, soạn thảo văn bản và xuất trực tiếp thành file Microsoft Word.

Tuy nhiên, **tạo được một file DOCX không đồng nghĩa với việc tạo được một tài liệu phù hợp với quy chuẩn trình bày tại Việt Nam.**

Nhiều hệ thống AI và công cụ tạo tài liệu sử dụng các mẫu định dạng phổ thông. Những mẫu này có thể phù hợp với một số quy ước quốc tế, nhưng chưa chắc đáp ứng yêu cầu của các cơ quan, tổ chức hoặc cơ sở giáo dục tại Việt Nam.

Chẳng hạn:

- **Văn bản hành chính:** Có quy định cụ thể về khổ giấy, căn lề, phông chữ, cỡ chữ, quốc hiệu, tiêu ngữ, số ký hiệu và các thành phần thể thức theo Nghị định 30/2020/NĐ-CP.
- **Báo cáo học thuật:** Thường có yêu cầu riêng về trang bìa, hệ thống đề mục, cách trình bày bảng biểu và cấu trúc nội dung.
- **Biên bản, tờ trình, đề xuất dự án:** Có mục đích sử dụng, bố cục và văn phong khác nhau.
- **Tài liệu theo mẫu riêng:** Phải ưu tiên quy định của trường học, cơ quan hoặc tổ chức cung cấp mẫu.

Một mẫu định dạng duy nhất không thể đáp ứng chính xác mọi tình huống.

Do đó, vấn đề không đơn thuần là làm sao để AI viết được tiếng Việt, mà là:

**Làm thế nào để AI hiểu và áp dụng đúng các quy ước trình bày tài liệu trong từng bối cảnh sử dụng tại Việt Nam?**

## 2. Từ vấn đề thực tế đến giải pháp

`vietnamese-docs-style` được xây dựng nhằm giải quyết khoảng cách giữa khả năng tạo tài liệu của AI và những yêu cầu trình bày thực tế tại Việt Nam.

Thay vì chỉ bổ sung một template Word cố định, dự án nghiên cứu và hệ thống hóa các quy tắc liên quan đến:

- Thể thức văn bản hành chính theo Nghị định 30/2020/NĐ-CP.
- Cấu trúc và cách trình bày báo cáo học thuật.
- Bố cục của các loại biên bản và đề xuất.
- Quy ước sử dụng tiếng Việt trong văn bản trang trọng.
- Những yêu cầu về căn lề, phông chữ, cỡ chữ, khoảng cách và tổ chức nội dung.

Các kiến thức này được tổ chức thành những tài liệu tham chiếu, hướng dẫn và công cụ để AI Agent có thể sử dụng lại khi tạo DOCX.

Mục tiêu là giúp AI không chỉ **viết đúng nội dung**, mà còn **lựa chọn cách trình bày phù hợp với loại tài liệu**.

## 3. Cách thức hoạt động

Skill triển khai quy trình gồm năm giai đoạn:

**Bước 1 — Nhận diện loại tài liệu**

AI Agent phân tích yêu cầu, xác định mục đích sử dụng và lựa chọn nhóm tài liệu tương ứng.

**Bước 2 — Lựa chọn quy tắc trình bày**

Agent đọc tài liệu tham chiếu phù hợp. Ví dụ, văn bản hành chính áp dụng hướng dẫn về Nghị định 30, còn báo cáo học thuật ưu tiên yêu cầu của cơ sở đào tạo.

**Bước 3 — Biên soạn nội dung**

Agent xây dựng nội dung theo cấu trúc và văn phong của loại tài liệu, đồng thời tránh tự tạo ra các dữ kiện chưa được cung cấp.

**Bước 4 — Tạo file DOCX**

Sử dụng các công cụ hỗ trợ dựa trên `python-docx` để thiết lập định dạng và xây dựng tài liệu Word.

**Bước 5 — Kiểm tra kết quả**

Kiểm tra các thuộc tính cấu trúc được hỗ trợ, phát hiện một số lỗi định dạng và rà soát tài liệu trước khi bàn giao.

## 4. Các loại tài liệu được hỗ trợ

| Nhóm tài liệu | Ví dụ |
|---|---|
| Văn bản hành chính | Công văn, quyết định, thông báo, tờ trình |
| Báo cáo học thuật | Báo cáo môn học, tiểu luận, báo cáo nghiên cứu |
| Đề xuất dự án | Kế hoạch, đề xuất triển khai |
| Biên bản hành chính | Biên bản sử dụng trong bối cảnh hành chính |
| Biên bản thông thường | Biên bản họp nhóm, biên bản nội bộ |
| Tài liệu tùy chỉnh | Tài liệu theo mẫu do người dùng cung cấp |

**Lưu ý:** Không phải mọi tài liệu tiếng Việt đều phải tuân theo Nghị định 30/2020/NĐ-CP. Skill phân biệt các nhóm tài liệu để lựa chọn quy tắc phù hợp.

## 5. Ví dụ sử dụng

Sau khi cài đặt skill vào AI Agent, người dùng có thể đưa ra những yêu cầu như:

> Tạo một công văn thông báo lịch họp, trình bày theo Nghị định 30/2020/NĐ-CP.

> Viết báo cáo môn học về trí tuệ nhân tạo, có trang bìa, mục lục và phân chia chương rõ ràng.

> Soạn biên bản họp nhóm bằng tiếng Việt, có nội dung thảo luận, kết luận và phân công công việc.

> Định dạng lại tài liệu này theo mẫu báo cáo của trường đại học mà tôi cung cấp.

## 6. Cấu trúc dự án

```text
vietnamese-docs-style/
├── SKILL.md
├── references/
│   ├── administrative-format-nd30.md
│   ├── academic-report.md
│   ├── document-profiles.md
│   ├── editorial-quality-vi.md
│   └── ...
├── scripts/
│   ├── build_docx.py
│   ├── validate_docx.py
│   └── render_docx.py
├── assets/
└── tests/
```

Trong đó:

- `SKILL.md`: Hướng dẫn quy trình dành cho AI Agent.
- `references/`: Cơ sở tri thức về định dạng, bố cục và văn phong.
- `scripts/`: Các công cụ hỗ trợ tạo và kiểm tra tài liệu.
- `assets/`: Tài nguyên và tài liệu mẫu.
- `tests/`: Các bài kiểm thử.

## 7. Giới hạn hiện tại

Dự án hỗ trợ AI tạo tài liệu có định dạng phù hợp hơn với các quy ước tại Việt Nam, nhưng không thay thế hoàn toàn bước kiểm tra của người sử dụng.

Cụ thể:

- Các script kiểm tra hiện tại chưa bao phủ mọi yêu cầu về thể thức và trình bày.
- Một số lỗi bố cục cần được kiểm tra trực quan trong Microsoft Word hoặc công cụ tương đương.
- Tài liệu hành chính vẫn cần được rà soát để bảo đảm phù hợp với quy định áp dụng và thẩm quyền ban hành.
- Chất lượng đầu ra phụ thuộc vào AI Agent, thông tin đầu vào và công cụ tạo DOCX được sử dụng.

## 8. Định hướng phát triển

Mục tiêu dài hạn của dự án là xây dựng một bộ quy tắc và công cụ bản địa hóa có thể tái sử dụng, giúp AI Agent làm việc hiệu quả hơn với tài liệu tiếng Việt.

Các hướng mở rộng có thể bao gồm việc hoàn thiện kiểm tra định dạng, cải thiện khả năng xử lý các template và bổ sung tài liệu mẫu có thể kiểm chứng.

---

**Phát triển bởi Bùi Gia Huy**

Dự án: [github.com/bGiaHuy/vietnamese-docs-style](https://github.com/bGiaHuy/vietnamese-docs-style)
