# Vietnamese Docs Style (vn-report-pro)

Kho lưu trữ chuyên môn dành cho AI Agent để tạo, định dạng, và biên tập các tài liệu tiếng Việt đạt chuẩn, bao gồm văn bản hành chính theo Nghị định 30/2020/NĐ-CP, báo cáo học thuật, đề xuất dự án, và các loại biên bản.

## Điểm nổi bật
- **Phân loại Profile rõ ràng**: Hỗ trợ định dạng chính xác theo ngữ cảnh (administrative, academic, proposal, minutes-administrative, minutes-general).
- **Tuân thủ NĐ30**: Đảm bảo khổ giấy A4, font Times New Roman, lề chuẩn (trái 30mm, phải 15mm), và đầy đủ các thành phần thể thức đối với văn bản hành chính.
- **Biên tập ngôn ngữ (Editorial Quality)**: Các hướng dẫn và pattern cụ thể để văn bản tiếng Việt tự nhiên, rõ ràng, tránh các lỗi phổ biến (bị động dư thừa, từ nối đơn điệu).
- **Executable Validation**: Kèm theo script xác thực (`validate_docx.py`) và tạo DOCX (`build_docx.py`) giúp AI dễ dàng kiểm tra tính đúng đắn trước khi bàn giao cho người dùng.

## Cấu trúc Repository

```text
├── SKILL.md                          # Entry point dành cho AI Agent
├── assets/
│   └── samples/                      # Các file mẫu .docx 
├── references/                       # Tài liệu chuyên môn (Agent đọc khi cần)
│   ├── academic-report.md            # Báo cáo học thuật
│   ├── administrative-format-nd30.md # Chuẩn NĐ30 cho VB hành chính
│   ├── document-profiles.md          # Bảng định tuyến profile tài liệu
│   ├── editorial-quality-vi.md       # Hướng dẫn biên tập văn phong
│   ├── meeting-minutes.md            # Các loại biên bản
│   ├── style-spec.md                 # Blueprint cho python-docx
│   ├── validation-checklist.md       # Danh sách kiểm tra chất lượng
│   └── source-documents/             # Tài liệu nguồn và tham chiếu cũ
├── scripts/                          # Script hỗ trợ Agent
│   ├── build_docx.py                 # Hàm helper tạo file DOCX
│   ├── render_docx.py                # Rendering stub
│   └── validate_docx.py              # Script kiểm tra cấu trúc file DOCX
└── tests/                            # Pytest fixtures và tests
    └── test_validation.py            
```

## Dành cho AI Agent

Hãy đọc trực tiếp file `SKILL.md` để hiểu quy trình làm việc, cách đọc các reference, và các chuẩn mực ưu tiên.

## Dành cho Người Dùng

Để sử dụng skill này, bạn chỉ cần yêu cầu AI Agent của mình:
- "Tạo một thông báo họp theo chuẩn NĐ30."
- "Viết báo cáo học thuật về AI, có trang bìa."
- "Chỉnh sửa lại file DOCX này sao cho đúng thể thức văn bản hành chính."

Hệ thống sẽ tự động đối chiếu các quy chuẩn trong kho lưu trữ này và áp dụng chính xác.
