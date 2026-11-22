# Công cụ và phạm vi kiểm tra

Python 3.10+; cài `pip install -e .` (hoặc `requirements.txt`), kiểm thử bằng `pip install -e ".[dev]"` (hoặc `requirements-dev.txt`).
Sau khi cài đặt editable/package, có thể gọi trực tiếp qua lệnh `vndocs <doctor|generate|inspect|validate|render|fill|edit|demo|templates|package>` hoặc chạy `python scripts/<tool>.py`.
Kiểm tra nhanh môi trường và LibreOffice renderer: `vndocs doctor`.
LibreOffice là dependency hệ thống riêng để render. Trên Windows có thể chỉ
định `--renderer` hoặc biến `VNDOCS_RENDERER`; không yêu cầu cài đè phiên bản sẵn.

## Tạo mới và giữ mẫu

`build_docx.setup_document(doc, profile, config=None)` áp cấu hình lên tài liệu
mới. Profile `office` cho thông báo/ghi nhớ/báo cáo nội bộ chung; học thuật/văn phòng không bắt buộc lề hành chính; custom không có
config giữ nguyên. Các helper thêm nội dung kế thừa style thay vì gán font trực
tiếp. `add_numbered(doc, items, start=1)` tạo danh sách Word thật, một cấp;
numbering nhiều cấp theo mẫu vẫn cần giữ hoặc cấu hình bằng công cụ phù hợp.

Không chạy `setup_document` hay `generate_docx.py` lên template cần bảo toàn.

### Sinh DOCX từ JSON (`generate`) và soi/chuẩn hóa mẫu (`inspect`)

```bash
python scripts/generate_docx.py assets/schemas/examples/nd30-dispatch.json output.docx --validate
python scripts/generate_docx.py - -o output.docx --validate < input.json
python scripts/generate_docx.py --json-string '{"profile":"party","body":[{"type":"paragraph","text":"Nội dung"}]}' -o output.docx
python scripts/inspect_docx.py template.docx --normalize-output clean-template.docx
```

- `generate_docx.py`: nhận file JSON, `-` (`stdin`) hoặc `--json-string` đúng theo `assets/schemas/document-input.schema.json`, hỗ trợ đa profile (`administrative`, `party`, `academic`, `office`, `proposal`, `minutes-general`, `minutes-administrative`, `custom`), tùy chọn `--config` (ví dụ `assets/rules/legal-normative.json`), `--validate` và `--draft`. Output bắt buộc có đuôi `.docx`.
- `inspect_docx.py`: báo cáo SHA256, chữ ký số, kích thước/lề từng section, style đã dùng, font/màu kế thừa, các loại placeholder (`{{...}}`, `[CẦN BỔ SUNG: ...]`, `???`, `<...>`, chuỗi dấu chấm dài) và phát hiện ký tự `Ð` (`U+00D0`), `ð` (`U+00F0`) hoặc lỗi `kính giửi`. Tùy chọn `--normalize-output` vá trực tiếp `<w:t>` theo kiểu giữ nguyên byte các ZIP member khác qua `docx_package` (từ chối file đã ký số).

Điền token `{{name}}`:

```bash
python scripts/fill_template.py template.docx output.docx --values values.json
```

JSON là map tên token sang chuỗi một dòng. Hỗ trợ body, bảng lồng, header/footer
và token chia qua nhiều run; phần thay thế lấy định dạng run đầu. Giữ run xung
quanh, section và style. Từ chối ghi đè nguồn, token thiếu/giá trị không dùng,
hoặc vùng token có field/hyperlink/control/drawing/revision. Tab/break ngoài vùng sửa được giữ; không thay text xuyên qua chúng.
`--draft` cho phép token thiếu, không bỏ qua giới hạn cấu trúc. Mẫu không có
token cần chỉnh sửa có chủ đích bằng công cụ Word/OOXML, không tự tìm dấu chấm
trống rồi thay tất cả.

## Cấu hình quy tắc

Ví dụ `assets/rules/custom-example.json`. Dùng cùng config khi tạo mới và kiểm tra.
Các khóa: `profile`, `source`, `font`, `color`, `body_size_pt`, `line_spacing`,
`page_cm`, `allow_landscape`, `margins_cm`, `sections`, `style_rules`, `required_text`,
`forbid_bullets`, `forbid_shading`, `line_spacing_exact_pt`, `page_number_size_pt`.
`line_spacing_exact_pt` (null hoặc số pt) đặt giãn dòng Exactly cho Normal, ưu tiên hơn
`line_spacing`; `page_number_size_pt` là cỡ chữ field PAGE của profile có số trang.
Profile `party` mặc định: lề 2,0/2,0/3,0/1,5 cm, body 14pt, Exactly 20pt, số trang
14pt cách mép trên 1 cm, ẩn trang đầu (HD 05-HD/VPTW). Lề mặt sau khi in hai mặt
(trái 15, phải 30 mm) chưa được tạo tự động. VBQPPL dùng
`--profile administrative --config assets/rules/legal-normative.json`. Không rõ khóa báo lỗi, không âm thầm bỏ qua.

`body_size_pt` và `line_spacing` cấu hình generator, chưa phải kiểm tra validator. `style_rules` ánh xạ tên paragraph style sang `font`, `color`, `size_pt: [min,max]`, `bold`, `italic`; generator thiết lập style, validator kiểm tra role được dùng, gồm override trực tiếp. Các thuộc tính chưa hỗ trợ vẫn ghi unchecked.
`font`/`color` null bỏ kiểm tra đó. `page_cm` null bỏ kiểm tra khổ. Lề dùng
`[min,max]` cm; các side không khai báo không được kiểm tra. `sections` là danh
sách theo thứ tự, nếu khai báo phải khớp số section; mỗi mục override khổ/hướng/lề.
`required_text` chỉ kiểm tra literal có xuất hiện, không hiểu ý nghĩa hoặc vị trí.
Bullet/shading chỉ bị cấm khi config yêu cầu, không phải lệnh cấm chung từ NĐ30.

## Kiểm tra và render

```bash
python scripts/validate_docx.py output.docx --profile custom --config rules.json --report report.json
python scripts/render_docx.py output.docx --png
```

Validator kiểm tra từng section, font/màu trực tiếp và chuỗi style kế thừa,
body/bảng lồng/header/footer, placeholder chia run, required literal và bullet
Word khi cấu hình cấm và tính toàn vẹn tham chiếu numbering. Chuỗi số thực tế và liên kết numbering style phức tạp chưa được kết luận. Không chứng nhận pháp lý, sự thật, bố cục hoặc numbering
liên tục. Chưa xử lý đầy đủ conditional table style, theme, text box, revision,
footnote, chart và field. Báo `unchecked` để người dùng biết phần chưa kiểm tra.

Status `pass`: các kiểm tra đã hỗ trợ không phát hiện lỗi; `fail`: có lỗi;
`partial`: thiếu quy tắc custom hoặc không giải được một số thuộc tính. Exit
code tương ứng 0/1/2. `--draft` chuyển placeholder thành cảnh báo. Boolean
`validate_doc` chỉ giữ tương thích, muốn biết kiểm tra thiếu phải dùng `inspect_doc`.

Render dùng profile LibreOffice tạm riêng và output chuyển đổi mới; không nhận
PDF cũ là thành công. Status `rendered` chỉ nghĩa đã tạo file, chưa phải người
đã xem mọi trang. `--png` cần PyMuPDF. Thiếu renderer hoặc PNG chưa tạo được trả
code 2; lỗi chuyển đổi trả code 1. Kiểm tra ảnh thật trước khi nghiệm thu bố cục.

## Mẫu kiểm tra tái lập

```bash
python scripts/create_demo.py --output-dir artifacts/demo
python scripts/render_docx.py artifacts/demo/office-filled.docx --png
python -m pytest -q
```

Demo là mẫu nội bộ do dự án viết với dữ liệu giả có nhãn, gồm trang dọc/trang
ngang, bảng, numbering và token chia run; không đại diện quy định của cơ quan.

## Mẫu không có token

Đọc và render bản nguồn trước, lưu hash và slot map. Sửa text ở vị trí đã xác minh:

```bash
python scripts/edit_docx.py source.docx output.docx --edits slots.json --source-sha256 HASH
```

Mỗi slot có `part` (ví dụ word/document.xml), `paragraph` (index 0-based theo
`docx_package.paragraphs`, gồm paragraph trong bảng), `before` và `after`.
Text before phải xuất hiện đúng một lần trong paragraph; chuỗi rỗng chỉ dùng
khi paragraph thực sự trống. Không thay theo số paragraph của Document.paragraphs
vì nó bỏ qua bảng. Không tự áp slot map của mẫu khác hoặc bản nguồn đã đổi.

Cả fill và edit patch trực tiếp XML trong gói ZIP; mọi ZIP member không sửa
được giữ nguyên byte. Run xung quanh, style, numbering, relationship và phần
opaque giữ nguyên. Không có nghĩa layout sẽ giống khi nội dung dài hơn: cần
render/so sánh sau sửa. Từ chối chỉnh gói có chữ ký số vì sẽ làm chữ ký mất hiệu lực.

## Starter đi kèm và đóng gói

`assets/templates/manifest.json`: ba starter do dự án viết (biên bản nội bộ,
báo cáo học thuật ngắn, thông báo hành chính tham chiếu mẫu 1.4). Có `.rules.json`
và `.values.json` giả định để thử lại; note thử nghiệm là token, không bị ép vào
nội dung thật. Chọn theo `starter-templates.md`; đây không phải chuẩn mọi trường/cơ quan.

Tái tạo bản thử: `python scripts/create_templates.py --output-dir artifacts/template-suite`.
Đóng gói local: `python scripts/package_skill.py --output dist/vietnamese-docx.zip`.
Chỉ đóng gói starter đã xem ảnh và khớp hash; không kèm dữ liệu nghiên cứu,
file bên thứ ba chưa rõ quyền, runtime, test cache hoặc thông tin Git.

## Số trang và kiểm tra phát triển

Tài liệu mới profile hành chính có field PAGE ở giữa header, 13pt; ẩn ở trang
đầu tài liệu, không ẩn trang đầu của các section sau. `add_page_numbers` có thể
được gọi có chủ đích cho tài liệu mới khác; không dùng để reset header của mẫu.
LibreOffice/Word tính field khi render; custom giữ field và cần kiểm tra refresh.
Renderer có thể thay font khi máy thiếu font, nên tên font trong XML không thay
việc xem bản render.

`python -m ruff check scripts tests` và `python -m pytest -q` là kiểm tra local
được CI chạy trong repo nguồn (tests và dependency phát triển không nằm trong ZIP cài skill).
Để render PDF/PNG ở local, cài LibreOffice theo hệ điều hành:
- Windows: `winget install TheDocumentFoundation.LibreOffice`
- macOS: `brew install --cask libreoffice`
- Ubuntu/Debian: `sudo apt-get install -y libreoffice-writer fonts-liberation`
Nếu `soffice` không nằm trong `PATH`, truyền `--renderer` hoặc đặt biến môi trường `VNDOCS_RENDERER` trỏ tới file thực thi LibreOffice.
