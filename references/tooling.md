# Công cụ và phạm vi kiểm tra

Python 3.10+; cài `pip install -e .` (hoặc `pip install -r requirements.txt` khi giải nén trực tiếp từ `vietnamese-docx.zip` không có `pyproject.toml`), kiểm thử bằng `pip install -e ".[dev]"` (hoặc `requirements-dev.txt`).
Sau khi cài đặt editable/package, gọi trực tiếp `vndocs <subcmd>`; khi chạy từ gói ZIP độc lập, thay `vndocs <subcmd>` bằng `python scripts/cli.py <subcmd>` (hỗ trợ đầy đủ: `doctor|generate|inspect|validate|render|fill|edit|demo|templates|map|layout-edit|diff-layout|paginate|package`) hoặc gọi riêng `python scripts/<tool>.py`.
Kiểm tra nhanh môi trường và LibreOffice renderer: `python scripts/cli.py doctor` (hoặc `vndocs doctor`).
LibreOffice là dependency hệ thống riêng để render. Trên Windows có thể chỉ
định `--renderer` hoặc biến `VNDOCS_RENDERER`; không yêu cầu cài đè phiên bản sẵn.

## Mục lục

- Tạo mới và giữ mẫu
- Cấu hình quy tắc
- Kiểm tra và render
- Chỉnh sửa bố cục có kiểm soát
- Mẫu kiểm tra tái lập
- Mẫu không có token
- Starter đi kèm và đóng gói
- Số trang và kiểm tra phát triển

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
đã xem mọi trang. `--png` cần PyMuPDF; thêm `--annotate-marks` để xuất thêm ảnh
`page-N-marks.png` có khung nhãn `[T#]`, `[T#:LAYOUT]`, `[IMG#]` phục vụ định vị
từ screenshot mà không làm biến dạng PDF sạch. Thiếu renderer hoặc PNG chưa tạo
được trả code 2; lỗi chuyển đổi trả code 1. Kiểm tra ảnh thật trước khi nghiệm thu bố cục.

## Chỉnh sửa bố cục có kiểm soát

Luồng chỉnh sửa bố cục trên `.docx` hiện có giữ nguyên cấu trúc ngoài vùng đích (`xml_ntpr == 1.0`):

1. `vndocs map src.docx [--format markdown|json] [--with-pdf | --pdf-path file.pdf] [--include-paragraphs]`
   Lập bản đồ section, bảng (`T#hash`), ảnh (`IMG#hash`) và đoạn (`P#hash`). Bảng bố cục (quốc hiệu, nơi nhận, chữ ký) được gắn `[LAYOUT:<role>]` và khóa mặc định.
2. `vndocs render src.docx --output-dir out --png [--annotate-marks] [--renderer word-win] [--cache-dir .vndocs-cache]`
   Đọc ảnh chụp người dùng và đối chiếu chữ/số dòng/cột với `vndocs map` trước; chỉ chạy `--annotate-marks` để xuất `page-N-marks.png` có nhãn Set-of-Mark `[T#]` / `[IMG#]` khi còn `≥ 2` ứng viên tương tự hoặc dữ liệu chữ từ `map` chưa đủ chốt. Hỗ trợ `--renderer word-win` (Word COM trên Windows), `--cache-dir` lưu cache PDF theo `SHA-256` kèm `render-manifest.json` chống stale render. Nếu có `≥ 2` bảng/ảnh giống hệt nhau mà ảnh bị crop không có tiêu đề/số trang để phân biệt, dừng lại hỏi người dùng thay vì đoán.
3. `vndocs layout-edit src.docx dst.docx --plan plan.json [--dry-run] [--allow-layout-tables] [--profile <name>] [--source-sha256 <hex>] [--track-changes] [--author <name>]`
   `plan.json` có dạng `{"operations": [...]}` (hoặc mảng JSON) với 13 op hợp lệ:
   - `table.scale_width`: `{op, target, factor | target_width_pct | target_width_cm}`
   - `table.set_col_widths`: `{op, target, widths: ["4cm", "5cm", ...]}` (đủ `grid_cols` cột)
   - `table.set_alignment`: `{op, target, align: "left|center|right"}`
   - `table.set_cell_margins`: `{op, target, top_cm?, bottom_cm?, left_cm?, right_cm?, clear_cell_overrides?}`
   - `table.set_font_size`: `{op, target, size_pt, preserve_header_size?}`
   - `table.set_pagination`: `{op, target, repeat_header?, header_rows?, cant_split?}`
   - `table.set_borders`: `{op, target, style?: "single|none|nil|double|dashed|dotted", size_eighth_pt?, color?, sides?}`
   - `table.set_shading`: `{op, target, fill: "<6-hex>|clear", scope?: "header|all|odd_rows|even_rows|row", row_index?}`
   - `image.scale`: `{op, target: "IMG...", factor | width_cm | width_pct, align?}` (giữ tỷ lệ khung hình, đồng bộ `wp:extent` + `a:ext`)
   - `paragraph.set_format`: `{op, target | targets, align?, space_before_pt?, space_after_pt?, line_spacing? | line_spacing_exact_pt?, first_line_indent_cm?, left_indent_cm?, right_indent_cm?, keep_with_next?}`
   - `paragraph.merge_runs`: `{op, target | targets}` (gộp các run văn bản thuần liền kề cùng `rPr`, giữ nguyên ký tự/trường)
   - `style.update`: `{op, style_name, font?, size_pt?, color?, bold?, italic?, align?, space_before_pt?, space_after_pt?, line_spacing? | line_spacing_exact_pt?, first_line_indent_cm?}` (sửa `word/styles.xml`, không chạm `word/document.xml`)
   - `section.set_layout`: `{op, section_index?, orientation?: "portrait|landscape", page_size?: "A4|Letter", margins_cm?: {top?, bottom?, left?, right?, header?, footer?}}`
4. `vndocs diff-layout src.docx dst.docx [--target <anchor>] [--output-dir out] [--crop-dir crops] [--renderer ...] [--cache-dir ...]`
   Đo lại kích thước bảng/ảnh trên PDF sau render (hỗ trợ `--cache-dir` và `compare_renderers`), phát hiện tràn lề phải (`overflow_right_mm`), kiểm tra nhảy trang ngoài vùng đích (`non_target_page_shifts`) và xuất crop PNG trước/sau của target (`max_iter = 2`).
5. `vndocs paginate src.docx dst.docx [--target-pages N] [--fix-orphan-signature] [--profile <name>] [--max-steps 5] [--renderer ...] [--cache-dir ...]`
   Co gọn bố cục theo bậc an toàn (lề ô bảng -> ảnh lớn -> cỡ chữ bảng theo sàn profile -> giãn dòng/đoạn không phải tiêu đề -> lề trên/dưới nếu profile cho phép), giữ nguyên 100% câu chữ và bảng bố cục NĐ30/Đảng, giúp gom chữ ký rơi lẻ sang trang cuối hoặc ép vừa số trang mục tiêu.

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
