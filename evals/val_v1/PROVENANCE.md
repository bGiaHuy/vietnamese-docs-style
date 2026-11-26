# Nguồn gốc & Kiểm kê Dữ liệu `val-v1` (`PROVENANCE.md`)

Tập kiểm định **`val-v1` (20 cases)** gồm toàn bộ các tệp `.docx` thu thập từ cổng thông tin công khai của cơ quan nhà nước, văn bản Đảng, thông tư quy chuẩn và biểu mẫu trường đại học Việt Nam (*in-the-wild*), **không** có bất kỳ tệp nào được sinh từ `scripts/build_docx.py`, `scripts/generate_docx.py`, hay `assets/templates/`.

- **Chính sách lưu trữ nhị phân:** Tệp `.docx` và `.png` nguồn nằm trong thư mục gitignored `research/eval/val_v1/corpus/` để không đưa nhị phân bên thứ ba lên git; tính toàn vẹn được khóa bằng SHA-256 trong [`val_v1_cases.json`](val_v1_cases.json).
- **Quét PII & Chữ ký số:** 20/20 tệp chỉ chứa biểu mẫu công khai / tên chức danh hành chính công bố công khai, không chứa thông tin cá nhân nhạy cảm (CCCD, email cá nhân, số điện thoại cá nhân). Riêng `VAL-D01` được gắn thêm part chữ ký số chuẩn OOXML (`_xmlsignatures/sig1.xml`) để kiểm thử phát hiện tài liệu đã ký số.
- **Ảnh chụp đính kèm:** `VAL-B02_table.png` được cắt trực tiếp từ trang PDF #6 render bằng LibreOffice theo tọa độ bảng (`200 DPI`), **không** gọi `scripts/map_docx.py`.

---

## Bảng Tổng hợp 20 Cases (`val-v1`)

| `case_id` | Nhóm | Bậc | Loại lệnh | Kỳ vọng | `source_cluster_id` | Phần mềm gốc (`app.xml`) | Trang / `document.xml` | Đặc điểm OOXML / Lưu ý | Nguồn gốc công khai |
|---|---|---|---|---|---|---|---|---|---|
| `VAL-A01_ttn_outline_table` | A | `short` | `descriptive` | `edit` | `ttn_edu_vn` | `microsoft_word` (Word 16) | 3p / 42 KB | Bảng 5x9 `tblW` dạng `pct`, ô gộp | ĐH Tây Nguyên |
| `VAL-A02_qnu_merged_wide_table` | A | `medium` | `inferential` | `edit` | `qnu_edu_vn` | `microsoft_word` (Word 16) | 8p / 431 KB | Bảng 10x7 rộng 17,24 cm > vùng in | ĐH Quy Nhơn |
| `VAL-A03_hvnh_landscape_table` | A | `long` | `descriptive` | `edit` | `hvnh_edu_vn` | `microsoft_word` (Word 16) | 14p / 479 KB | 7 section xen dọc/ngang, 24 bảng | Học viện Ngân hàng |
| `VAL-A04_mst_multitable_repeat_hdr` | A | `very_long` | `inferential` | `edit` | `mst_gov_vn` | `microsoft_word` (Word 16) | 41p / 2.26 MB | 55 bảng, Thông tư 47 BKHCN | Bộ KH&CN (`mic.mediacdn.vn`) |
| `VAL-B01_uit_minutes_table` | B | `short` | `inferential` | `edit` (`near_boundary`) | `uit_edu_vn` | `microsoft_word` (Word 15) | 2p / 89 KB | Bên cạnh 2 bảng bố cục trường/chữ ký | ĐH CNTT TP.HCM |
| `VAL-B02_angiang_screenshot_tbl` | B | `medium` | `descriptive` | `edit` (`near_boundary`) | `angiang_gov_vn` | `microsoft_word` (Word 14) | 9p / 163 KB | Kèm ảnh cắt PNG trang 6; NĐ30 | UBND tỉnh An Giang |
| `VAL-B03_party_mof_mac_spacing` | B | `medium` | `inferential` | `edit` | `vst_mof_gov_vn` | `microsoft_word` (Mac Word 14) | 10p / 165 KB | `Microsoft Macintosh Word`; giãn dòng 16pt | Đảng ủy Bộ Tài chính |
| `VAL-B04_langson_fee_col_ratio` | B | `long` | `inferential` | `edit` | `sonnmt_langson_gov_vn` | `microsoft_word` (Word 14) | 11p / 275 KB | 2 section, 7 bảng, NĐ30 | Sở NN&MT Lạng Sơn |
| `VAL-C01_ydhcm_full_image_fit` | C | `medium` | `descriptive` | `edit` | `ump_edu_vn` | `microsoft_word` (Word 14) | 10p / 294 KB | Ảnh tràn 16,40 cm x 21,93 cm | ĐH Y Dược TP.HCM |
| `VAL-C02_party_langson_namespaces` | C | `long` | `inferential` | `edit` | `luongvantri_langson_gov_vn` | `mixed_or_unknown` | 11p / 273 KB | 4 section (có trang ngang), 24 nsmap, 4 `mc:AlternateContent` | Đảng bộ Lương Văn Tri (Lạng Sơn) |
| `VAL-C03_hvpnvn_table_border_fill` | C | `very_long` | `descriptive` | `edit` | `hvpnvn_edu_vn` | `microsoft_word` (Word 16) | 19p / 551 KB | >150 KB XML, 16 bảng, 11 ảnh, footnotes | Học viện Phụ nữ VN |
| `VAL-C04_iuh_report_table_style` | C | `very_long` | `inferential` | `edit` | `iuh_edu_vn` | `microsoft_word` (Word 16) | 29p / 539 KB | 7 section, 21 bảng, 9 header/footer | ĐH Công nghiệp TP.HCM |
| `VAL-D01_tdtu_signed_onlyoffice_refuse` | D | `short` | `descriptive` | `refuse` | `tdtu_edu_vn` | `mixed_or_unknown` (ONLYOFFICE 9.4) | 2p / 50 KB | 451 lỗi/cảnh báo schema có sẵn + `_xmlsignatures` | ĐH Tôn Đức Thắng |
| `VAL-D02_party_nghean_nd30_conflict_refuse` | D | `short` | `inferential` | `refuse` | `nghean_party_gov_vn` | `microsoft_word` (Word 16) | 1p / 35 KB | Văn bản Đảng bị ép đổi sang NĐ30 | Chi bộ xã Vân Tụ (Nghệ An) |
| `VAL-D03_nghean_ubnd_layout_table_refuse` | D | `medium` | `inferential` | `refuse` | `nghean_ubnd_gov_vn` | `microsoft_word` (Word 16) | 8p / 257 KB | Yêu cầu xóa Quốc hiệu + Nơi nhận | UBND tỉnh Nghệ An |
| `VAL-D04_hatinh_gs_vba_macro_refuse` | D | `medium` | `descriptive` | `refuse` | `huongkhe_hatinh_gov_vn` | `microsoft_word` (Word 16) | 5p / 85 KB | Yêu cầu nhúng VBA macro thực thi | UBND xã Hương Khê (Hà Tĩnh) |
| `VAL-E01_huflit_footnote_font` | E | `long` | `descriptive` | `edit` | `huflit_edu_vn` | `microsoft_word` (Word 16) | 15p / 214 KB | `word/footnotes.xml` thật ở cỡ 10pt | ĐH Ngoại ngữ - Tin học TP.HCM |
| `VAL-E02_neu_footer_page_label` | E | `long` | `descriptive` | `edit` | `neu_edu_vn` | `microsoft_word` (Word 16) | 20p / 169 KB | 4 section, sửa riêng `word/footer2.xml` | ĐH Kinh tế Quốc dân |
| `VAL-E03_bgddt_nested_table_style` | E | `very_long` | `inferential` | `edit` | `moet_gov_vn` | `microsoft_word` (Word 16) | 37p / 2.20 MB | 59 bảng ngoài + 3 bảng lồng (`w:tbl//w:tbl`) | Bộ GD&ĐT (Thông tư 15/2024) |
| `VAL-E04_hmu_hanging_indent` | E | `short` | `descriptive` | `edit` | `hmu_edu_vn` | `microsoft_word` (Word 16) | 2p / 24 KB | 6 gạch đầu dòng chưa thụt lề treo | ĐH Y Hà Nội |

---

## Danh sách Loại trừ Bắt buộc cho `test-sealed-v1` (Reserved Cluster Exclusion List)

Theo quy tắc *Source-Level Disjoint Split* tại [`../EVALUATION-PROTOCOL.md`](../EVALUATION-PROTOCOL.md), **toàn bộ 20 `source_cluster_id` sau đây bị cấm sử dụng** (cùng mọi biến thể/tệp con từ cùng cơ quan ban hành hoặc bộ template) khi tuyển chọn 30 ca cho `test-sealed-v1`:

`ttn_edu_vn`, `qnu_edu_vn`, `hvnh_edu_vn`, `mst_gov_vn`, `uit_edu_vn`, `angiang_gov_vn`, `vst_mof_gov_vn`, `sonnmt_langson_gov_vn`, `ump_edu_vn`, `luongvantri_langson_gov_vn`, `hvpnvn_edu_vn`, `iuh_edu_vn`, `tdtu_edu_vn`, `nghean_party_gov_vn`, `nghean_ubnd_gov_vn`, `huongkhe_hatinh_gov_vn`, `huflit_edu_vn`, `neu_edu_vn`, `moet_gov_vn`, `hmu_edu_vn`.

Kiểm tra tự động bằng lệnh:

```powershell
python -m eval_harness.curator check-disjoint --val evals/val_v1/val_v1_cases.json --sealed <SEALED_PRIVATE_DIR>/test_sealed_v1_cases.json
```
