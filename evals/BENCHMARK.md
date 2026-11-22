# Public DOCX Validation Benchmark (`0/8` Baseline vs. `8/8` Skill-Guided)

Tài liệu này công bố bộ benchmark tái lập được cho công cụ tạo và kiểm tra `.docx` trong `vietnamese-docs-style` trên **16 trường hợp xác định (8 Baseline mặc định vs. 8 Skill-guided)**.

> **Phạm vi và giới hạn trung thực:** Đây là **benchmark deterministic của bộ công cụ DOCX** ([`scripts/validate_docx.py`](../scripts/validate_docx.py), [`scripts/build_docx.py`](../scripts/build_docx.py), [`scripts/fill_template.py`](../scripts/fill_template.py)), **không phải** bảng xếp hạng đo độ chính xác sinh văn bản tự do của từng mô hình LLM độc lập. Để xem kết quả đối chiếu văn bản thật do cán bộ Việt Nam soạn thảo và kiểm thử hành vi thủ công, xem [audit-comparison.md](audit-comparison.md), [manual-results.md](manual-results.md) và 14 tình huống tại [behavior-cases.json](behavior-cases.json).

---

## Tái lập trực tiếp trên máy của bạn

```bash
# In bảng Markdown ra terminal
python evals/run_validation_benchmark.py

# Xuất kết quả JSON thô
python evals/run_validation_benchmark.py --json
```

Mỗi lần push lên GitHub, `tests/test_validation.py` (`test_validation_benchmark_matches_published_json`) và workflow [.github/workflows/tests.yml](../.github/workflows/tests.yml) tự động chạy lại benchmark và đối chiếu khớp hoàn toàn với tệp kết quả công bố tại [benchmark-results.json](benchmark-results.json).

---

## Tổng hợp theo nhóm kiểm tra

| Nhóm kiểm tra (`validate_docx.py`) | Baseline mặc định | Đầu ra dùng Skill | Ý nghĩa thực tế |
| --- | :---: | :---: | --- |
| **Tổng tỷ lệ `pass`** | **0 / 8 (0%)** | **8 / 8 (100%)** | Bao phủ đủ hành chính NĐ30, Đảng, VBQPPL, học thuật, biên bản, văn phòng và custom config |
| **Khổ giấy & lề NĐ30** (`administrative`, `minutes-administrative`) | 0 / 3 pass | 2 / 2 pass | Bắt lỗi khổ US Letter (`21.59 × 27.94 cm`), lề mặc định `2.54 / 3.18 cm`, lề trái biên bản `< 3.0 cm` |
| **Văn bản của Đảng (`party` — HD 05-HD/VPTW)** | 0 / 1 pass | 1 / 1 pass | Chặn lỗi áp nhầm lề co giãn NĐ30 (`3.5 / 2.5 cm`) lên văn bản Đảng đòi hỏi lề cố định `2.0 / 2.0 / 3.0 / 1.5 cm`, `14pt`, giãn dòng `Exactly 20pt` |
| **VBQPPL (`administrative` + `legal-normative.json` — NĐ 187/2025)** | 0 / 1 pass | 1 / 1 pass | Bắt lỗi lề trên/dưới `2.5 cm` khi Phụ lục I NĐ 187/2025/NĐ-CP yêu cầu `2.0 cm` và cỡ chữ `14pt` |
| **Font & màu chữ (`font`, `color`)** | Bị bắt ở case Arial & heading xanh | 0 lỗi | Ngăn font Arial/Calibri theme và màu tiêu đề xanh mặc định Word (`#2F5496`) |
| **Toàn vẹn template (`placeholder`, `numbering`)** | 0 / 3 pass | Đạt ở mọi starter & demo | Phát hiện `{{token}}` chưa điền, `[CẦN BỔ SUNG: ...]` còn sót ở bản cuối và `numId` trỏ hỏng (`99999`) |

---

## Chi tiết 16 trường hợp (Output trực tiếp từ `evals/run_validation_benchmark.py`)

```text
Baseline pass: 0/8 | Skill-guided pass: 8/8
```

| Group | Case | Profile | Status | Issues | Codes |
| --- | --- | --- | --- | ---: | --- |
| baseline | Unguided ND30 notice (Letter, Calibri, default margins) | `administrative` | `fail` | 4 | margin, page |
| baseline | Unguided ND30 dispatch (Arial font, 2.54 cm margins) | `administrative` | `fail` | 4 | font, margin |
| baseline | Unguided minutes (blue Word heading, 2.0 cm left margin) | `minutes-administrative` | `fail` | 2 | color, margin |
| baseline | Unresolved {{token}} left in final DOCX | `administrative` | `fail` | 1 | placeholder |
| baseline | Unresolved [CẦN BỔ SUNG: ...] left in final mode | `office` | `fail` | 1 | placeholder |
| baseline | Broken Word numbering reference (numId=99999) | `academic` | `fail` | 1 | numbering |
| baseline | Unguided Party doc (ND30 3.5/2.5 cm margins instead of HD 05 fixed) | `party` | `fail` | 2 | margin |
| baseline | Unguided VBQPPL doc (2.5 cm top/bottom margins vs. ND 187/2025 2.0 cm) | `administrative+legal-normative` | `fail` | 2 | margin |
| skill | ND30 notice starter filled + style role validation | `administrative` | `pass` | 0 | — |
| skill | Academic short report starter filled + validation | `academic` | `pass` | 0 | — |
| skill | Office minutes starter filled + validation | `minutes-general` | `pass` | 0 | — |
| skill | Skill-generated 2-page ND30 notice (table + PAGE field) | `administrative` | `pass` | 0 | — |
| skill | Multi-section office demo (portrait + landscape annex) | `minutes-general` | `pass` | 0 | — |
| skill | Custom organization config (Arial 12pt, no landscape) | `custom` | `pass` | 0 | — |
| skill | Skill-guided Party doc (HD 05-HD/VPTW fixed margins, 14pt, 20pt exact) | `party` | `pass` | 0 | — |
| skill | Skill-guided VBQPPL doc (ND 187/2025 Appendix I rules, 14pt, 2.0 cm top/bottom) | `administrative+legal-normative` | `pass` | 0 | — |

---

## Tệp liên quan

- Script chạy benchmark: [`evals/run_validation_benchmark.py`](run_validation_benchmark.py)
- Kết quả JSON chuẩn máy đọc: [`evals/benchmark-results.json`](benchmark-results.json)
- Đối chiếu văn bản thật (UBND xã Hương Khê 11/2026 & Cục QLĐKKD Mẫu 27): [`evals/audit-comparison.md`](audit-comparison.md)
- Đánh giá hành vi agent (14 tình huống gồm 4 tình huống Đảng/NĐ30/VBQPPL do `@triensslci` đóng góp): [`evals/behavior-cases.json`](behavior-cases.json) · [`evals/manual-results.md`](manual-results.md)
