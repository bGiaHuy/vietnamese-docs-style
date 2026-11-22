# Vietnamese Docs Style

Skill và bộ công cụ Python (`vndocs`) giúp người dùng AI coding agent, cán bộ hành chính, nhân viên văn phòng, sinh viên và lập trình viên tạo, điền mẫu và kiểm tra file Word (`.docx`) theo đúng quy ước trình bày Việt Nam (Nghị định 30/2020/NĐ-CP, văn bản của Đảng theo QĐ 399-QĐ/TW & HD 05-HD/VPTW, văn bản quy phạm pháp luật, báo cáo học thuật, biên bản họp).

*Agent skill and local Python CLI (`vndocs`) for generating, template-filling, and validating Vietnamese `.docx` documents under Decree 30/2020/NĐ-CP administrative, Communist Party (QĐ 399-QĐ/TW & HD 05-HD/VPTW), legal-normative, academic, and office conventions.*

[![DOCX Tests](https://github.com/bGiaHuy/vietnamese-docs-style/actions/workflows/tests.yml/badge.svg)](https://github.com/bGiaHuy/vietnamese-docs-style/actions/workflows/tests.yml)
[![Static Checks](https://github.com/bGiaHuy/vietnamese-docs-style/actions/workflows/lint.yml/badge.svg)](https://github.com/bGiaHuy/vietnamese-docs-style/actions/workflows/lint.yml)
[![Benchmark](https://img.shields.io/badge/benchmark-0%2F8%20baseline%20vs%208%2F8%20skill-brightgreen)](evals/BENCHMARK.md)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Python 3.10+](https://img.shields.io/badge/python-3.10%2B-blue.svg)](pyproject.toml)
[![Platforms](https://img.shields.io/badge/platforms-Windows%20%7C%20macOS%20%7C%20Linux-informational)](#cách-bắt-đầu)
[![AI Agents](https://img.shields.io/badge/agents-Claude%20Code%20%7C%20Codex%20%7C%20Cursor%20%7C%20Gemini%20CLI%20%7C%20Copilot%20%7C%20Windsurf%20%7C%20Antigravity-6f42c1)](#tương-thích-ai-agent)

**Trạng thái:** `v0.2.0` (`53/53` automated tests pass trên Python `3.10` và `3.12`) · **Giấy phép:** [MIT](LICENSE) · **Cài đặt:** chưa đóng gói bản GitHub Release riêng; cài skill nhanh qua `npx skills add bGiaHuy/vietnamese-docs-style`, cài từ mã nguồn (`pip install -e .`), hoặc xuất ZIP qua `vndocs package`.

**Ngôn ngữ / Language:** [Tiếng Việt](#tiếng-việt) | [English](#english)

---

# Tiếng Việt

## Bằng chứng trực quan

Ảnh dưới là kết quả render trực tiếp (`vndocs render --png` bằng LibreOffice và PyMuPDF) từ **Thông báo hành chính 2 trang do skill tạo ra với dữ liệu giả định**:

| Trang 1 — Header NĐ30, căn cứ và bảng lặp tiêu đề | Trang 2 — Số trang `2` căn giữa, nơi nhận và khối ký |
| :---: | :---: |
| ![Thông báo 2 trang tạo bằng skill — Trang 1](docs/images/skill-generated-notice-p1.png) | ![Thông báo 2 trang tạo bằng skill — Trang 2](docs/images/skill-generated-notice-p2.png) |

Tệp nguồn vượt qua `validate_docx.py --profile administrative` với 0 lỗi: khổ A4 (`21.0 × 29.7 cm`), lề NĐ30 (`3.0 / 1.5 / 2.0 / 2.0 cm`), bảng đầu 2×2 giữ `Số:` và `Ngày` cùng hàng ngang, đường kẻ dưới Tiêu ngữ dài `4.95 cm` khớp bề ngang chữ, ẩn số trang ở trang 1 và hiện `2` căn giữa ở đầu trang 2. File minh họa giữ nguyên vùng ký trống, không tạo chữ ký hay con dấu giả.

<details>
<summary>So sánh thêm: file Word mặc định (Before) vs. starter sau audit (After) và đối chiếu văn bản cơ quan nhà nước</summary>

| Mặc định không hướng dẫn (`before-naive`) | Starter NĐ30 sau audit (`after-audit-notice`) |
| :---: | :---: |
| ![Before: file docx mặc định](docs/images/before-naive.png) | ![After: starter NĐ30 sau audit](docs/images/after-audit-notice.png) |
| **`validate_docx.py` → `fail` (4 lỗi):** khổ US Letter (`21.59 × 27.94 cm`), lề mặc định (`3.18 / 2.54 cm`), font Calibri theme, dùng dấu `---` làm đường kẻ, thiếu bố cục `Số:`/`Ngày` và `Nơi nhận:`. | **`validate_docx.py` → `pass` (0 lỗi):** khổ A4, lề chuẩn NĐ30, Times New Roman `#000000`, `Số:` và `Ngày` thẳng hàng trong bảng 2×2, đường kẻ Tiêu ngữ `4.95 cm`, ghi chú ký in nghiêng khóa `cantSplit`. |

Xem thêm so sánh Starter Trước/Sau Audit và kết quả chạy validator trên 2 văn bản `.docx` thật tại Việt Nam (**UBND xã Hương Khê, Hà Tĩnh 11/2026** và **Cục Quản lý đăng ký kinh doanh Mẫu 27, 2025**) ở [evals/audit-comparison.md](evals/audit-comparison.md).

</details>

## Giới thiệu ngắn

Khi LLM hoặc script Python tự tạo một file `.docx` mới mà thiếu quy ước địa phương, file thường rơi về khổ US Letter, lề `2.54 cm`, font Calibri hoặc Arial và tiêu đề màu xanh mặc định của Word. Nhiều prompt còn khiến AI áp nhầm Quốc hiệu, Tiêu ngữ của Nghị định 30/2020/NĐ-CP vào báo cáo học thuật hoặc thông báo công ty.

`vietnamese-docs-style` tách rõ hai phần: bộ hướng dẫn theo ngữ cảnh ([SKILL.md](SKILL.md), [AGENTS.md](AGENTS.md), `references/`) để agent chọn đúng thể thức ngay từ đầu, đi kèm CLI `vndocs` để dựng file, điền mẫu giữ nguyên XML gốc, kiểm tra định dạng và render ảnh từng trang.

## Tính năng chính

- Chọn đúng thể thức theo loại văn bản và nơi tiếp nhận qua 8 profile (`administrative`, `party`, `academic`, `minutes-general`, `office`, `proposal`, `minutes-administrative`, `custom`); NĐ30/2020 chỉ áp dụng cho văn bản hành chính nhà nước, tách biệt với văn bản của Đảng (QĐ 399-QĐ/TW, HD 05-HD/VPTW), văn bản quy phạm pháp luật (Phụ lục I NĐ 187/2025/NĐ-CP), bài học thuật hay tài liệu nội bộ doanh nghiệp.
- Sinh file `.docx` đa profile trực tiếp từ JSON đúng [JSON Schema](assets/schemas/document-input.schema.json) (`vndocs generate ... --validate`), điền token `{{name}}` tách nhiều run (`vndocs fill`) hoặc vá đúng vị trí văn bản bằng `vndocs edit --source-sha256`, giữ nguyên từng byte của các thành phần XML ngoài vùng sửa.
- Soi nhanh file mẫu và phát hiện lỗi encoding `Ð` (`U+00D0` -> `Đ` `U+0110`) / placeholder / `kính giửi` qua `vndocs inspect` (kèm `--normalize-output` sửa không làm hỏng XML khác), rồi kiểm tra khổ giấy, lề từng section, font/màu kế thừa, named style role và tham chiếu numbering qua `vndocs validate`.
- Render PDF và PNG từng trang bằng LibreOffice + PyMuPDF trong profile tạm cô lập (`vndocs render --png`), tránh trường hợp PDF cũ nằm sẵn trên đĩa bị nhận nhầm là bản mới.
- Kèm bộ câu hỏi phỏng vấn lấy dữ liệu ([interview-questions.md](references/interview-questions.md)), hướng dẫn xem trước Markdown 2 cột khi chat thuần text ([markdown-preview-fallback.md](references/markdown-preview-fallback.md)), chỉ mục agent ([llms.txt](llms.txt)) và đóng gói ZIP gọn sạch (`vndocs package`) không dính dữ liệu nghiên cứu hay test.

### Phạm vi hỗ trợ theo Profile

| Profile | Nhóm tài liệu | Rule & Validator | Starter template / Cấu hình tham chiếu |
| --- | --- | --- | --- |
| `administrative` | Thông báo, công văn, quyết định, tờ trình thuộc NĐ30/2020 ([29 loại & ký hiệu](references/administrative-doc-types-nd30.md), [văn phong](references/administrative-writing-style.md)) | Đầy đủ | [`administrative-notice`](assets/templates/administrative-notice.docx) (tham chiếu Mẫu 1.4 NĐ30); VBQPPL dùng kèm [`legal-normative.json`](assets/rules/legal-normative.json) ([tham chiếu](references/legal-normative-documents.md)) |
| `party` | Văn bản của Đảng theo QĐ 399-QĐ/TW & HD 05-HD/VPTW ([tham chiếu](references/party-documents.md)) | Builder & validator (lề cố định `2.0/2.0/3.0/1.5 cm`, `14pt`, giãn dòng `Exactly 20pt`, số trang `14pt` cách mép trên `1.0 cm`) | Dựng qua `build_docx` (`new_doc('party')`) hoặc mẫu cấp ủy |
| `academic` | Báo cáo học thuật ngắn, tiểu luận, học phần | Đầy đủ | [`academic-short-report`](assets/templates/academic-short-report.docx) |
| `minutes-general` | Biên bản họp nhóm, họp nội bộ doanh nghiệp | Đầy đủ | [`office-minutes`](assets/templates/office-minutes.docx) |
| `office` | Thông báo, ghi nhớ, báo cáo nội bộ văn phòng | Builder & validator | Dựng qua `build_docx` hoặc mẫu tổ chức |
| `proposal` | Đề xuất dự án, kế hoạch công việc | Builder & validator | Dựng qua `build_docx` hoặc mẫu tổ chức |
| `minutes-administrative` | Biên bản thủ tục hành chính cụ thể | Rule lề/font/số trang | Chọn mẫu đúng thủ tục thực tế |
| `custom` | Biểu mẫu `.docx` riêng của trường/cơ quan/công ty | Giữ mẫu; kiểm theo `--config` | Người dùng cung cấp ([ví dụ config](assets/rules/custom-example.json)) |

### Benchmark công cụ công khai (Baseline 0/8 vs. Skill 8/8)

Chạy qua `python evals/run_validation_benchmark.py` trên **16 case deterministic** (8 file baseline mặc định / áp nhầm lề vs. 8 đầu ra qua starter/builder/config của skill, được đối chiếu liên tục trong CI). Đây là **benchmark của bộ công cụ kiểm tra DOCX** (`validate_docx.py`), không phải bảng xếp hạng đo độ chính xác sinh văn bản tự do của từng mô hình LLM:

| Hạng mục kiểm tra (`validate_docx.py`) | Baseline mặc định | Đầu ra dùng Skill |
| --- | :---: | :---: |
| **Tổng số case đạt `pass`** | **0 / 8 (0%)** | **8 / 8 (100%)** |
| Khổ giấy & lề hành chính NĐ30 (`administrative`, `minutes-administrative`) | 0 / 3 pass (`margin`, `page`) | 0 lỗi |
| Văn bản của Đảng (`party` — HD 05-HD/VPTW vs. áp nhầm lề NĐ30) | 0 / 1 pass (`margin`) | 1 / 1 pass |
| VBQPPL (`administrative` + `legal-normative.json` — NĐ 187/2025) | 0 / 1 pass (`margin`) | 1 / 1 pass |
| Font & màu chữ (`font`, `color`) | 2/2 case gán sai bị bắt | 0 lỗi |
| Placeholder sót / `numId` hỏng (`placeholder`, `numbering`) | 3/3 case lỗi bị bắt | 0 lỗi |

Xem bảng kết quả đầy đủ 16 case, file JSON cho CI và đối chiếu văn bản thật tại: [`evals/BENCHMARK.md`](evals/BENCHMARK.md) · [`evals/benchmark-results.json`](evals/benchmark-results.json) · [`evals/audit-comparison.md`](evals/audit-comparison.md) · [`evals/manual-results.md`](evals/manual-results.md).

## Cách bắt đầu

**Yêu cầu hệ thống:** Node.js/npm (nếu cài skill bằng `npx skills`) và Python `3.10+` (`python-docx`, `PyMuPDF`, `lxml` cài tự động cho bộ công cụ `vndocs`). LibreOffice chỉ cần khi xuất PDF/PNG bằng `vndocs render`.

### Cài và cập nhật skill qua `npx skills`

```bash
# Cài skill
npx skills add bGiaHuy/vietnamese-docs-style

# Kiểm tra bản cập nhật
npx skills check

# Cập nhật các skill đã cài
npx skills update
```

### Cài bộ công cụ Python `vndocs` từ mã nguồn

```bash
git clone https://github.com/bGiaHuy/vietnamese-docs-style.git
cd vietnamese-docs-style
python -m pip install -e .
# Nếu chỉ chạy trực tiếp `python scripts/...py` và không dùng `vndocs`:
# python -m pip install -r requirements.txt
vndocs doctor
```

Người dùng cuối không cần cài `requirements-dev.txt`.

### Tương thích AI Agent

| AI Agent / IDE | Tệp cấu hình trong repo | Cách tích hợp |
| --- | --- | --- |
| **Claude Code** | [SKILL.md](SKILL.md) | Đọc `SKILL.md` trong repo hoặc cài từ ZIP đóng gói qua `vndocs package --output dist/vietnamese-docx.zip` |
| **OpenAI Codex** | [AGENTS.md](AGENTS.md), [SKILL.md](SKILL.md) | Tự đọc `AGENTS.md` ở thư mục gốc và nạp `SKILL.md` theo nhu cầu |
| **Cursor** | [AGENTS.md](AGENTS.md), [SKILL.md](SKILL.md) | Mở repo làm workspace để nhận `AGENTS.md`, hoặc gắn `SKILL.md` vào rule/context |
| **Gemini CLI** | [SKILL.md](SKILL.md), [AGENTS.md](AGENTS.md) | Cài qua `npx skills add bGiaHuy/vietnamese-docs-style` hoặc đọc `SKILL.md` / `references/` |
| **Antigravity** | [SKILL.md](SKILL.md), [AGENTS.md](AGENTS.md) | Cài qua `npx skills` hoặc nhận diện tự động `SKILL.md` / `AGENTS.md` tại thư mục gốc |
| **GitHub Copilot (VS Code)** | [AGENTS.md](AGENTS.md), [SKILL.md](SKILL.md) | Đọc hướng dẫn repo qua `AGENTS.md` / `SKILL.md` trong workspace |
| **Windsurf** | [AGENTS.md](AGENTS.md), [SKILL.md](SKILL.md) | Mở workspace chứa `AGENTS.md` và trỏ agent tới `SKILL.md` |
| **OpenCode / Cline / Roo Code** | [AGENTS.md](AGENTS.md), [SKILL.md](SKILL.md) | Đọc chuẩn `AGENTS.md` ở gốc repo và gọi lệnh `vndocs` |

## Cách sử dụng

### Chạy thử bộ demo

```bash
vndocs demo --output-dir artifacts/demo
vndocs validate artifacts/demo/office-filled.docx --profile minutes-general
vndocs render artifacts/demo/office-filled.docx --png
```

### Thao tác thường dùng với CLI và Agent

- **Sinh DOCX từ JSON đa profile + tự kiểm tra ngay:**
  ```bash
  vndocs generate assets/schemas/examples/nd30-dispatch.json artifacts/cong-van.docx --validate
  vndocs generate assets/schemas/examples/party-resolution.json artifacts/nq-chibo.docx --validate
  ```
- **Soi mẫu cơ quan và chuẩn hóa ký tự `Ð` (`U+00D0` -> `Đ`) giữ nguyên XML khác:**
  ```bash
  vndocs inspect source.docx --normalize-output artifacts/clean-template.docx
  ```
- **Điền starter có sẵn (`{{token}}`):**
  ```bash
  vndocs fill assets/templates/administrative-notice.docx artifacts/notice.docx --values assets/templates/administrative-notice.values.json
  vndocs validate artifacts/notice.docx --profile administrative --config assets/templates/administrative-notice.rules.json
  ```
- **Sửa giữ nguyên mẫu có sẵn không dùng token:**
  ```bash
  vndocs edit source.docx artifacts/edited.docx --edits slots.json --source-sha256 <SHA256>
  ```
- **Prompt mẫu gửi AI Agent:**
  - *Hành chính:* `Tạo file DOCX thông báo lịch họp theo thể thức NĐ30/2020; chỗ nào thiếu thông tin ký hoặc số văn bản thì đánh dấu [CẦN BỔ SUNG: ...].`
  - *Học thuật:* `Viết báo cáo học thuật ngắn về ứng dụng AI trong giáo dục (chưa có mẫu trường), không tự thêm số liệu nghiên cứu.`
  - *Biên bản họp:* `Chuyển ghi chú họp sau thành biên bản họp nội bộ gồm thảo luận, quyết định đã chốt và bảng phân công.`
  - *Mẫu riêng:* `Điền dữ liệu vào mẫu DOCX đính kèm và giữ nguyên style, bảng, header/footer ngoài vùng sửa.`

Danh sách đầy đủ subcommand, cấu hình JSON và exit codes (`0` pass, `1` fail, `2` partial) nằm tại [references/tooling.md](references/tooling.md).

## Giới hạn và tích hợp

- **Cài đặt LibreOffice (tùy chọn để render PNG):**
  - Windows: `winget install TheDocumentFoundation.LibreOffice`
  - macOS: `brew install --cask libreoffice`
  - Ubuntu / Debian: `sudo apt-get update && sudo apt-get install -y libreoffice-writer fonts-liberation`
  - Nếu `soffice` không có trong `PATH`, truyền `--renderer` hoặc đặt biến môi trường `VNDOCS_RENDERER`.
- **Tài khoản và quota:** CLI `vndocs` và toàn bộ script trong `scripts/` chạy offline trên máy, không cần API key hay tài khoản. Khi chạy qua AI Agent, lượng token sử dụng phụ thuộc vào gói tài khoản hoặc API của agent bạn đang dùng.
- **Phạm vi chưa hỗ trợ:**
  - `validate_docx.py` không chứng nhận hợp lệ pháp lý hay tính chính xác của dữ liệu; văn bản hành chính chính thức vẫn cần người chịu trách nhiệm kiểm tra trước khi phát hành.
  - Công cụ từ chối sửa các file `.docx` đã ký số (`_xmlsignatures/`) và không tạo chữ ký hay con dấu giả.
  - Phần `unchecked` trong báo cáo kiểm tra gồm: theme font/color phức tạp, conditional table style, text box, chart, footnote/endnote, tracked revisions và tính đúng của thứ tự đánh số/cập nhật field.

## Dữ liệu và quyền riêng tư

- Toàn bộ xử lý `.docx` qua `vndocs` và `scripts/` diễn ra cục bộ trên máy. Đầu ra mặc định lưu trong `artifacts/` và `dist/` (cả hai đã nằm trong `.gitignore`).
- Bộ công cụ Python không gửi telemetry, không đồng bộ đám mây và không mở kết nối mạng ra ngoài; khi render, LibreOffice chạy với profile tạm trong thư mục temp cục bộ.
- Nếu bạn sử dụng skill thông qua AI Agent chạy trên đám mây, nội dung prompt tuân theo chính sách dữ liệu của nhà cung cấp mô hình đó — không dán tài liệu mật hoặc thông tin nhạy cảm chưa được phép lên dịch vụ ngoài.

## Thông tin cho người đóng góp

- **Công nghệ chính:** Python 3.10+, `python-docx`, `lxml`, `PyMuPDF`, LibreOffice (render), `pytest`, `ruff`.
- **Cài môi trường dev và chạy bộ kiểm tra:**
  ```bash
  python -m pip install -e ".[dev]"
  python -m ruff check .
  python -m compileall -q scripts tests evals
  python -m pytest -q
  python evals/run_validation_benchmark.py
  ```
- **Báo lỗi và đóng góp:** Mở issue tại [GitHub Issues](https://github.com/bGiaHuy/vietnamese-docs-style/issues) hoặc gửi Pull Request theo hướng dẫn tại [CONTRIBUTING.md](CONTRIBUTING.md).

<details>
<summary>Cấu trúc thư mục chính của repository</summary>

```text
vietnamese-docs-style/
├── README.md                  # Tài liệu chính (Việt / Anh)
├── LICENSE                    # Giấy phép MIT
├── CONTRIBUTING.md            # Hướng dẫn review và đóng góp
├── pyproject.toml             # Đóng gói PEP 621 & CLI `vndocs`
├── SKILL.md                   # Hướng dẫn chính cho AI Agent
├── AGENTS.md                  # Quy tắc chung cho Codex, Cursor, Copilot, Windsurf, OpenCode
├── llms.txt                   # Chỉ mục tài liệu & thứ tự đọc cho web/sandbox agents
├── requirements.txt           # Runtime dependencies
├── requirements-dev.txt       # Dev/CI dependencies (pytest, PyYAML, ruff)
├── assets/                    # Starter templates, JSON Schema, ví dụ đầu vào & custom rules
├── docs/images/               # Ảnh preview Before/After và đầu ra tạo bằng skill
├── docs/internal/             # Hồ sơ triển khai và nghiệm thu nội bộ
├── evals/                     # Benchmark tái lập, báo cáo audit & manual evals
├── references/                # Tài liệu quy chuẩn NĐ30, văn bản Đảng, VBQPPL, học thuật, biên bản, tooling
├── scripts/                   # Package Python & CLI `vndocs`
└── tests/                     # 53 automated tests (pytest + LibreOffice render)
```

</details>

## Credits và giấy phép

- **Tác giả:** Bùi Gia Huy ([bGiaHuy](https://github.com/bGiaHuy)) · [vietnamese-docs-style](https://github.com/bGiaHuy/vietnamese-docs-style).
- **Đóng góp & nguồn cảm hứng:**
  - Cảm ơn [@triensslci](https://github.com/triensslci) đã đóng góp profile `party` cùng bộ tài liệu quy chuẩn văn bản của Đảng ([party-documents.md](references/party-documents.md) theo QĐ 399-QĐ/TW và HD 05-HD/VPTW), văn bản quy phạm pháp luật ([legal-normative-documents.md](references/legal-normative-documents.md), [legal-normative.json](assets/rules/legal-normative.json) theo Luật 64/2025, Luật 87/2025, NĐ 78/2025 và Phụ lục I NĐ 187/2025), danh mục 29 loại văn bản hành chính & thẩm quyền ký NĐ30 ([administrative-doc-types-nd30.md](references/administrative-doc-types-nd30.md)) và hướng dẫn văn phong hành chính – công vụ ([administrative-writing-style.md](references/administrative-writing-style.md)).
  - Cảm ơn [@kanazawahere/nd30](https://github.com/kanazawahere/nd30) (giấy phép MIT) và nguồn nghiệp vụ gốc [biencuong/vbhc](https://github.com/biencuong/vbhc) (Unlicense) về ý tưởng kiến trúc sinh `.docx` từ JSON Schema (`vndocs generate`), lệnh soi mẫu & phát hiện lỗi encoding `Ð` (`U+00D0`) (`vndocs inspect`), bộ câu hỏi phỏng vấn người dùng ([interview-questions.md](references/interview-questions.md)), hướng dẫn xem trước Markdown ([markdown-preview-fallback.md](references/markdown-preview-fallback.md)) và chỉ mục [llms.txt](llms.txt) — đã được thiết kế lại phù hợp kiến trúc đa profile và OOXML lossless của dự án.
- **Tham chiếu & thư viện nền tảng:** Đối chiếu thể thức từ **Nghị định 30/2020/NĐ-CP**, **Quy định 399-QĐ/TW**, **Hướng dẫn 05-HD/VPTW** và **Nghị định 187/2025/NĐ-CP**; xây dựng trên `python-docx`, `lxml`, `PyMuPDF` và `LibreOffice`.
- **Giấy phép:** Mã nguồn, script và tài liệu của dự án dùng giấy phép **MIT** ([LICENSE](LICENSE)). Các văn bản, biểu mẫu tham khảo của bên thứ ba giữ nguyên bản quyền và điều khoản gốc của đơn vị ban hành.

---

# English

## Visual Proof

Real LibreOffice renders (`vndocs render --png`) of a **2-page Vietnamese administrative notice generated with the skill and simulated data**:

| Page 1 — NĐ30 Header, Legal Bases & Repeating Table Header | Page 2 — Centered Top Page `2` & Recipients / Signer Block |
| :---: | :---: |
| ![Skill-generated 2-page notice — Page 1](docs/images/skill-generated-notice-p1.png) | ![Skill-generated 2-page notice — Page 2](docs/images/skill-generated-notice-p2.png) |

`validate_docx.py --profile administrative` passes with 0 errors: A4 (`21.0 × 29.7 cm`), NĐ30 margins (`3.0 / 1.5 / 2.0 / 2.0 cm`), 2×2 aligned `Số:` / `Ngày` header row, `4.95 cm` motto underline, page 1 number hidden, and page `2` centered in the page 2 header (no fabricated signatures or stamps).

See Before vs. After comparisons and real-world audits against human-authored Vietnamese government/enterprise `.docx` files in [evals/audit-comparison.md](evals/audit-comparison.md).

## Overview

Out of the box, LLMs and Python DOCX scripts default to US Letter pages, `2.54 cm` margins, Calibri or Arial fonts, blue Word headings, or force Decree 30/2020/NĐ-CP headers onto academic essays and private company notices.

`vietnamese-docs-style` combines context-routed instructions ([SKILL.md](SKILL.md), [AGENTS.md](AGENTS.md), `references/`) with the local `vndocs` CLI so agents pick the right conventions first and verify `.docx` structure and rendered pages before handoff.

## Key Features

- Context-routed formatting across 8 profiles (`administrative`, `party`, `academic`, `minutes-general`, `office`, `proposal`, `minutes-administrative`, `custom`) plus [`legal-normative.json`](assets/rules/legal-normative.json) for legal-normative documents (VBQPPL), keeping Decree 30/2020/NĐ-CP, Party guidelines (QĐ 399-QĐ/TW & HD 05-HD/VPTW), academic reports, and corporate minutes clearly separated.
- Declarative multi-profile JSON-to-DOCX generation (`vndocs generate --validate`, [JSON Schema](assets/schemas/document-input.schema.json)), lossless `{{token}}` template filling (`vndocs fill`, including split-run tokens), and SHA256-verified slot patching (`vndocs edit`) while keeping untouched ZIP members byte-identical.
- Template inspection & lossless `Ð` (`U+00D0` -> `Đ` `U+0110`) encoding normalization (`vndocs inspect --normalize-output`), plus structural and style validation (`vndocs validate`) with an explicit `unchecked` list.
- Real PDF and PNG rendering (`vndocs render --png`) using LibreOffice in an isolated temporary profile so stale PDFs never trigger false passes.
- Root agent sitemap ([llms.txt](llms.txt)), pre-drafting interview checklist ([interview-questions.md](references/interview-questions.md)), chat Markdown layout fallback ([markdown-preview-fallback.md](references/markdown-preview-fallback.md)), and clean skill ZIP packaging (`vndocs package`) that excludes research corpora and tests.

## Getting Started & Agent Support

Requires Node.js/npm (for `npx skills`) and Python `3.10+` (`python-docx`, `PyMuPDF`, `lxml` for the local `vndocs` CLI); LibreOffice is optional and only needed for `vndocs render`.

```bash
# Install the skill
npx skills add bGiaHuy/vietnamese-docs-style

# Check for updates
npx skills check

# Update installed skills
npx skills update
```

To install the local Python `vndocs` CLI from source:

```bash
git clone https://github.com/bGiaHuy/vietnamese-docs-style.git
cd vietnamese-docs-style
python -m pip install -e .
# Or install runtime deps directly without the `vndocs` CLI:
# python -m pip install -r requirements.txt
vndocs doctor
```

| AI Agent / IDE | Entry files | Integration |
| --- | --- | --- |
| **Claude Code** | [SKILL.md](SKILL.md) | Point to `SKILL.md` or install the ZIP built by `vndocs package` |
| **OpenAI Codex / Cursor / Copilot / Windsurf / OpenCode** | [AGENTS.md](AGENTS.md), [SKILL.md](SKILL.md) | Open the repo as workspace (`AGENTS.md` auto-discovered) and invoke `vndocs` |
| **Gemini CLI / Antigravity** | [SKILL.md](SKILL.md), [AGENTS.md](AGENTS.md) | Install via `npx skills add bGiaHuy/vietnamese-docs-style` or load `SKILL.md` / `AGENTS.md` |

## Usage

```bash
vndocs demo --output-dir artifacts/demo
vndocs validate artifacts/demo/office-filled.docx --profile minutes-general
vndocs render artifacts/demo/office-filled.docx --png
```

Common workflows also include `vndocs generate ... --validate`, `vndocs inspect --normalize-output ...`, `vndocs fill`, `vndocs edit --source-sha256 ...`, `vndocs templates`, and `vndocs package` (or calling `python scripts/<tool>.py`). See [references/tooling.md](references/tooling.md) for JSON rule keys and exit codes (`0` pass, `1` fail, `2` partial).

### Public Reproducible Benchmark (`0/8` Baseline vs. `8/8` Skill)

Run `python evals/run_validation_benchmark.py` to execute all **16 deterministic validation cases** covering Decree 30/2020/NĐ-CP administrative notices/dispatches, Communist Party layouts (`party`, HD 05-HD/VPTW), legal-normative rules (`VBQPPL` / Decree 187/2025/NĐ-CP), font/color enforcement, placeholder checks, `numId` integrity, and custom organization configs.

- **Result:** `0 / 8` naive baseline documents pass vs. `8 / 8` skill-guided outputs pass (verified against published JSON in CI).
- **Scope note:** Measures deterministic `validate_docx.py` tooling behavior rather than independent multi-LLM prompt accuracy.
- **Full reports:** [`evals/BENCHMARK.md`](evals/BENCHMARK.md) · [`evals/benchmark-results.json`](evals/benchmark-results.json) · [`evals/audit-comparison.md`](evals/audit-comparison.md) · [`evals/manual-results.md`](evals/manual-results.md).

## Limitations & Integrations

- **LibreOffice (optional for rendering):** `winget install TheDocumentFoundation.LibreOffice` (Windows), `brew install --cask libreoffice` (macOS), or `sudo apt-get install -y libreoffice-writer fonts-liberation` (Ubuntu/Debian); override via `--renderer` or `VNDOCS_RENDERER`.
- **Accounts & Quotas:** `vndocs` runs offline with zero API keys or accounts. Agent execution uses whichever model subscription or quota you configure in your agent client.
- **Unsupported / uncertified scope:** Not a legal compliance certificate; refuses editing digitally signed `_xmlsignatures/` files; never fabricates signatures or stamps; leaves theme fonts/colors, conditional table styles, text boxes, charts, footnotes, and tracked revisions in `unchecked`. Human review is required before official issuance.

## Data & Privacy

- All `vndocs` / `scripts/` execution runs locally with zero telemetry, zero cloud sync, and zero outbound network requests. Local outputs go to gitignored `artifacts/` and `dist/`, and LibreOffice renders in an isolated temp directory.
- When using cloud-hosted AI agents, prompt text is governed by that provider's privacy terms — avoid sending confidential records without authorization.

## Contributing

- **Stack:** Python 3.10+, `python-docx`, `lxml`, `PyMuPDF`, LibreOffice, `pytest`, `ruff`.
- **Checks:**
  ```bash
  python -m pip install -e ".[dev]"
  python -m ruff check .
  python -m compileall -q scripts tests evals
  python -m pytest -q
  python evals/run_validation_benchmark.py
  ```
- Report bugs or submit PRs via [GitHub Issues](https://github.com/bGiaHuy/vietnamese-docs-style/issues) following [CONTRIBUTING.md](CONTRIBUTING.md).

## Credits & License

- **Author:** Bùi Gia Huy ([bGiaHuy](https://github.com/bGiaHuy)) · [vietnamese-docs-style](https://github.com/bGiaHuy/vietnamese-docs-style).
- **Contributors & Inspiration:**
  - Special thanks to [@triensslci](https://github.com/triensslci) for contributing the `party` profile and reference guides for Communist Party documents ([party-documents.md](references/party-documents.md) per QĐ 399-QĐ/TW & HD 05-HD/VPTW), legal-normative documents ([legal-normative-documents.md](references/legal-normative-documents.md) and [legal-normative.json](assets/rules/legal-normative.json) per Laws 64/2025 & 87/2025, Decrees 78/2025 & 187/2025), the 29 Decree 30/2020/NĐ-CP document types & signing rules ([administrative-doc-types-nd30.md](references/administrative-doc-types-nd30.md)), and administrative writing style guidelines ([administrative-writing-style.md](references/administrative-writing-style.md)).
  - Appreciation to [@kanazawahere/nd30](https://github.com/kanazawahere/nd30) (MIT License) and upstream [biencuong/vbhc](https://github.com/biencuong/vbhc) (Unlicense) for inspiring the declarative JSON-to-DOCX generator workflow (`vndocs generate`), template inspection & `U+00D0` (`Ð -> Đ`) encoding detection (`vndocs inspect`), pre-drafting interview questions ([interview-questions.md](references/interview-questions.md)), Markdown layout fallback ([markdown-preview-fallback.md](references/markdown-preview-fallback.md)), and root [llms.txt](llms.txt) index, re-engineered natively for multi-profile OOXML workflows.
- **Credits:** Formatting conventions referenced against Decree 30/2020/NĐ-CP, Regulation 399-QĐ/TW, Guideline 05-HD/VPTW, and Decree 187/2025/NĐ-CP; built on `python-docx`, `lxml`, `PyMuPDF`, and `LibreOffice`.
- **License:** Project-authored source, scripts, and documentation are licensed under the **MIT License** ([LICENSE](LICENSE)). Third-party reference materials remain under their original owners' licenses.
