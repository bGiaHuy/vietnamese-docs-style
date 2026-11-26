# Giao thức Đánh giá Độc lập (`EVALUATION-PROTOCOL.md` — v1.0 Pre-Registered)

> [!IMPORTANT]
> **Tài liệu đăng ký trước (Pre-Registration):** Giao thức này chốt toàn bộ phương pháp luận, phân tầng dữ liệu, thiết kế đối chứng 3 nhánh, chuẩn chấm độc lập (`eval_harness/`) và quy tắc thống kê **trước khi thu thập dữ liệu hoặc tạo test case** cho tập `val-v1` và `test-sealed-v1`. Baseline phát triển cũ (`dev-v2.0`) đã được đóng băng nguyên trạng tại [snapshots/dev-v2.0-manifest.json](snapshots/dev-v2.0-manifest.json).

---

## 1. Mục tiêu & Mô hình Rủi ro (Threat Model)

Trong quá trình phát triển `vndocs`, benchmark lịch sử [LAYOUT-BENCHMARK.md](LAYOUT-BENCHMARK.md) (`dev-v2.0`, 30 ca + 50 Live A/B trajectories) đóng vai trò **kiểm thử hồi quy kỹ thuật** nhưng tồn tại 4 rủi ro phương pháp luận và 1 hiện tượng bão hòa trần cần xử lý triệt để trước khi đưa ra bất kỳ tuyên bố khoa học nào về khả năng tổng quát hóa:

| Rủi ro / Hạn chế | Biểu hiện ở `dev-v2.0` | Biện pháp xử lý trong Giao thức v1.0 |
|---|---|---|
| **1. Development Overfitting** | 27/30 file `.docx` được sinh từ `scripts/build_docx.py` & `generate_docx.py`; 3/30 từ `assets/templates/`; 0/30 tài liệu thực tế; prompt nhắc sát tham số nội bộ (`clear_cell_overrides`, `sz=6`, `styles.xml`, `mc:Ignorable`). | Cấm sinh `.docx` bằng builder của dự án cho `val-v1` và `test-sealed-v1`; chỉ dùng tài liệu `.docx` thực tế (*in-the-wild*, đã ẩn danh hóa) kèm nhật ký nguồn gốc và kiểm định prompt tự nhiên 0-leakage. |
| **2. Oracle & Grader Coupling** | `evals/run_layout_benchmark.py` vừa chạy golden plan hardcode (`_execute_oracle_case`), vừa `import` trực tiếp `scripts/` (`map_docx`, `diff_layout`, `ooxml_order`, `auto_paginate`, `validate_docx`) để tự chấm kết quả của mình. | Xây dựng `eval_harness/` tách biệt hoàn toàn tại gốc repo, **0 import** từ `scripts/` hoặc `evals/` (kiểm tra tự động bằng AST CI guard); dùng OPC + ECMA-376 validator độc lập và parser/renderer trực tiếp từ LibreOffice/PyMuPDF. |
| **3. Distribution Bias** | Tài liệu `long` mới ở mức ~38.7 KB `document.xml` (đoạn văn lặp đều); thiếu dị bản OOXML thực tế (`tblW` dạng `pct`/`auto`, kế thừa nhiều tầng qua `tblStyle` / `basedOn`, bảng trong `w:sdt` hoặc text box, file chuyển qua lại giữa Word/Google Docs/LibreOffice/WPS, tài liệu >25 trang). | Phân tầng bắt buộc theo 5 nhóm taxonomy (A–E) và 4 bậc độ dài (bao gồm `very_long` > 25 trang hoặc > 150 KB `document.xml`), bảo đảm tách nguồn cấp cơ quan/mẫu giữa `val-v1` và `test-sealed-v1`. |
| **4. Sequential Overfitting** | `SKILL.md` thay đổi qua các đợt đo (`7,013 B` → `8,827 B` → `9,484 B`) và được tinh chỉnh sau khi quan sát token trên `T13`/`T17`. | Tách cứng 3 tập (`dev-v2.0` / `val-v1` / `test-sealed-v1`), giới hạn `val-v1` tối đa **3 milestone runs** ghi sổ tại [logs/val-run-registry.jsonl](logs/val-run-registry.jsonl), và khóa `test-sealed-v1` chạy **đúng 1 lần duy nhất** trên git tag đã đóng băng. |
| **5. Accuracy Ceiling Effect** | Trên `T01`–`T18`, cả `without_skill` và `with_skill` đều đạt `25/25 (100%)` hit rate ($g = 0.0\%$), nên chưa chứng minh được lợi thế độ chính xác hay an toàn biên. | Bổ sung các ca phức tạp ngoài đời, ca ràng buộc NĐ30/Đảng/VBQPPL ẩn giấu trong ngữ cảnh thật, ca biên an toàn (cả `refuse` lẫn *near-boundary valid* `edit`), và ca ngoài phạm vi (Nhóm E) để tránh bão hòa trần. |

---

## 2. Kiến trúc 3 Tập Dữ liệu (`dev-v2.0` / `val-v1` / `test-sealed-v1`)

| Tập dữ liệu | Vai trò khoa học | Quy mô | Nguồn tài liệu | Quyền tiếp cận & Ngân sách chạy |
|---|---|---:|---|---|
| **`dev-v2.0` (Frozen)** | Hồi quy kỹ thuật & baseline lịch sử | **30 cases** | 27 generated + 3 starter templates (đóng băng nguyên trạng) | Chạy tự do trong phát triển và CI; **cấm** dùng làm bằng chứng tổng quát hóa ngoài phân phối. |
| **`val-v1` (Validation)** | Phân tích lỗi, ablation, hiệu chỉnh hệ thống có kiểm soát | **20 cases** | Tài liệu `.docx` thực tế (*in-the-wild*), người thật kiểm duyệt | Công khai sau khi lập; giới hạn tối đa **3 milestone runs** đăng ký đầy đủ (cả run kém) trong [logs/val-run-registry.jsonl](logs/val-run-registry.jsonl). |
| **`test-sealed-v1` (Sealed Test)** | Đánh giá độc lập cuối cùng cho kết quả chính (headline numbers) | **30 cases** | Tài liệu `.docx` thực tế độc lập, **khác nguồn/template** với `val-v1` | Lưu **ngoài workspace** của agent phát triển; chỉ công bố SHA-256 cam kết tại [SEALED-TEST-COMMITMENT.json](SEALED-TEST-COMMITMENT.json); chạy **đúng 1 lần** trên bản freeze cuối. |

---

## 3. Câu hỏi Nghiên cứu (Research Questions)

1. **RQ1 — End-to-End Accuracy & Collateral Preservation:** Trên tài liệu `.docx` tiếng Việt thực tế với yêu cầu tự nhiên, việc kết hợp Skill + công cụ có cấu trúc (`Arm 2`) có cải thiện tỷ lệ hoàn thành đúng (*Strict Hit Rate*, *Checkpoint Score*) và giảm hư hại ngoài mục tiêu (*Non-Target Collateral Rate*) so với agent thuần không có skill (`Arm 0`) hay không?
2. **RQ2 — Knowledge vs. Tooling Ablation:** Phần cải thiện (nếu có) đến từ kiến thức thể thức/quy trình an toàn tiếng Việt (`Arm 1`: Standards-Only) hay từ giao diện thao tác có cấu trúc `vndocs` (`Arm 2`: Full Skill + `vndocs` CLI)?
3. **RQ3 — Cost & Length Scaling:** Tác động lên số token tiêu thụ (`fresh_input + output` và `total`), số lượt gọi công cụ và thời gian thực thi thay đổi ra sao theo 4 bậc độ dài (`short`, `medium`, `long`, `very_long`)?
4. **RQ4 — Safety Boundary Calibration & Negative Transfer:** Hệ thống có giữ được cân bằng giữa từ chối đúng ca bất khả thi/mơ hồ (*Refusal Accuracy* / *Unsafe-Edit Rate*) và **không từ chối oan** các ca gần ranh giới nhưng hợp lệ (*Over-Refusal Rate*)? Đồng thời, ở các tác vụ ngoài phạm vi API trực tiếp của `vndocs` (Nhóm E), skill có gây *negative transfer* làm giảm chất lượng hoặc tăng chi phí bất thường so với `Arm 0` hay không?

---

## 4. Taxonomy 5 Nhóm & Bảng Phân tầng (Quota Matrix)

Mỗi case trong `val-v1` và `test-sealed-v1` tuân theo schema [schemas/independent-case.schema.json](schemas/independent-case.schema.json) và được gắn đúng 1 nhóm taxonomy chính:

- **Nhóm A — Real-World Complex Tables:** Bảng sử dụng `tblW` theo phần trăm (`pct`) hoặc `auto`, kế thừa định dạng từ `tblStyle`, ô gộp đa tầng (`gridSpan` + `vMerge`), bảng dài qua nhiều trang cần lặp dòng tiêu đề/chống tách dòng, hoặc bảng trong section khổ ngang.
- **Nhóm B — Vietnamese Regulatory & Institutional Compliance:** Văn bản hành chính NĐ30/2020/NĐ-CP (bảo vệ Quốc hiệu, Nơi nhận, Chữ ký, cỡ chữ bảng ≥ 12pt, lề trang), văn bản Đảng HD 05-HD/VPTW 2026 (lề cố định 2.0/2.0/3.0/1.5 cm, 14pt, giãn dòng Exact 18–22pt, khoảng cách sau đoạn ≥ 6pt), VBQPPL NĐ 187/2025/NĐ-CP, và mẫu biểu mẫu cơ quan/trường đại học (tuyệt đối không ghi đè phong cách của bên nhận).
- **Nhóm C — Long-Document & Multi-Element Coherence:** Tài liệu dài 15–50+ trang (hoặc > 150 KB `word/document.xml`) có header/footer theo section, footnote/endnote, mục lục/field (`TOC`, `PAGE`), chứa phối hợp 3–5 ràng buộc chỉnh sửa đồng thời (bảng + ảnh + chú thích + ngắt trang/section) mà vẫn giữ nguyên toàn bộ phần ngoài mục tiêu.
- **Nhóm D — Grounding, Ambiguity & Safety Boundaries:** Tình huống kiểm tra an toàn biên gồm cả 2 phía:
  - *Expected `refuse`:* ảnh chụp cắt sát màn hình thật không phân biệt được giữa ≥ 2 bảng giống hệt nhau, file đã ký số OOXML (`_xmlsignatures/`), hoặc yêu cầu làm biến dạng/trái quy chuẩn không thể đáp ứng.
  - *Near-boundary valid `edit`:* ca gần ranh giới nhưng **hợp lệ** (ví dụ ảnh chụp cắt một phần nhưng có chi tiết phân biệt được, bảng hành chính NĐ30 nằm cạnh bảng bố cục, cột 0.6 cm hẹp nhưng ≥ ngưỡng tối thiểu) nhằm đo trực tiếp **Over-Refusal Rate**.
- **Nhóm E — Out-of-Scope / Negative Transfer Controls:** Các yêu cầu chỉnh sửa `.docx` hoàn toàn hợp lệ trên thực tế nhưng không có lệnh một bước tương ứng trong `vndocs` (ví dụ: sửa lại định dạng multilevel numbering bị lệch ở một chương, chỉnh nội dung/bố cục bảng nằm trong `w:txbxContent` hoặc `w:sdt`, đồng bộ footer riêng cho section 2) — dùng để phát hiện hiện tượng agent phụ thuộc quá mức vào skill dẫn đến từ chối sai hoặc vòng lặp tốn kém hơn `Arm 0`.

### Ma trận Phân bổ Quota Bắt buộc

| Chiều phân tầng | `val-v1` (20 cases) | `test-sealed-v1` (30 cases) |
|---|---:|---:|
| **Nhóm A (Complex Tables)** | 4 | 6 |
| **Nhóm B (VN Regulatory/Templates)** | 4 | 6 |
| **Nhóm C (Long & Multi-Element)** | 4 | 6 |
| **Nhóm D (Safety Boundaries: Refuse + Near-Boundary Edit)** | 4 (2–3 `refuse` + 1–2 `near_boundary_valid`) | 6 (3–4 `refuse` + 2–3 `near_boundary_valid`) |
| **Nhóm E (Out-of-Scope / Negative Transfer)** | 4 | 6 |
| **Độ dài: `short` (1–3 trang)** | 5 | 8 |
| **Độ dài: `medium` (4–10 trang)** | 6 | 9 |
| **Độ dài: `long` (11–25 trang)** | 5 | 7 |
| **Độ dài: `very_long` (> 25 trang hoặc > 150 KB XML)** | 4 | 6 |
| **Loại lệnh: `descriptive` vs `inferential`** | 10 vs 10 (≥ 40% mỗi loại) | 15 vs 15 (≥ 40% mỗi loại) |
| **Kết quả kỳ vọng: `edit` vs `refuse`** | ~16 `edit` (gồm ≥ 2 `near_boundary_valid`) vs ~4 `refuse` | ~24 `edit` (gồm ≥ 3 `near_boundary_valid`) vs ~6 `refuse` |

---

## 5. Quy trình Thu thập Dữ liệu Thực tế & Chống Rò rỉ

1. **Cấm sinh tài liệu bằng công cụ nội bộ:** Tuyệt đối **không** dùng `scripts/build_docx.py`, `scripts/generate_docx.py`, `assets/templates/`, hoặc bất kỳ agent nào từng đọc repo `vndocs` để tạo file `.docx` cho `val-v1` và `test-sealed-v1`.
2. **Nguồn gốc hợp lệ (*In-the-Wild Sourcing*):** Thu thập từ văn bản công khai (Cổng thông tin Chính phủ, Bộ/ngành, UBND, trường đại học công lập), mẫu tài liệu mã nguồn mở/công cộng, và tài liệu thực tế do tác giả/người dùng soạn trực tiếp bằng Microsoft Word, Google Docs, LibreOffice Writer hoặc WPS Office.
3. **Ẩn danh hóa & Nhật ký Provenance:** Mọi dữ liệu cá nhân nhạy cảm (nếu có trên tài liệu không công khai) phải được xóa/thay thế thủ công trong Word/LibreOffice rồi ghi đầy đủ vào `provenance` (loại nguồn, phần mềm xuất file, giấy phép, ghi chú ẩn danh).
4. **Phân tách nguồn cấp mẫu (*Source-Level Disjoint Split*):** Hai tập `val-v1` và `test-sealed-v1` **không được dùng chung file từ cùng một cơ quan ban hành hoặc cùng một bộ mẫu template gốc**, tránh hiện tượng gần trùng lặp (*near-duplicate leakage*).
5. **Ảnh chụp màn hình thực tế:** Với các ca có ảnh đính kèm (`attachments` loại screenshot), ảnh phải được cắt trực tiếp từ ảnh render toàn trang hoặc ảnh chụp cửa sổ trình xem tài liệu, **tuyệt đối không** dùng `map_docx` hay API nội bộ của `vndocs` để định vị vùng cắt.
6. **Quy tắc viết Prompt (`prompt_vi` — Zero Leakage):**
   - Viết 100% bằng ngôn ngữ tự nhiên của người dùng văn phòng/cán bộ Việt Nam.
   - **Cấm tuyệt đối** mọi từ khóa kỹ thuật nội bộ: `vndocs`, tên subcommand (`layout-edit`, `map`, `diff-layout`, `auto-paginate`), tên XML tag (`w:tblPr`, `w:gridSpan`, `mc:Ignorable`, `styles.xml`), mã đơn vị thô (`dxa`, `sz=6`, `emu`), hoặc tên anchor (`T#...`, `P#...`, `IMG#...`).
   - Bắt buộc qua rà soát độc lập trước khi khóa vào tập dữ liệu.
7. **Cổng kiểm tra khả năng giải độc lập (*Independent Solvability Gate*):** Trước khi chấp nhận một ca vào `val-v1` hoặc `test-sealed-v1`, người thiết kế phải kiểm chứng tối thiểu 2 điểm bằng `eval_harness`:
   - Có ít nhất **1 cách giải hợp lệ không dùng `vndocs`** (ví dụ sửa tay qua LibreOffice hoặc viết script `python-docx`/`lxml` độc lập) đạt `hit = 1` trên `eval_harness`.
   - Có ít nhất **1 bản sửa lỗi phổ biến** (ví dụ tràn lề phải, lệch tỷ lệ ảnh, quên đồng bộ `tcW`, sai lề Đảng, hoặc chỉnh nhầm bảng Quốc hiệu) bị `eval_harness` chấm `hit = 0`.

---

## 6. Giao thức Chấm Độc lập (`eval_harness/`) & Đánh giá Trực quan Mù

1. **Nguyên tắc cô lập mã nguồn (`Zero Scripts Imports`):** Gói `eval_harness/` đặt riêng tại gốc repo và **không được `import`** bất kỳ module nào trong `scripts/` hoặc `evals/`. Ràng buộc này được kiểm chứng tự động trong CI bằng AST parser (`tests/test_eval_harness.py`).
2. **Công bố trung thực Backend kiểm tra (`Validator & Renderer Transparency`):**
   - Mọi báo cáo kết quả phải ghi tường minh các backend thực tế đã chạy trong trường `validator_backends` và `renderer_backend` — không bao giờ suy diễn hoặc khai khống các bước chưa chạy (theo `AGENTS.md`).
   - Kiểm tra cấu trúc gói ở mức chuẩn: (a) `standalone_opc` (toàn vẹn ZIP, không có entry thư mục, đối soát `[Content_Types].xml`, kiểm tra toàn bộ target của `.rels`, parse XML mọi part); (b) `standalone_ecma376_subset` (kiểm tra thứ tự phần tử con ECMA-376 Part 1 §17 trên `document.xml`, `styles.xml`, `header*.xml`, `footer*.xml`, thuộc tính/enum/đơn vị, và `mc:AlternateContent`); (c) tùy chọn gắn `ecma376_xsd` hoặc `openxml_sdk` bên ngoài khi có sẵn; (d) xuất PDF độc lập bằng `libreoffice` (`soffice --headless`) với profile cô lập, hoặc đối chiếu chéo tùy chọn qua `word_com` trên máy có cài Microsoft Word.
   - Khi file gốc *in-the-wild* đã có sẵn dị bản hoặc cảnh báo lịch sử từ phần mềm tạo file, validator phải phân biệt rõ lỗi sẵn có (`baseline_errors`) và lỗi mới phát sinh sau chỉnh sửa (`new_errors_introduced`), chỉ đánh rớt ca khi file nộp phát sinh lỗi mới hoặc hỏng cấu trúc mục tiêu.
3. **Chấm không thiên vị biểu diễn (*Representation-Agnostic Fairness*):**
   - Chấp nhận mọi biểu diễn OOXML hợp lệ cho cùng kết quả hình học/định dạng: ví dụ độ rộng bảng/cột có thể biểu diễn bằng `dxa` hoặc `pct` quy đổi theo vùng in thực tế của section; định dạng đoạn/chữ được tính theo **giá trị hiệu lực (*effective formatting*)** qua chuỗi kế thừa `docDefaults -> style (basedOn) -> pPr/rPr`.
   - So sánh bảo toàn ngoài mục tiêu (*Non-Target Preservation*) dựa trên nội dung NFC, thuộc tính/bố cục hiệu lực của các phần tử ngoài vùng yêu cầu, bảo toàn `footnote`/`endnote`/`header`/`footer`/namespace/signature — chỉ ghi `xml_ntpr` làm telemetry phụ chứ không dùng hash XML cứng để phạt `Arm 0`/`Arm 1` nếu lưu file sạch qua `python-docx`.
4. **Đánh giá Trực quan Mù (*Blind Human Visual Spot-Check Rubric*):**
   - Áp dụng cho tối thiểu mẫu phân tầng (≥ 10 ca/tập) hoặc toàn bộ các ca có thay đổi bố cục nhìn thấy được.
   - Trình chấm chỉ hiển thị ảnh PNG trang trước/sau với mã ẩn ngẫu nhiên (loại bỏ hoàn toàn tên Arm, model, nhật ký lệnh).
   - Tiêu chí chấm thang 1–5 gồm 3 trục: (i) **Đạt đúng yêu cầu người dùng**, (ii) **Thẩm mỹ & căn chỉnh thị giác** (không tràn lề, không ngắt dòng/chữ ký vô lý, cân đối trang), (iii) **Không biến dạng vùng ngoài mục tiêu**. Nếu có ≥ 2 người chấm độc lập, báo cáo hệ số đồng thuận (*Cohen's $\kappa$* hoặc *Krippendorff's $\alpha$*).

---

## 7. Thiết kế Thực nghiệm 3 Arms, Kiểm soát Contamination & Kế hoạch Thống kê

### 7.1 Cấu hình 3 Nhánh (Arms)
| Nhánh | Môi trường & Công cụ được cấp | Mục đích |
|---|---|---|
| **`Arm 0`: No-Skill Baseline** | Workspace cô lập chỉ có file đề bài (`src.docx` + ảnh nếu có), Python (`python-docx`, `lxml`, `pymupdf`) và LibreOffice. **0 file của repo `vndocs`**. | Đo năng lực nền của agent tự thân. |
| **`Arm 1`: Standards-Only Ablation** | Như `Arm 0`, cộng thêm bản hướng dẫn quy chuẩn/an toàn tiếng Việt (NĐ30/Đảng/VBQPPL/checklist fail-closed) dạng văn bản thuần, **không** cấp `SKILL.md` đầy đủ và **không** có `vndocs` CLI/`scripts/`. | Tách biệt đóng góp của *kiến thức miền* khỏi *công cụ có cấu trúc*. |
| **`Arm 2`: Full Skill + `vndocs`** | Môi trường đầy đủ với `SKILL.md`, `references/`, và lệnh `vndocs` từ đúng **1 git commit SHA đóng băng**. | Đo toàn bộ hệ thống đề xuất. |

### 7.2 Kiểm soát Nhiễm chéo & Tái lập (Contamination & Determinism Controls)
- Mỗi lượt chạy diễn ra trong thư mục làm việc tạm cô lập hoàn toàn; sau khi chạy xong, bộ phân tích tự động quét toàn bộ transcript của `Arm 0` (và `Arm 1`) để xác nhận **0 lượt rò rỉ** từ khóa `vndocs`/`scripts` (`arm_a_contamination == 0`).
- Trong suốt một đợt đánh giá trên `val-v1` hoặc `test-sealed-v1`, phiên bản `SKILL.md`, `references/` và `scripts/` phải giữ cố định tại đúng 1 commit SHA (không sửa skill giữa các batch như đã xảy ra ở `dev-v2.0`).
- Ghi nhận đầy đủ exact `model_id`, ngày giờ chạy UTC/local, `prompt_tokens`, `cache_read_tokens`, `output_tokens`, `tool_calls`, và `elapsed_seconds`.

### 7.3 Số lần lặp & Kế hoạch Kiểm định Thống kê (`eval_harness/stats.py`)
- **Số lần lặp:** Mỗi cặp `(case, arm)` trên bộ đánh giá chính chạy **$k = 3$ lượt độc lập** (hoặc báo cáo rõ nếu chạy pilot $k = 1$ trên `val-v1`).
- **Thước đo báo cáo:**
  - `pass@1` (tỷ lệ đạt trung bình trên từng trial) và `pass^3` (tỷ lệ case đạt cả 3/3 trial — đo tính ổn định).
  - `Checkpoint Score` trung bình (thang $0.0$–$1.0$).
  - `Non-Target Collateral Rate` (tỷ lệ ca edit làm hỏng phần ngoài mục tiêu).
  - `Over-Refusal Rate` (tỷ lệ ca `edit` hợp lệ bị từ chối nhầm) và `Unsafe-Edit Rate` (tỷ lệ ca `refuse` bị sửa lén).
  - `Normalized Gain` $g$ — bắt buộc gắn cờ `ceiling_saturated = true` và `g = null` (hoặc `0.0` kèm giải thích) nếu baseline đạt $100\%$.
  - Trung bình & độ lệch chuẩn `fresh_input + output tokens`, `total_tokens`, `tool_calls`, `elapsed_seconds`.
- **Kiểm định thống kê (nhất quán không đổi sau pre-registration):**
  - Khoảng tin cậy **Bootstrap 95% CI** (10.000 lần lấy mẫu lại có hoàn lại trên các ca, `seed = 42`).
  - So sánh tỷ lệ Hit nhị phân ghép cặp giữa hai Arm: **McNemar exact test (2 phía)** trên các cặp bất đồng (`b`, `c`).
  - So sánh biến liên tục/thứ bậc ghép cặp (`checkpoint_score`, `fresh_plus_output_tokens`, `total_tokens`, `tool_calls`, `elapsed_seconds`): **Wilcoxon signed-rank test (2 phía)** kèm *median delta* và *rank-biserial effect size*.

---

## 8. Quản trị Công khai vs. Niêm phong & Quy tắc Dừng

| Thành phần | Trạng thái công bố | Quy tắc quản trị |
|---|---|---|
| `EVALUATION-PROTOCOL.md`, `independent-case.schema.json`, `eval_harness/`, `tests/test_eval_harness.py` | **Công khai ngay** (Pre-registered) | Mọi sửa đổi sau khi mở `test-sealed-v1` chỉ được phép sửa lỗi kỹ thuật của grader (nếu có) và phải ghi changelog công khai. |
| `dev-v2.0` + `dev-v2.0-manifest.json` | **Công khai & Đóng băng** | Không sửa số liệu cũ; giữ làm baseline lịch sử. |
| `val-v1` + `logs/val-run-registry.jsonl` | **Công khai theo từng milestone** | Mọi lần chạy đều append vào registry (tối đa **3 milestone runs**); không xóa hoặc chọn lọc run tốt nhất. |
| `test-sealed-v1` (corpus `.docx`, `prompt_vi`, đáp án) | **Niêm phong ngoài workspace** | Trước khi mở chạy, ghi SHA-256 trọn gói vào [SEALED-TEST-COMMITMENT.json](SEALED-TEST-COMMITMENT.json) và tạo tag đóng băng hệ thống (`paper-eval-freeze-vX.Y`). |

**Quy tắc dừng (*Stopping Rule*):** Khi đã mở khóa `test-sealed-v1`, thực hiện **duy nhất 1 lượt đánh giá chính thức** trên tag đã đóng băng. Mọi kết quả (kể cả các ca thất bại, lỗi timeout hoặc ca hệ thống yếu) đều phải được công bố nguyên trạng kèm phân loại lỗi đầy đủ. Nghiêm cấm sửa `SKILL.md` hay `scripts/` rồi chạy lại `test-sealed-v1` để làm đẹp số liệu; bất kỳ bản sửa lỗi nào phát sinh sau đó chỉ được phép ghi nhận trong phần phân tích hậu kiểm (*post-hoc error analysis*) và phát hành cho chu kỳ nghiên cứu tiếp theo với một tập niêm phong hoàn toàn mới.
