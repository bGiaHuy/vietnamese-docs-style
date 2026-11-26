# Báo cáo Đánh giá Độc lập `val-v1` (`VAL-V1-BENCHMARK.md`)

Tài liệu này báo cáo toàn bộ kết quả thực nghiệm độc lập **3 nhánh đối chứng** (`Arm 0: No Skill`, `Arm 1: Standards-Only`, `Arm 2: Full SKILL.md + vndocs CLI`) trên tập **`val-v1` (20 tài liệu `.docx` thực tế *in-the-wild*)**, thực hiện theo giao thức đăng ký trước [`EVALUATION-PROTOCOL.md`](EVALUATION-PROTOCOL.md) bằng bộ chấm tách biệt [`eval_harness/`](../eval_harness/).

> [!IMPORTANT]
> **Tính tách biệt & minh bạch dữ liệu:**
> - Toàn bộ baseline phát triển nội bộ cũ (`dev-v2.0`, 30 ca) vẫn được giữ nguyên trạng tại [`LAYOUT-BENCHMARK.md`](LAYOUT-BENCHMARK.md) và [`snapshots/dev-v2.0-manifest.json`](snapshots/dev-v2.0-manifest.json).
> - Không có tệp `.docx` nào trong `val-v1` được sinh bởi `scripts/build_docx.py`, `scripts/generate_docx.py`, hay `assets/templates/`.
> - Toàn bộ **420 lượt chạy được chấm điểm chính thức** (`300` ở Milestone #1 + `120` ở Milestone #2) đạt **`contamination_count = 0`** ở mọi nhánh đối chứng `Arm 0` và `Arm 1`.
> - Đã sử dụng **2 / 3 milestone slot** được phép trên `val-v1` tại [`logs/val-run-registry.jsonl`](logs/val-run-registry.jsonl) (còn `1` slot dự phòng); tập khóa [`SEALED-TEST-COMMITMENT.json`](SEALED-TEST-COMMITMENT.json) vẫn giữ trạng thái `pending_phase3_curation` bên ngoài workspace theo quy tắc Air-Gap.

---

## 1. Thiết kế Đối chứng 3 Nhánh & Kiểm định Cô lập

| Nhánh | Thành phần cung cấp cho Agent | Mục đích khoa học |
|---|---|---|
| **`arm0_no_skill`** | Chỉ prompt tiếng Việt tự nhiên + tệp `src.docx` (và ảnh cắt nếu có); workspace temp hoàn toàn cô lập, chặn skill hệ thống. | Đo năng lực nền của mô hình khi thao tác OOXML/Python thuần (`RQ1`). |
| **`arm1_standards_only`** | Như `Arm 0` + hướng dẫn thể thức/OOXML độc lập [`prompts/arm1_standards_only.md`](prompts/arm1_standards_only.md) (đã kiểm định 0 từ khóa nội bộ `vndocs`). | Tách riêng đóng góp của **tri thức thể thức Việt Nam & quy tắc an toàn** khỏi công cụ `vndocs` (`RQ2`). |
| **`arm2_full_skill`** | Toàn bộ `SKILL.md` + `references/` + `scripts/` (`vndocs` CLI) được sao chép cục bộ vào `./skill/`. | Đo tác động toàn phần của Skill + giao diện thao tác cấu trúc (`RQ1–RQ4`). |

---

## 2. Milestone #1 (`val-v1.0`, Baseline Đóng băng, `SKILL.md v0.4.0` — 300 lượt)

- **Registry ID:** `val_v1_milestone_1_20261126` (`milestone_index = 1`)
- **Commit SHA:** `1696eee3dbe8a5671bc6b0f9e1ab1cc75074561b` | **`SKILL.md` SHA-256:** `190abe4fa1f241f1da23f14896f03dd7c67bc927a7657c9bd4559edf4004fa9a`
- **Nguồn báo cáo gốc:** [`research/eval/val_v1/milestone_1/MILESTONE-1-REPORT.md`](../research/eval/val_v1/milestone_1/MILESTONE-1-REPORT.md)

### 2.1 Ma trận Chính (`5` cấu hình Model × Reasoning Effort × `3` Nhánh)

| Cấu hình | Model ID | Nhánh | Runs | `pass@1` | Hit 95% CI | Checkpoint Mean | Collateral | Over-Refusal | Unsafe-Edit | Total Tok | Calls | Time (s) |
|---|---|---|---:|---:|---|---:|---:|---:|---:|---:|---:|---:|
| `ag_flash` | `gemini-4-hydrogen-tiered-high` | `arm0_no_skill` | 20 | 65.0% | [45.0%, 85.0%] | 89.4% | 18.8% | 6.2% | 0.0% | 430,683 | 13.9 | 133.2 |
| `ag_flash` | `gemini-4-hydrogen-tiered-high` | `arm1_standards_only` | 20 | 80.0% | [60.0%, 95.0%] | 91.8% | 12.5% | 6.2% | 0.0% | 523,682 | 18.1 | 129.9 |
| `ag_flash` | `gemini-4-hydrogen-tiered-high` | `arm2_full_skill` | 20 | **85.0%** | [70.0%, 100.0%] | **92.5%** | 12.5% | **0.0%** | 0.0% | 1,158,247 | 31.7 | 174.5 |
| `ag_pro` | `gemini-4-hydrogen-tiered-xhigh` | `arm0_no_skill` | 20 | 70.0% | [50.0%, 90.0%] | 86.1% | 18.8% | 12.5% | 0.0% | 859,326 | 20.7 | 195.3 |
| `ag_pro` | `gemini-4-hydrogen-tiered-xhigh` | `arm1_standards_only` | 20 | 80.0% | [60.0%, 95.0%] | 91.5% | 18.8% | 0.0% | 0.0% | 894,662 | 22.6 | 203.1 |
| `ag_pro` | `gemini-4-hydrogen-tiered-xhigh` | `arm2_full_skill` | 20 | **80.0%** | [60.0%, 95.0%] | **91.5%** | 18.8% | **0.0%** | 0.0% | 2,496,574 | 50.0 | 259.1 |
| `codex_gpt61_low` | `gpt-6.1-sol` (`low`) | `arm0_no_skill` | 20 | 50.0% | [30.0%, 70.0%] | 81.8% | 18.8% | 6.2% | 0.0% | 139,726 | 4.3 | 277.4 |
| `codex_gpt61_low` | `gpt-6.1-sol` (`low`) | `arm1_standards_only` | 20 | **80.0%** | [60.0%, 95.0%] | 87.5% | 18.8% | 12.5% | 0.0% | 294,414 | 8.8 | 156.8 |
| `codex_gpt61_low` | `gpt-6.1-sol` (`low`) | `arm2_full_skill` | 20 | 75.0% | [55.0%, 90.0%] | **90.0%** | **12.5%** | 6.2% | 0.0% | 325,936 | 10.4 | 142.6 |
| `codex_gpt61_medium` | `gpt-6.1-sol` (`medium`) | `arm0_no_skill` | 20 | 45.0% | [25.0%, 65.0%] | 83.9% | 25.0% | 6.2% | 0.0% | 210,608 | 7.0 | 103.7 |
| `codex_gpt61_medium` | `gpt-6.1-sol` (`medium`) | `arm1_standards_only` | 20 | 75.0% | [55.0%, 90.0%] | 89.0% | 25.0% | 6.2% | 0.0% | 429,919 | 11.4 | 128.0 |
| `codex_gpt61_medium` | `gpt-6.1-sol` (`medium`) | `arm2_full_skill` | 20 | **80.0%** | [60.0%, 95.0%] | **93.8%** | **12.5%** | **0.0%** | 0.0% | 478,333 | 14.5 | 115.7 |
| `codex_gpt61_high` | `gpt-6.1-sol` (`high`) | `arm0_no_skill` | 20 | 60.0% | [40.0%, 80.0%] | 88.5% | 18.8% | 6.2% | 0.0% | 277,955 | 7.3 | 134.9 |
| `codex_gpt61_high` | `gpt-6.1-sol` (`high`) | `arm1_standards_only` | 20 | 80.0% | [60.0%, 95.0%] | 90.2% | 18.8% | 6.2% | 0.0% | 515,299 | 15.0 | 171.7 |
| `codex_gpt61_high` | `gpt-6.1-sol` (`high`) | `arm2_full_skill` | 20 | **85.0%** | [70.0%, 100.0%] | **96.5%** | **12.5%** | **0.0%** | 0.0% | 647,854 | 18.8 | 169.1 |

### 2.2 So sánh Thống kê Ghép cặp & Reasoning Effort Scaling (`M1`)

- **Hiệu ứng tri thức chuẩn (`Arm 0 → Arm 1`):** Tăng từ `+10.00 pp` đến `+30.00 pp` trên cả `5/5` cấu hình (đạt $p = 0.0312$ theo McNemar exact trên `codex_gpt61_low` và `medium`).
- **Hiệu ứng tổng thể (`Arm 0 → Arm 2`):** Tăng `+10.00` đến `+35.00 pp`, Normalized Gain $g \in [0.3333, 0.6364]$ (`codex_gpt61_medium`: McNemar $p = 0.0156$, Wilcoxon Checkpoint $p = 0.0156$; `codex_gpt61_high`: $p = 0.0625$).
- **Effort Scaling trên `gpt-6.1-sol` (`low → medium → high` ở `Arm 2`):**
  - `pass@1`: `75.0% → 80.0% → 85.0%`
  - `Checkpoint Mean`: `90.0% → 93.8% → 96.5%`
- **Phát hiện chẩn đoán từ `M1`:** (i) Nhánh `Arm 2` có chi phí token cao (`~2.2–2.9x` trên `ag_flash`/`ag_pro`) vì đọc nhiều reference và render toàn bộ trang; (ii) `codex_gpt61_low` tụt xuống `25.0%` ở Nhóm E (`75.0%` trên `Arm 1`) do lỗi thứ tự thẻ con OOXML khi tự sửa part ngoài phạm vi `layout-edit`; (iii) `VAL-B03` và `VAL-E02` rớt `100%` trên cả `5` cấu hình, dẫn tới cuộc kiểm tra hậu kiểm phát hiện 2 errata dưới đây.

---

## 3. Errata `val-v1.1` (Minh bạch Hậu kiểm) & Re-score Milestone #1

Để bảo đảm tách biệt hoàn toàn ảnh hưởng của lỗi đề/oracle khỏi nâng cấp hệ thống, 2 errata được ghi nhận công khai tại [`val_v1/PROVENANCE.md`](val_v1/PROVENANCE.md) (số liệu Milestone #1 gốc ở Mục 2 giữ nguyên không sửa):

1. **`VAL-B03_party_mof_mac_spacing`:** Khoảng từ đoạn mở đầu `'Thực hiện kế hoạch số 05- KH/ĐUK...'` (`P5`) đến hết đoạn `'Giảm 01 chi bộ so 2011...'` (`P15`) gồm đúng **11 đoạn** (`P5..P15`), còn `P16` là tiêu đề mục II (`'II. Kết quả...'`) vốn bị prompt yêu cầu giữ nguyên. Đã sửa `prompt_vi` (`"12 đoạn"` → `"11 đoạn"`) và `target_paragraph_ordinals` (`[5..15]`).
2. **`eval_harness/diff_ntpr.py` trên `VAL-E02_neu_footer_page_label`:** Khi tách liên kết footer của một section để chỉ thêm nhãn `'Trang '` vào riêng `word/footer2.xml` mà vẫn bảo toàn chữ `'i'` ở section sau qua part mới `word/footer3.xml`, kiểm tra `hf_text_ok` trước đó báo lỗi sai lệch. Đã sửa để chấp nhận part `header/footer` mới sinh nếu giữ nguyên `text_nfc` của part trong `allow_header_footer_parts`.

---

## 4. Milestone #2 (`val-v1.1`, `SKILL.md v0.4.1` + `normalize-ooxml` — 120 lượt hoàn chỉnh)

- **Registry ID:** `val_v1_milestone_2_20261126` (`milestone_index = 2`)
- **Commit gốc code `val-v1.1` + `v0.4.1`:** trên nhánh `main` (kế thừa từ `abe0675`) | **`SKILL.md` SHA-256:** `9103a7e50f0544883f3e582abc07d8141117e2a3475734aa88d9d0b3581ac6ff`
- **Nguồn báo cáo:** [`research/eval/val_v1/milestone_2/MILESTONE-2-REPORT.md`](../research/eval/val_v1/milestone_2/MILESTONE-2-REPORT.md)

| Cấu hình | Nhánh | Giai đoạn | `pass@1` | Ckpt Mean | Collateral | Over-Refusal | Nhóm E `pass@1` | `ooxml_order_clean` | Total Tok | Calls | Time (s) |
|---|---|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|
| `ag_flash` | `arm0_no_skill` | M1 (v1.0) | 65.0% | 89.4% | 18.8% | 6.2% | 75.0% | 75.0% | 430,683 | 13.9 | 133.2 |
| `ag_flash` | `arm0_no_skill` | M1 (v1.1 re-score) | 75.0% | 91.9% | 12.5% | 6.2% | 100.0% | 75.0% | 430,683 | 13.9 | 133.2 |
| `ag_flash` | `arm0_no_skill` | **M2 (v1.1+v0.4.1)** | 75.0% | 87.9% | 12.5% | 12.5% | 100.0% | 70.0% | 377,646 | 12.7 | 102.9 |
| `ag_flash` | `arm1_standards_only` | M1 (v1.0) | 80.0% | 91.8% | 12.5% | 6.2% | 75.0% | 75.0% | 523,682 | 18.1 | 129.9 |
| `ag_flash` | `arm1_standards_only` | M1 (v1.1 re-score) | 90.0% | 94.3% | 6.2% | 6.2% | 100.0% | 75.0% | 523,682 | 18.1 | 129.9 |
| `ag_flash` | `arm1_standards_only` | **M2 (v1.1+v0.4.1)** | 85.0% | 97.5% | 6.2% | 0.0% | 100.0% | 80.0% | 479,038 | 16.9 | 106.9 |
| `ag_flash` | `arm2_full_skill` | M1 (v1.0) | 85.0% | 92.5% | 12.5% | 0.0% | 75.0% | 75.0% | 1,158,247 | 31.7 | 174.5 |
| `ag_flash` | `arm2_full_skill` | M1 (v1.1 re-score) | 95.0% | 95.0% | 6.2% | 0.0% | 100.0% | 75.0% | 1,158,247 | 31.7 | 174.5 |
| `ag_flash` | `arm2_full_skill` | **M2 (v1.1+v0.4.1)** | **95.0%** | **95.0%** | **6.2%** | 6.2% | **100.0%** | 75.0% | **1,002,925** | **28.6** | **145.9** |
| `codex_gpt61_low` | `arm0_no_skill` | M1 (v1.0) | 50.0% | 81.8% | 18.8% | 6.2% | 50.0% | 50.0% | 139,726 | 4.3 | 277.4 |
| `codex_gpt61_low` | `arm0_no_skill` | M1 (v1.1 re-score) | 55.0% | 83.1% | 18.8% | 6.2% | 50.0% | 50.0% | 139,726 | 4.3 | 277.4 |
| `codex_gpt61_low` | `arm0_no_skill` | **M2 (v1.1+v0.4.1)** | 55.0% | 80.9% | 18.8% | 12.5% | 50.0% | 45.0% | 143,773 | 4.9 | 71.2 |
| `codex_gpt61_low` | `arm1_standards_only` | M1 (v1.0) | 80.0% | 87.5% | 18.8% | 12.5% | 75.0% | 70.0% | 294,414 | 8.8 | 156.8 |
| `codex_gpt61_low` | `arm1_standards_only` | M1 (v1.1 re-score) | 90.0% | 90.0% | 12.5% | 12.5% | 100.0% | 70.0% | 294,414 | 8.8 | 156.8 |
| `codex_gpt61_low` | `arm1_standards_only` | **M2 (v1.1+v0.4.1)** | 85.0% | 87.5% | 18.8% | 12.5% | 75.0% | 70.0% | 299,871 | 8.9 | 88.1 |
| `codex_gpt61_low` | `arm2_full_skill` | M1 (v1.0) | 75.0% | 90.0% | 12.5% | 6.2% | 25.0% | 65.0% | 325,936 | 10.4 | 142.6 |
| `codex_gpt61_low` | `arm2_full_skill` | M1 (v1.1 re-score) | 85.0% | 92.5% | 6.2% | 6.2% | 50.0% | 65.0% | 325,936 | 10.4 | 142.6 |
| `codex_gpt61_low` | `arm2_full_skill` | **M2 (v1.1+v0.4.1)** | **95.0%** | **95.0%** | **6.2%** | 6.2% | **100.0%** | **75.0%** | **284,509** | **9.2** | **70.4** |

> [!NOTE]
> - **Mức cải thiện kỹ thuật thuần (`M2` vs `M1 Re-scored v1.1` trên `Arm 2`):**
>   - `codex_gpt61_low`: `pass@1` tăng **`+10.0 pp`** (`85.0% → 95.0%`; McNemar `0 → 2` **$p = 0.0078$**, Normalized Gain $g = 0.8889$), Nhóm E tăng **`+50.0 pp`** (`50.0% → 100.0%`), `ooxml_order_clean` tăng **`+10.0 pp`**, Total Tokens giảm **`-12.7%`** (xuống thấp hơn cả `Arm 1`: `284,509` vs `299,871`), Calls giảm **`-12.0%`**, thời gian trung bình giảm từ `142.6s → 70.4s` (nhanh hơn cả `Arm 0` `71.2s`).
>   - `ag_flash`: giữ vững **`95.0%`** (`100.0%` trên cả 4 nhóm B, C, D, E; $g_{0\to2} = 0.8000$), giảm **`-13.4%`** Total Tokens và **`-9.6%`** Calls.
> - **Ghi chú `codex_gpt61_medium` M2:** Chạy được `44/60` (`rc=0`, gồm `20/20` `Arm 0`, `20/20` `Arm 1`, `4/20` `Arm 2`) trước khi chạm hạn mức tuần Codex (`"try again at Nov 30th, 2026 12:55 PM"`), không tính vào bảng `120` lượt hoàn chỉnh của Milestone #2.

---

## 5. Giới hạn Hiện tại & Kiểm định Độc lập Tiếp theo (`test-sealed-v1`)

- **Ngân sách `val-v1`:** Đã dùng `2/3` slot ([`logs/val-run-registry.jsonl`](logs/val-run-registry.jsonl)); còn `1` slot nếu muốn chạy bổ sung `codex_gpt61_medium` / `high` cho `v0.4.1`.
- **Chấm trực quan mù (Blind Human Visual Audit):** Đã xuất sẵn bộ `35` mẫu PNG trước/sau của `12` ca phân tầng từ `ag_flash` M2 kèm phiếu chấm [`scoring_sheet.csv`](../research/eval/val_v1/milestone_2/blind_audit/scoring_sheet.csv) và giao diện [`viewer.html`](../research/eval/val_v1/milestone_2/blind_audit/viewer.html) (`0` nhãn Arm/model bị lộ).
- **Tập khóa `test-sealed-v1` (30 ca):** Giữ nguyên trạng thái `pending_phase3_curation` tại [`SEALED-TEST-COMMITMENT.json`](SEALED-TEST-COMMITMENT.json) bên ngoài workspace để tránh rò rỉ dữ liệu. Quy trình tuyển chọn, kiểm định SHA-256 (`seal-commitment` / `verify-commitment`), chạy 1 lần duy nhất (`mark-evaluated`), và xuất bộ chấm mù được chuẩn hóa đầy đủ tại [`val_v1/SEALED-CURATION-GUIDE.md`](val_v1/SEALED-CURATION-GUIDE.md).
