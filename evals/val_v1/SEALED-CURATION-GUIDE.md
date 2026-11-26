# Hướng dẫn Tuyển chọn & Niêm phong Độc lập `test-sealed-v1` (`SEALED-CURATION-GUIDE.md`)

Tài liệu này hướng dẫn người tuyển chọn độc lập (hoặc cộng tác viên bên ngoài không làm việc trên mã nguồn `vndocs`) cách tạo tập **`test-sealed-v1` (30 cases)** mà **không làm rò rỉ nội dung hay đáp án** vào workspace của agent phát triển.

---

## 1. Quy tắc Cách ly Môi trường (Air-Gap Rule)

1. Đặt toàn bộ tệp nguồn `.docx`, ảnh đính kèm `.png`, và `test_sealed_v1_cases.json` ở **thư mục bên ngoài repository** (ví dụ `../sealed-eval-v1/` hoặc máy đánh giá tách biệt), **không** lưu hay commit vào repo phát triển `vietnamese-docs-style`.
2. Người tạo câu hỏi **không** dùng `scripts/build_docx.py`, `scripts/generate_docx.py`, `assets/templates/` hay chạy agent đã đọc `SKILL.md` để sinh đề hoặc kiểm tra đáp án bằng `vndocs`.

---

## 2. Phân bổ Quota Bắt buộc cho `test-sealed-v1` (30 Cases)

Phải khớp [`../EVALUATION-PROTOCOL.md`](../EVALUATION-PROTOCOL.md) và [`../schemas/independent-case.schema.json`](../schemas/independent-case.schema.json):

- **Tổng số ca:** `30` (`case_id` dạng `SEAL-A01_<slug>` .. `SEAL-E06_<slug>`, `split: "test-sealed-v1"`).
- **Nhóm taxonomy (6 ca mỗi nhóm):** `A`: 6, `B`: 6, `C`: 6, `D`: 6, `E`: 6.
- **Bậc độ dài:** `short`: 8 (1–3 trang), `medium`: 9 (4–10 trang), `long`: 7 (11–25 trang), `very_long`: 6 (> 25 trang hoặc > 150 KB `word/document.xml`).
- **Loại lệnh:** ≥ 40% `descriptive` (khuyến nghị 15) và ≥ 40% `inferential` (khuyến nghị 15).
- **Kết quả kỳ vọng:** `refuse` từ 4 đến 7 ca (khuyến nghị 6), còn lại là `edit` (≥ 3 ca `near_boundary_valid: true`).
- **Đa dạng phần mềm soạn thảo:** bao gồm tài liệu từ Microsoft Word, Google Docs export, LibreOffice Writer, và/hoặc WPS Office.
- **Độc lập nguồn (Disjoint Clusters):** toàn bộ `source_cluster_id` và `source_docx_sha256` phải khác 100% với 20 ca trong [`PROVENANCE.md`](PROVENANCE.md) và [`val_v1_cases.json`](val_v1_cases.json).

---

## 3. Quy trình 5 Bước Kiểm định & Niêm phong (`eval_harness.curator`)

### Bước 1 — Viết manifest và kiểm định Zero-Leakage & Quotas
Tham khảo cấu trúc mẫu tổng hợp tại [`../schemas/sealed-manifest-example.json`](../schemas/sealed-manifest-example.json) (khớp [`../schemas/independent-case.schema.json`](../schemas/independent-case.schema.json)), rồi kiểm định tệp manifest thực tế:
```powershell
python -m eval_harness.curator validate-manifest `
  --manifest <SEALED_DIR>/test_sealed_v1_cases.json `
  --corpus-dir <SEALED_DIR>/corpus
```

### Bước 2 — Kiểm tra Tách nguồn hoàn toàn với `val-v1`
```powershell
python -m eval_harness.curator check-disjoint `
  --val-manifest evals/val_v1/val_v1_cases.json `
  --sealed-manifest <SEALED_DIR>/test_sealed_v1_cases.json
```
Yêu cầu kết quả: `"disjoint": true`, `"overlapping_clusters": []`, `"overlapping_hashes": []`.

### Bước 3 — Cổng Khả năng Giải Độc lập (*Independent Solvability Gate*)
Chuẩn bị thư mục đối chứng độc lập (viết hoàn toàn bằng `python-docx` / `lxml` / sửa tay trong LibreOffice, **không** dùng `vndocs`):
- `<SEALED_DIR>/solvability/positive/<case_id>/` chứa `<case_id>.docx` & `<case_id>.result.json` đạt yêu cầu.
- `<SEALED_DIR>/solvability/negative/<case_id>/` chứa bản lỗi phổ biến (ví dụ quên đồng bộ `tcW`, tràn lề phải, sai lề Đảng, từ chối nhầm hoặc sửa lén ca `refuse`, không cập nhật nested table/footnote/footer).

Chạy cổng kiểm định:
```powershell
python -m eval_harness.curator verify-solvability `
  --manifest <SEALED_DIR>/test_sealed_v1_cases.json `
  --corpus-dir <SEALED_DIR>/corpus `
  --positive-dir <SEALED_DIR>/solvability/positive `
  --negative-dir <SEALED_DIR>/solvability/negative `
  --output <SEALED_DIR>/solvability-report.json
```
Chỉ chấp nhận khi `all_passed == true` (30/30 positive đạt `hit = 1`, 30/30 negative bị `hit = 0`).

### Bước 4 — Niêm phong SHA-256 Commit Công khai
Sau khi đóng băng `test-sealed-v1` và đẩy `SKILL.md` / `vndocs` lên git tag cuối (ví dụ `eval-freeze-v1`):
```powershell
python -m eval_harness.curator seal-commitment `
  --sealed-dir <SEALED_DIR> `
  --sealed-manifest-relpath test_sealed_v1_cases.json `
  --val-manifest evals/val_v1/val_v1_cases.json `
  --output evals/SEALED-TEST-COMMITMENT.json `
  --frozen-tag eval-freeze-v1 `
  --frozen-sha <GIT_COMMIT_SHA>
```
Chỉ file băm SHA-256 `evals/SEALED-TEST-COMMITMENT.json` được ghi vào repository; toàn bộ nội dung `test-sealed-v1` giữ ngoài workspace cho tới khi chạy đánh giá chính thức **1 lần duy nhất**.

### Bước 5 — Kiểm chứng Niêm phong & Chạy Đánh giá 1 Lần Duy nhất
Ngay trước khi chạy benchmark chính thức trên bản tag đã đóng băng (`eval-freeze-v1`), kiểm tra lại toàn vẹn SHA-256 và đảm bảo tập niêm phong chưa từng được mở:
```powershell
python -m eval_harness.curator verify-commitment `
  --commitment evals/SEALED-TEST-COMMITMENT.json `
  --sealed-dir <SEALED_DIR> `
  --sealed-manifest-relpath test_sealed_v1_cases.json
```
Chạy đánh giá 3 nhánh (thực hiện bên ngoài quy trình phát triển, sau đó đóng dấu đã đánh giá 1 lần):
```powershell
python -m eval_harness.runner run `
  --manifest <SEALED_DIR>/test_sealed_v1_cases.json `
  --corpus-dir <SEALED_DIR>/corpus `
  --out-dir <SEALED_EVAL_OUT_DIR> `
  --sandbox danger-full-access `
  --workers 3

python -m eval_harness.curator mark-evaluated `
  --commitment evals/SEALED-TEST-COMMITMENT.json `
  --notes "Evaluated once on tag eval-freeze-v1"
```

### Bước 6 — Tạo Bộ Chấm Trực quan Mù (*Blind Audit Pack*)
Dùng CLI xuất gói PNG trước/sau ẩn toàn bộ tên Arm/model kèm phiếu chấm `scoring_sheet.csv` và giao diện HTML tương tác `viewer.html` cho người chấm độc lập:
```powershell
python -m eval_harness.curator prepare-blind-audit `
  --report <OUT_DIR>/benchmark_report.json `
  --manifest <MANIFEST_JSON> `
  --corpus-dir <CORPUS_DIR> `
  --out-dir <OUT_DIR>/blind_audit `
  --max-cases 12 `
  --seed 42
```
