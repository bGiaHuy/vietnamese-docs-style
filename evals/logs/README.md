# Validation Run Registry (`evals/logs/`)

Thư mục này lưu nhật ký append-only cho toàn bộ các đợt chạy trên tập `val-v1` tại [`val-run-registry.jsonl`](val-run-registry.jsonl) nhằm chống *Sequential Overfitting*.

## Quy định ghi nhận
- Mỗi dòng trong `val-run-registry.jsonl` là một bản ghi JSON cho **1 milestone run** (tối đa **3 milestone runs** được phép cho toàn bộ giai đoạn phát triển trên `val-v1`).
- Không được xóa, sửa đè, hay cherry-pick các dòng đã ghi.
- Trường tối thiểu của mỗi dòng JSONL:
  - `run_id`: chuỗi định danh duy nhất (vd. `val-v1-milestone-1`)
  - `milestone_index`: số nguyên `1`, `2`, hoặc `3`
  - `timestamp`: ISO-8601
  - `git_commit_sha`: commit SHA hệ thống được đánh giá
  - `skill_md_sha256`: SHA-256 của `SKILL.md`
  - `model_ids`: danh sách model IDs chính xác
  - `arms`: danh sách nhánh đã chạy (vd. `["arm0_no_skill", "arm1_standards_only", "arm2_full_skill"]`)
  - `trials_per_case`: số lần lặp `k`
  - `summary`: tóm tắt `pass_at_1`, `pass_hat_k`, `over_refusal_rate`, `unsafe_edit_rate`, `contamination_count`
  - `notes`: tóm tắt lỗi phát hiện và hành động (nếu có) trước milestone kế tiếp
