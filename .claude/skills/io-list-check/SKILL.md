---
name: io-list-check
description: Kiểm tra chất lượng Instrument I/O List (Signal I/O List) dạng PDF có lớp text hoặc Excel trước khi nộp/duyệt — trùng loop tag, lệch bảng I/O Summary, kiểu tín hiệu ngoài bảng viết tắt, thiếu range/burnout/N.O-N.C, tag sai ISA-5.1 so với kiểu tín hiệu, ô còn Hold/TBC, trùng kênh Node/Slot/Channel. Dùng khi user gửi I/O list và yêu cầu check, review, kiểm tra, rà soát, đối chiếu I/O count, hoặc chuẩn bị trả lời comment về I/O list. Trigger: "I/O list", "IO list", "signal list", "danh mục tín hiệu", "check I/O", "rà soát I/O", "I/O summary", "I/O count".
---

# io-list-check

Script làm toàn bộ việc đọc và kiểm tra. **Không đọc PDF bằng ảnh** — chỉ chạy script và đọc tóm tắt nó in ra (tiết kiệm token).

## Cách chạy

```bash
pip install -q pdfplumber openpyxl   # nếu thiếu
python .claude/skills/io-list-check/scripts/check.py <file.pdf|xlsx> --out <thư mục> [--tag-scope document|section] [--map "Tên Summary=Tên section"]
```

- Output: `Claude_IOCheck_<tên file>.xlsx` gồm 3 sheet `Findings` (sắp theo mức độ), `Summary_Recount`, `Data` (bảng đã chuẩn hoá — dùng tiếp cho loop diagram, MTO).
- Màn hình in tóm tắt: số tín hiệu, số finding theo rule/mức, 15 mục High/Medium đầu.

## Trước khi chạy — hỏi user nếu chưa rõ

1. **Phạm vi duy nhất của loop tag** (`--tag-scope`): `document` (mặc định, chặt) hay `section` (dự án cho phép cùng tag ở section Hard-wired và Communication). Không chắc → hỏi.
2. PDF phải có lớp text. Nếu `Tín hiệu: 0` → có thể là PDF scan hoặc header lạ: báo user, đừng tự OCR hàng loạt.

## Đọc kết quả

| Rule | Nội dung | Mức |
|---|---|---|
| R01 | Trùng loop tag (theo `tag_scope`) | High / Info |
| R02 | Đếm lại theo section × type so với I/O Summary; tự cộng hàng của bảng Summary | High / Medium |
| R03 | Kiểu tín hiệu không có trong bảng ABBREVIATIONS của chính tài liệu | Medium |
| R04 | Logic theo type: AI/AO range + unit, AI burnout U/D, DI/DO spec N.O/N.C, contact M/S, voltage | Medium / Low |
| R05 | Chữ cái tag ISA-5.1 ↔ type (vd `LSH` mà AI, `XSP` mà AI) | Medium |
| R06 | Ô còn Hold/TBC | Info |
| R07 | Trùng Node/Slot/Channel; card chứa nhiều kiểu tín hiệu (chỉ khi đã gán kênh) | High / Medium |

Chi tiết rule, logic ISA-5.1 và cách chỉnh: `references/rules.md`. Ngưỡng/danh mục chỉnh trong `references/rules.json`.

- R02 in ra **mapping Summary → section** trong sheet `Summary_Recount`: luôn kiểm tra mapping đúng trước khi kết luận lệch; sai → chạy lại với `--map`.
- R05 là **gợi ý kiểm tra**, không phải lỗi chắc chắn: tag theo quy ước riêng dự án (vd `ZTRO` = travelling to open) sẽ bị báo. Trình bày cho user như "cần verify".
- Báo cáo cho user: nêu High trước, gom theo rule, dẫn vị trí (`p.<trang> #<STT>` hoặc `Sheet!R<hàng>`). Không tự sửa file I/O list gốc.

## Test

```bash
python -m pytest .claude/skills/io-list-check/tests -q          # Excel giả lập có lỗi cài sẵn
IO_SAMPLE_PDF=<file.pdf> python -m pytest .claude/skills/io-list-check/tests -q   # thêm PDF thật
```
