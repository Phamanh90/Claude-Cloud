# io-list-check — Chi tiết rule

Mọi danh mục/ngưỡng nằm trong `rules.json`; sửa JSON, không cần sửa code. Có thể truyền file JSON riêng cho từng dự án bằng `--config`.

## Đọc dữ liệu (extract.py)

- PDF: `pdfplumber.extract_tables()` từng trang. Excel: từng sheet (`data_only`, nên công thức phải đã được tính trong file).
- **Section**: chuỗi trong ngoặc ở tiêu đề, dạng `(DCS TANK / HARD-WIRED - …)` → lấy phần trước ` - ` = `DCS TANK / HARD-WIRED`. Bắt buộc có `/` để không nhầm với `(DCS)` trong header.
- **Cột** map theo từ khoá header (Anh + Việt), không cố định vị trí (`COLUMN_KEYS` trong `extract.py`). Bảng hợp lệ khi có tối thiểu `loop_tag` + `type`.
- **Hàng dữ liệu**: ô đầu là số thứ tự.
- **I/O Summary**: bảng có cột `SYSTEM` và ≥3 cột kiểu tín hiệu. Bảng toàn cột `COM-*` = bảng communication.
- **Abbreviations**: trang/sheet có chữ `ABBREVIATION`/`VIẾT TẮT`, regex `XX : Analog/Discrete/Pulse…`. Không thấy → dùng `signal_types_default`.

## R01 — Trùng loop tag
- `tag_scope = document`: tag xuất hiện >1 lần trong cả tài liệu → High.
- `tag_scope = section`: trùng trong cùng section → High; dùng ở nhiều section → Info (để vẫn thấy).
- Field tag **không** kiểm tra trùng (1 thiết bị có nhiều tín hiệu là bình thường).

## R02 — I/O Summary
- Hard-wired: map từng hàng Summary sang section bằng token (bỏ DCS/RIO/AREA/HARD/WIRED, `TANKS→TANK`); số hiệu khu vực (VC-5002…) phải khớp tuyệt đối; chọn section trùng nhiều token nhất, ít token thừa nhất. Sai → `--map "RIO-VC-Refueller=DCS HYDRANT_REFUELLER AREA / HARD-WIRED"`.
- Communication: Summary thường chia theo giao thức (Modbus MOV/TGS…) không trùng section → chỉ đối chiếu **hàng Total** với tổng tín hiệu `COM-*`.
- Tự kiểm tra cộng: tổng các hàng ≠ hàng Total → Medium.
- Chỉ đếm các type trong `summary_types_*`; type lạ (`MOD`, `-`) không được đếm và đã báo ở R03.

## R03 — Kiểu tín hiệu
Type không có trong bảng Abbreviations của chính tài liệu → Medium (gom theo section + type).

## R04 — Logic theo type (chỉ hard-wired, trừ spec DI/DO áp cho cả COM)
| Type | Kiểm tra |
|---|---|
| AI, AO | lower/upper là số, lower < upper (hỗ trợ `0,0` và `0.0`); có unit. Hold/TBC → để R06 |
| AI | burnout ∈ `burnout_values` (U/D) |
| AO, DI, DO, PI | không nên có burnout U/D (Low) |
| DI, DO (+COM) | spec ∈ N.O/N.C (bỏ dấu chấm, khoảng trắng) |
| DI, DO | contact ∈ M/S; có circuit voltage |

## R05 — ISA-5.1: chữ cái tag ↔ type
Tách tiền tố chữ của loop tag (`tag_regex`), bỏ tag trong `non_isa_prefixes` hoặc không khớp regex (báo Info tổng hợp). Bỏ modifier của chữ đầu (`D F J K Q`). Xét các chữ còn lại theo thứ tự, gặp điều kiện đầu tiên thì dừng:

| Điều kiện | Type hợp lệ | Ví dụ |
|---|---|---|
| `SV`/`SE` ngay sau chữ đầu (safety) | DI, DO, AI | PSV, PSE |
| có `S` (switch), chữ đầu `H` | DI, DO | HS, HSS, HSR, HSO |
| có `S` (switch), chữ đầu khác | DI | LSH, ZSO, ZSC, XSP |
| có `C` (controller) | AI, AO, DO | PIC, FIC, PICA |
| có `V`/`Y`/`Z` (final element/relay/actuator) | AO, DO | FV, XY, HZ |
| có `T`/`E`/`I`/`R` | AI, PI | PT, TE, PIA, TI, FQI |
| chỉ có `A` (alarm) | AI, DI | TAH, LAH, XA |

So theo type gốc (`COM-DI` → `DI`). Kết quả là **gợi ý kiểm tra** (Medium), gom theo (tiền tố, type).

## R06 — Hold/TBC
Ô có giá trị trong `placeholder_values` → Info, gom theo section + cột. Dùng để theo dõi dữ liệu còn thiếu qua các Rev.

## R07 — Kênh I/O
Chỉ chạy khi có cột Channel được điền.
- Trùng (Node, Slot, Channel) → High.
- Một card (Node, Slot) chứa nhiều type → Medium.
- **Chưa có**: kiểm tra vượt số kênh card, spare % — cần model card + số kênh tra từ catalog hãng (Siemens…); bổ sung khi user cung cấp.
