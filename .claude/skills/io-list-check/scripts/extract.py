"""
extract.py — Đọc I/O list (PDF có lớp text hoặc Excel) thành dữ liệu chuẩn hoá.

Ý tưởng: mọi nguồn đều được quy về "grid" (danh sách hàng, mỗi hàng là list ô),
kèm vị trí (trang PDF hoặc sheet Excel). Sau đó dùng chung một bộ phân tích:
  - tìm tiêu đề section "(DCS TANK / HARD-WIRED - ...)"
  - map cột theo header (không cố định vị trí cột)
  - lấy hàng dữ liệu (ô đầu là số thứ tự)
  - tách bảng I/O SUMMARY và bảng chữ viết tắt (ABBREVIATIONS)
"""
import re
from dataclasses import dataclass, field
from pathlib import Path

# ---------------------------------------------------------------------------
# Map cột: tên chuẩn -> từ khoá trong header. Thứ tự quan trọng: khoá cụ thể
# (CONTROL ACTION) phải đứng trước khoá chung (ACTION).
# "leaf" = so khớp với ô header thấp nhất của cột; "full" = toàn bộ chữ header.
# ---------------------------------------------------------------------------
COLUMN_KEYS = [
    ("alarm_al", "leaf", ["AL"]),
    ("alarm_hh", "leaf", ["HH"]),
    ("alarm_h", "leaf", ["H"]),
    ("alarm_l", "leaf", ["L"]),
    ("alarm_ll", "leaf", ["LL"]),
    ("no", "full", ["SỐ TT", "S/N", "NO."]),
    ("field_tag", "full", ["FIELD", "INSTRUMENT TAG", "TÊN THIẾT BỊ"]),
    ("loop_tag", "full", ["LOOP TAG", "TÊN LOOP"]),
    ("service", "full", ["SERVICE", "PHỤC VỤ"]),
    ("from", "full", ["FROM"]),
    ("to", "full", ["TO/", "TO /", "TỚI"]),
    ("pid", "full", ["P&ID", "PID"]),
    ("type", "full", ["TYPE", "LOẠI"]),
    ("spec", "full", ["SPEC", "ĐẶC TÍNH"]),
    ("contact", "full", ["CONTACT"]),
    ("control_action", "full", ["CONTROL ACTION", "CONTROL"]),
    ("action", "full", ["ACTION", "TÁC ĐỘNG"]),
    ("voltage", "full", ["VOLTAGE", "ĐIỆN ÁP"]),
    ("lower", "full", ["LOWER", "THẤP"]),
    ("upper", "full", ["UPPER", "CAO"]),
    ("unit", "full", ["UNIT", "ĐƠN VỊ"]),
    ("node", "full", ["NODE"]),
    ("slot", "full", ["SLOT"]),
    ("channel", "full", ["CHANNEL"]),
    ("interlock", "full", ["INTERLOCK", "KHÓA LIÊN ĐỘNG"]),
    ("burnout", "full", ["BURNOUT"]),
    ("remarks", "full", ["REMARK", "GHI CHÚ"]),
]
# Header phải có tối thiểu các cột này mới được coi là bảng I/O
REQUIRED_COLS = {"loop_tag", "type"}

# Tiêu đề section luôn có dạng "(HỆ / LOẠI ...)" — bắt buộc có "/" để không nhầm với "(DCS)" trong header
TITLE_RE = re.compile(r"\(\s*((?:DCS|PLC|SIS|ESD|F&G|FGS|RIO|BPCS)[^()]*?/[^()]*?)\s*\)")
ABBR_RE = re.compile(r"\b(COM-[A-Z]{2}|[A-Z]{2})\s*:\s*(Analog|Pulse|Discrete|Digital)", re.I)


@dataclass
class Grid:
    rows: list          # list[list[str]]
    loc: str            # "p.5" hoặc "Sheet1"
    page_text: str = ""  # toàn bộ chữ của trang (PDF) / sheet (Excel)
    row_locs: list = None  # vị trí từng hàng (Excel: số hàng thật)


@dataclass
class Record:
    loc: str            # vị trí để người dùng tra lại: "p.5 #21" / "IO!R30"
    section: str        # "DCS TANK / HARD-WIRED"
    data: dict = field(default_factory=dict)

    def get(self, k):
        return self.data.get(k, "")


@dataclass
class IOList:
    source: str
    records: list
    summary: list        # [{"table": "hardwired|comm", "name", "counts": {type: int}, "loc"}]
    abbreviations: set   # các kiểu tín hiệu khai báo trong tài liệu
    sections: list       # tên section theo thứ tự xuất hiện


def clean(v):
    """Chuẩn hoá ô: None -> '', gộp xuống dòng, bỏ khoảng trắng thừa."""
    if v is None:
        return ""
    s = str(v).replace("\n", " ").replace("\r", " ")
    return re.sub(r"\s+", " ", s).strip()


# ---------------------------------------------------------------------------
# Đọc nguồn -> grids
# ---------------------------------------------------------------------------
def grids_from_pdf(path):
    import pdfplumber
    grids = []
    with pdfplumber.open(path) as pdf:
        for i, page in enumerate(pdf.pages, start=1):
            text = page.extract_text() or ""
            tables = page.extract_tables()
            if not tables:
                grids.append(Grid([], f"p.{i}", text))
            for t in tables:
                grids.append(Grid([[clean(c) for c in r] for r in t], f"p.{i}", text))
    return grids


def grids_from_xlsx(path):
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    grids = []
    for ws in wb.worksheets:
        rows, locs = [], []
        for r_idx, r in enumerate(ws.iter_rows(values_only=True), start=1):
            rows.append([clean(c) for c in r])
            locs.append(r_idx)
        text = "\n".join(" ".join(c for c in r if c) for r in rows)
        grids.append(Grid(rows, ws.title, text, locs))
    return grids


# ---------------------------------------------------------------------------
# Phân tích grid
# ---------------------------------------------------------------------------
def _is_data_row(row):
    return bool(row) and re.fullmatch(r"\d+", row[0] or "") is not None


def _section_title(grid):
    """Lấy tiêu đề section từ các hàng trên header hoặc chữ trên trang."""
    for r in grid.rows[:6]:
        for c in r:
            m = TITLE_RE.search(c or "")
            if m:
                return _short_section(m.group(1))
    m = TITLE_RE.search(grid.page_text or "")
    return _short_section(m.group(1)) if m else ""


def _short_section(s):
    # "DCS TANK / HARD-WIRED - HỆ THỐNG ..." -> "DCS TANK / HARD-WIRED"
    return re.split(r"\s+-\s+", s)[0].strip()


def _map_columns(header_rows, ncols):
    """Ghép header nhiều tầng theo cột rồi so từ khoá -> {tên chuẩn: chỉ số cột}."""
    pieces = [[] for _ in range(ncols)]
    for r in header_rows:
        for j in range(min(ncols, len(r))):
            if r[j]:
                pieces[j].append(r[j].upper())
    mapping = {}
    for j in range(ncols):
        if not pieces[j]:
            continue
        leaf, full = pieces[j][-1], " ".join(pieces[j])
        for name, mode, keys in COLUMN_KEYS:
            if name in mapping:
                continue
            if mode == "leaf":
                if leaf in keys and len(pieces[j]) >= 1 and leaf == pieces[j][-1]:
                    mapping[name] = j
                    break
            elif any(k in full for k in keys):
                mapping[name] = j
                break
    return mapping


def _parse_io_grid(grid):
    """Trả về list Record nếu grid là bảng I/O, ngược lại []."""
    rows = grid.rows
    first = next((i for i, r in enumerate(rows) if _is_data_row(r)), None)
    if first is None:
        return []
    # Header = các hàng phía trên hàng dữ liệu đầu, bỏ hàng tiêu đề
    head = [r for r in rows[:first] if not any(TITLE_RE.search(c or "") or "I/O LIST" in (c or "").upper() for c in r)]
    ncols = max(len(r) for r in rows)
    cols = _map_columns(head, ncols)
    if not REQUIRED_COLS <= cols.keys():
        return []
    section = _section_title(grid) or grid.loc
    out = []
    for i, r in enumerate(rows[first:], start=first):
        if not _is_data_row(r):
            continue
        data = {k: (r[j] if j < len(r) else "") for k, j in cols.items()}
        if grid.row_locs:
            loc = f"{grid.loc}!R{grid.row_locs[i]}"
        else:
            loc = f"{grid.loc} #{data.get('no', '')}"
        out.append(Record(loc, section, data))
    return out


def _parse_summary_grid(grid, cfg):
    """Tách bảng I/O SUMMARY: hàng header có 'SYSTEM' + các cột kiểu tín hiệu."""
    hw, cm = set(cfg["summary_types_hardwired"]), set(cfg["summary_types_comm"])
    out, i = [], 0
    rows = grid.rows
    while i < len(rows):
        r = [c.upper() for c in rows[i]]
        types = {j: c for j, c in enumerate(r) if c in hw | cm}
        if any("SYSTEM" in c for c in r) and len(types) >= 3:
            table = "comm" if all(c in cm for c in types.values()) else "hardwired"
            name_col = next(j for j, c in enumerate(r) if "SYSTEM" in c)
            i += 1
            while i < len(rows) and rows[i] and rows[i][name_col]:
                name = rows[i][name_col]
                counts = {}
                for j, t in types.items():
                    v = rows[i][j] if j < len(rows[i]) else ""
                    counts[t] = int(float(v)) if re.fullmatch(r"\d+(\.0+)?", v or "") else 0
                loc = f"{grid.loc}!R{grid.row_locs[i]}" if grid.row_locs else grid.loc
                out.append({"table": table, "name": name, "counts": counts, "loc": loc})
                i += 1
            continue
        i += 1
    return out


def load(path, cfg):
    """Điểm vào chính: đọc file, trả về IOList."""
    p = Path(path)
    ext = p.suffix.lower()
    if ext == ".pdf":
        grids = grids_from_pdf(p)
    elif ext in (".xlsx", ".xlsm"):
        grids = grids_from_xlsx(p)
    else:
        raise ValueError(f"Định dạng không hỗ trợ: {ext} (chỉ .pdf, .xlsx, .xlsm)")

    records, summary, abbr, sections = [], [], set(), []
    seen_text = set()
    for g in grids:
        recs = _parse_io_grid(g)
        if recs:
            records += recs
            for r in recs:
                if r.section not in sections:
                    sections.append(r.section)
        else:
            summary += _parse_summary_grid(g, cfg)
        if g.page_text and g.page_text not in seen_text:
            seen_text.add(g.page_text)
            if "ABBREVIATION" in g.page_text.upper() or "VIẾT TẮT" in g.page_text.upper():
                abbr |= {m.group(1).upper() for m in ABBR_RE.finditer(g.page_text)}
    return IOList(str(p), records, summary, abbr, sections)
