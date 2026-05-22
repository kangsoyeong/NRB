import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "도급기성청구현황"

# ── 스타일 ──────────────────────────────────────────────────
thin = Side(style="thin")

def tb():
    return Border(left=thin, right=thin, top=thin, bottom=thin)

def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def font(bold=False, color="000000", size=10, italic=False):
    return Font(name="맑은 고딕", bold=bold, color=color, size=size, italic=italic)

def align(h="center", v="center", wrap=True):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

C_TITLE = "1F4E79"
C_HEAD1 = "2E75B6"
C_HEAD2 = "BDD7EE"
C_ALT   = "EBF3FA"
C_WHITE = "FFFFFF"
C_WARN  = "FFF2CC"   # 주의 강조
NUM_FMT = '#,##0'

# ── 열 너비 ─────────────────────────────────────────────────
col_widths = {
    "A": 7,   # 회차
    "B": 8,   # 정식/약식
    "C": 20,  # 품의번호
    "D": 13,  # 기안일자
    "E": 13,  # 처리일자(총괄표 기준)
    "F": 14,  # 청구월
    "G": 18,  # 금회금액(합계)
    "H": 17,  # 극동건설
    "I": 16,  # HJ중공업
    "J": 16,  # 서한
    "K": 16,  # 대저건설
    "L": 16,  # 엔알비
    "M": 18,  # 누계금액
    "N": 18,  # 비고
}
for col, w in col_widths.items():
    ws.column_dimensions[col].width = w

# ── 제목 ────────────────────────────────────────────────────
ws.merge_cells("A1:N1")
c = ws["A1"]
c.value     = "의왕초평 A-4BL  도급 기성청구 현황"
c.font      = Font(name="맑은 고딕", bold=True, size=15, color="FFFFFF")
c.alignment = align()
c.fill      = fill(C_TITLE)
c.border    = tb()
ws.row_dimensions[1].height = 38

# ── 계약금액 안내 ────────────────────────────────────────────
ws.merge_cells("A2:E2")
ws["A2"].value     = "계약금액 (도급기성청구 소계)"
ws["A2"].font      = font(bold=True, size=10)
ws["A2"].alignment = align(h="left")
ws["A2"].fill      = fill(C_HEAD2)
ws["A2"].border    = tb()

ws.merge_cells("F2:G2")
ws["F2"].value         = 86_000_000_000
ws["F2"].font          = font(bold=True, size=10)
ws["F2"].alignment     = align()
ws["F2"].fill          = fill(C_HEAD2)
ws["F2"].border        = tb()
ws["F2"].number_format = NUM_FMT

ws.merge_cells("H2:N2")
ws["H2"].value     = "※ 추가정산 포함 총계약 90,000,000,000원  /  단위 : 원, VAT포함"
ws["H2"].font      = font(italic=True, size=9, color="595959")
ws["H2"].alignment = align(h="left")
ws["H2"].fill      = fill(C_HEAD2)
ws["H2"].border    = tb()
ws.row_dimensions[2].height = 20

# ── 헤더 (3행) ───────────────────────────────────────────────
headers = [
    ("회차",             "A"),
    ("구분\n(정식/약식)", "B"),
    ("품의번호",          "C"),
    ("기안일자",          "D"),
    ("처리일자\n(총괄표 기준)", "E"),
    ("청구월",            "F"),
    ("금회금액\n(합계)",  "G"),
    ("극동건설(주)\n51%", "H"),
    ("(주)HJ중공업\n15%","I"),
    ("(주)서한\n14%",    "J"),
    ("(주)대저건설\n10%","K"),
    ("(주)엔알비\n10%",  "L"),
    ("누계금액",          "M"),
    ("비고",              "N"),
]
for label, col in headers:
    c = ws[f"{col}3"]
    c.value     = label
    c.font      = Font(name="맑은 고딕", bold=True, size=10, color="FFFFFF")
    c.alignment = align()
    c.fill      = fill(C_HEAD1)
    c.border    = tb()
ws.row_dimensions[3].height = 40

# ── 데이터 ──────────────────────────────────────────────────
# (회차, 정식/약식, 품의번호, 기안일자, 처리일자, 청구월,
#  합계, 극동, HJ, 서한, 대저, 엔알비, 누계, 비고)
data = [
    (
        "1회", "정식", "(주)엔알비-2025-17478",
        "2025-11-29", "2025-12-04", "2025년 12월",
        5_076_000_000,
        2_588_760_000, 761_400_000, 710_640_000, 507_600_000, 507_600_000,
        5_076_000_000,
        "기안일(11/29) vs 처리일(12/04) 월 상이"
    ),
    (
        "2회", "약식", "(주)엔알비-2026-1859",
        "2026-02-02", "2026-02-04", "2026년 2월",
        624_000_000,
        318_240_000, 93_600_000, 87_360_000, 62_400_000, 62_400_000,
        5_700_000_000,
        ""
    ),
    (
        "3회", "약식", "(주)엔알비-2026-4026",
        "2026-03-05", "2026-03-06", "2026년 3월",
        1_640_000_000,
        836_400_000, 246_000_000, 229_600_000, 164_000_000, 164_000_000,
        7_340_000_000,
        ""
    ),
    (
        "4회", "정식", "(주)엔알비-2026-5251",
        "2026-03-25", "2026-04-06", "2026년 4월",
        2_930_000_000,
        1_494_300_000, 439_500_000, 410_200_000, 293_000_000, 293_000_000,
        10_270_000_000,
        "기안일(3/25) vs 처리일(4/06) 월 상이"
    ),
    (
        "5회", "약식", "(주)엔알비-2026-7755",
        "미확인", "2026-05-12", "2026년 5월",
        3_980_000_000,
        2_029_800_000, 597_000_000, 557_200_000, 398_000_000, 398_000_000,
        14_250_000_000,
        "기안서 1p 누락 – 처리일은 총괄표 기준"
    ),
]

WARN_ROWS = {0, 3, 4}  # 1회·4회·5회 강조

for i, row_data in enumerate(data):
    row = 4 + i
    bg = C_WARN if i in WARN_ROWS else (C_ALT if i % 2 == 0 else C_WHITE)

    for col_idx, val in enumerate(row_data, start=1):
        c = ws.cell(row=row, column=col_idx, value=val)
        c.fill   = fill(bg)
        c.border = tb()

        if col_idx == 1:   # 회차
            c.font      = font(bold=True, size=11)
            c.alignment = align()
        elif col_idx == 2: # 정식/약식
            c.font      = font(bold=True, size=9,
                               color=("C00000" if val == "정식" else "2E75B6"))
            c.alignment = align()
        elif col_idx in (3, 4, 5, 14): # 품의번호, 날짜, 비고
            c.font      = font(size=9)
            c.alignment = align(h="left")
        elif col_idx == 6: # 청구월
            c.font      = font(bold=True, size=10)
            c.alignment = align()
        elif col_idx in range(7, 14):  # 금액
            c.font          = font(
                bold=(col_idx in (7, 13)),
                size=10,
                color=("C00000" if col_idx == 7 else "000000")
            )
            c.alignment     = align(h="right")
            c.number_format = NUM_FMT
        else:
            c.font      = font(size=9)
            c.alignment = align(h="left")

    ws.row_dimensions[row].height = 24

# ── 합계 행 ─────────────────────────────────────────────────
total_row = 4 + len(data)
ws.merge_cells(f"A{total_row}:F{total_row}")
c = ws[f"A{total_row}"]
c.value     = "합   계"
c.font      = Font(name="맑은 고딕", bold=True, size=11, color="FFFFFF")
c.alignment = align()
c.fill      = fill(C_TITLE)
c.border    = tb()

for col_idx in range(7, 14):
    col_letter = get_column_letter(col_idx)
    val = (14_250_000_000 if col_idx == 13
           else f"=SUM({col_letter}4:{col_letter}{total_row-1})")
    c = ws.cell(row=total_row, column=col_idx, value=val)
    c.font          = Font(name="맑은 고딕", bold=True, size=10, color="FFFFFF")
    c.alignment     = align(h="right")
    c.fill          = fill(C_TITLE)
    c.border        = tb()
    c.number_format = NUM_FMT

c = ws.cell(row=total_row, column=14, value="")
c.fill = fill(C_TITLE); c.border = tb()
ws.row_dimensions[total_row].height = 26

# ── 안내 ────────────────────────────────────────────────────
note_row = total_row + 2
ws.merge_cells(f"A{note_row}:N{note_row}")
c = ws[f"A{note_row}"]
c.value     = ("※ 청구월은 사업비회수금 총괄표의 처리일자 기준.  "
               "1회(기안 11/29→처리 12/04), 4회(기안 3/25→처리 4/06)는 기안월과 처리월이 상이함.")
c.font      = Font(name="맑은 고딕", italic=True, size=9, color="C00000")
c.alignment = align(h="left")

# ── 인쇄 설정 ────────────────────────────────────────────────
ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
ws.page_setup.fitToPage   = True
ws.page_setup.fitToWidth  = 1
ws.freeze_panes           = "G4"

output_path = "/home/user/NRB/의왕초평_도급기성청구현황.xlsx"
wb.save(output_path)
print(f"저장 완료: {output_path}")
