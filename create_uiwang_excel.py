import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "도급기성청구현황"

# ── 스타일 ──────────────────────────────────────────────────
thin  = Side(style="thin")
med   = Side(style="medium")

def tb(l=thin, r=thin, t=thin, b=thin):
    return Border(left=l, right=r, top=t, bottom=b)

def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def font(bold=False, color="000000", size=10, italic=False):
    return Font(name="맑은 고딕", bold=bold, color=color, size=size, italic=italic)

def align(h="center", v="center", wrap=True):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

C_TITLE  = "1F4E79"   # 진한 파랑
C_HEAD1  = "2E75B6"   # 중간 파랑 (대분류 헤더)
C_HEAD2  = "BDD7EE"   # 연한 파랑 (소분류 헤더)
C_ALT    = "EBF3FA"   # 교대 행
C_TOTAL  = "D6E4F0"   # 합계 행
C_NOTE   = "FFF2CC"   # 비고 (미확인) 행
C_WHITE  = "FFFFFF"

# ── 열 너비 ─────────────────────────────────────────────────
col_widths = {
    "A": 7,   # 회차
    "B": 20,  # 품의번호
    "C": 14,  # 기안일자
    "D": 13,  # 청구월
    "E": 17,  # 금회금액(합계)
    "F": 17,  # 극동건설
    "G": 16,  # HJ중공업
    "H": 16,  # 서한
    "I": 16,  # 대저건설
    "J": 16,  # 엔알비
    "K": 17,  # 누계금액
    "L": 14,  # 비고
}
for col, w in col_widths.items():
    ws.column_dimensions[col].width = w

NUM_FMT = '#,##0'

# ── 제목 ────────────────────────────────────────────────────
ws.merge_cells("A1:L1")
c = ws["A1"]
c.value      = "의왕초평 A-4BL 도급 기성청구 현황"
c.font       = Font(name="맑은 고딕", bold=True, size=15, color="FFFFFF")
c.alignment  = align()
c.fill       = fill(C_TITLE)
c.border     = tb()
ws.row_dimensions[1].height = 38

# ── 계약 정보 ────────────────────────────────────────────────
ws.merge_cells("A2:D2")
ws["A2"].value = "계약금액 (도급기성청구 합계)"
ws["A2"].font  = font(bold=True, size=10)
ws["A2"].alignment = align(h="left")
ws["A2"].fill  = fill(C_HEAD2)
ws["A2"].border = tb()

ws.merge_cells("E2:F2")
ws["E2"].value       = 86_000_000_000
ws["E2"].font        = font(bold=True, size=10)
ws["E2"].alignment   = align()
ws["E2"].fill        = fill(C_HEAD2)
ws["E2"].border      = tb()
ws["E2"].number_format = NUM_FMT

ws.merge_cells("G2:J2")
ws["G2"].value     = "※ 추가정산 포함 총계약 90,000,000,000원"
ws["G2"].font      = font(italic=True, size=9, color="595959")
ws["G2"].alignment = align(h="left")
ws["G2"].fill      = fill(C_HEAD2)
ws["G2"].border    = tb()

ws.merge_cells("K2:L2")
ws["K2"].fill   = fill(C_HEAD2)
ws["K2"].border = tb()
ws.row_dimensions[2].height = 22

# ── 헤더 (3행) ───────────────────────────────────────────────
headers = [
    ("회차",         "A3"),
    ("품의번호",      "B3"),
    ("기안일자",      "C3"),
    ("청구월",        "D3"),
    ("금회금액\n(합계)","E3"),
    ("극동건설(주)\n51%","F3"),
    ("(주)HJ중공업\n15%","G3"),
    ("(주)서한\n14%","H3"),
    ("(주)대저건설\n10%","I3"),
    ("(주)엔알비\n10%","J3"),
    ("누계금액",      "K3"),
    ("비고",          "L3"),
]
for label, addr in headers:
    c = ws[addr]
    c.value     = label
    c.font      = Font(name="맑은 고딕", bold=True, size=10, color="FFFFFF")
    c.alignment = align()
    c.fill      = fill(C_HEAD1)
    c.border    = tb()
ws.row_dimensions[3].height = 36

# ── 데이터 ──────────────────────────────────────────────────
data = [
    # (회차, 품의번호, 기안일자, 청구월, 합계, 극동, HJ, 서한, 대저, 엔알비, 누계, 비고)
    (
        "1회", "(주)엔알비-2025-17478", "2025-11-29", "2025년 11월",
        5_076_000_000,
        2_588_760_000, 761_400_000, 710_640_000, 507_600_000, 507_600_000,
        5_076_000_000, ""
    ),
    (
        "2회", "(주)엔알비-2026-1859", "2026-02-02", "2026년 2월",
        624_000_000,
        318_240_000, 93_600_000, 87_360_000, 62_400_000, 62_400_000,
        5_700_000_000, ""
    ),
    (
        "3회", "(주)엔알비-2026-4026", "2026-03-05", "2026년 3월 (초)",
        1_640_000_000,
        836_400_000, 246_000_000, 229_600_000, 164_000_000, 164_000_000,
        7_340_000_000, ""
    ),
    (
        "4회", "(주)엔알비-2026-5251", "2026-03-25", "2026년 3월 (말)",
        2_930_000_000,
        1_494_300_000, 439_500_000, 410_200_000, 293_000_000, 293_000_000,
        10_270_000_000, ""
    ),
    (
        "5회", "(주)엔알비-2026-7755", "미확인", "미확인",
        3_980_000_000,
        2_029_800_000, 597_000_000, 557_200_000, 398_000_000, 398_000_000,
        14_250_000_000, "PDF 1p(기안서 헤더) 누락 – 날짜 확인 필요"
    ),
]

for i, row_data in enumerate(data):
    row = 4 + i
    bg = C_NOTE if row_data[2] == "미확인" else (C_ALT if i % 2 == 0 else C_WHITE)

    for col_idx, val in enumerate(row_data, start=1):
        c = ws.cell(row=row, column=col_idx, value=val)
        c.fill   = fill(bg)
        c.border = tb()

        # 회차: 볼드 중앙
        if col_idx == 1:
            c.font      = font(bold=True, size=11)
            c.alignment = align()
        # 품의번호, 기안일자, 청구월, 비고: 왼쪽 정렬
        elif col_idx in (2, 3, 4, 12):
            c.font      = font(size=9)
            c.alignment = align(h="left")
        # 숫자 열 (합계, 업체별, 누계)
        elif col_idx in range(5, 12):
            c.font          = font(bold=(col_idx in (5, 11)), size=10,
                                   color=("C00000" if col_idx == 5 else "000000"))
            c.alignment     = align(h="right")
            c.number_format = NUM_FMT
        else:
            c.font      = font(size=9)
            c.alignment = align(h="left")

    ws.row_dimensions[row].height = 24

# ── 합계 행 ─────────────────────────────────────────────────
total_row = 4 + len(data)
ws.merge_cells(f"A{total_row}:D{total_row}")
c = ws[f"A{total_row}"]
c.value     = "합   계"
c.font      = Font(name="맑은 고딕", bold=True, size=11, color="FFFFFF")
c.alignment = align()
c.fill      = fill(C_TITLE)
c.border    = tb()

for col_idx in range(5, 12):
    col_letter = get_column_letter(col_idx)
    if col_idx == 11:
        # 누계는 최종 누계값
        c = ws.cell(row=total_row, column=col_idx, value=14_250_000_000)
    else:
        c = ws.cell(
            row=total_row, column=col_idx,
            value=f"=SUM({col_letter}4:{col_letter}{total_row-1})"
        )
    c.font          = Font(name="맑은 고딕", bold=True, size=10, color="FFFFFF")
    c.alignment     = align(h="right")
    c.fill          = fill(C_TITLE)
    c.border        = tb()
    c.number_format = NUM_FMT

# 비고 합계 칸
c = ws.cell(row=total_row, column=12, value="")
c.fill   = fill(C_TITLE)
c.border = tb()
ws.row_dimensions[total_row].height = 26

# ── 안내 ────────────────────────────────────────────────────
note_row = total_row + 2
ws.merge_cells(f"A{note_row}:L{note_row}")
c = ws[f"A{note_row}"]
c.value     = "※ 5회차 기안일자 미확인 (PDF 첫 페이지 누락). 품의번호는 파일명 기준 추정값."
c.font      = Font(name="맑은 고딕", italic=True, size=9, color="C00000")
c.alignment = align(h="left")

# ── 인쇄 설정 ────────────────────────────────────────────────
ws.page_setup.orientation  = ws.ORIENTATION_LANDSCAPE
ws.page_setup.fitToPage    = True
ws.page_setup.fitToWidth   = 1
ws.freeze_panes            = "E4"

output_path = "/home/user/NRB/의왕초평_도급기성청구현황.xlsx"
wb.save(output_path)
print(f"저장 완료: {output_path}")
