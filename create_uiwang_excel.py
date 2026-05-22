import openpyxl
from openpyxl.styles import (
    Font, Alignment, PatternFill, Border, Side, numbers
)
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "도급기성청구현황"

# --- 스타일 정의 ---
header_font = Font(name="맑은 고딕", bold=True, size=11)
body_font   = Font(name="맑은 고딕", size=10)
title_font  = Font(name="맑은 고딕", bold=True, size=14)

center = Alignment(horizontal="center", vertical="center", wrap_text=True)
left   = Alignment(horizontal="left",   vertical="center", wrap_text=True)

thin = Side(style="thin")
medium = Side(style="medium")

def thin_border():
    return Border(left=thin, right=thin, top=thin, bottom=thin)

def medium_border():
    return Border(left=medium, right=medium, top=medium, bottom=medium)

header_fill  = PatternFill("solid", fgColor="1F4E79")  # 진한 파랑
subhdr_fill  = PatternFill("solid", fgColor="BDD7EE")  # 연한 파랑
alt_fill     = PatternFill("solid", fgColor="EBF3FA")  # 더 연한 파랑
white_fill   = PatternFill("solid", fgColor="FFFFFF")

header_font_white = Font(name="맑은 고딕", bold=True, size=11, color="FFFFFF")

# --- 제목 ---
ws.merge_cells("A1:O1")
ws["A1"] = "의왕초평 프로젝트 – 도급 기성청구 현황"
ws["A1"].font = title_font
ws["A1"].alignment = center
ws["A1"].fill = PatternFill("solid", fgColor="1F4E79")
ws["A1"].font = Font(name="맑은 고딕", bold=True, size=14, color="FFFFFF")
ws.row_dimensions[1].height = 36

# --- 컬럼 헤더 (2행) ---
months = [f"{m}월" for m in range(1, 13)]  # 1월~12월

headers = ["회차", "공종/업체명", "계약금액"] + months + ["합계", "비고"]
col_count = len(headers)  # 3 + 12 + 2 = 17

# 열 너비
ws.column_dimensions["A"].width = 8    # 회차
ws.column_dimensions["B"].width = 22   # 공종/업체명
ws.column_dimensions["C"].width = 16   # 계약금액
for i in range(4, 16):                 # D~O (1월~12월)
    ws.column_dimensions[get_column_letter(i)].width = 12
ws.column_dimensions["P"].width = 16   # 합계
ws.column_dimensions["Q"].width = 14   # 비고

# 2행 헤더 작성
for col_idx, hdr in enumerate(headers, start=1):
    cell = ws.cell(row=2, column=col_idx, value=hdr)
    cell.font = header_font_white
    cell.alignment = center
    cell.fill = header_fill
    cell.border = thin_border()

ws.row_dimensions[2].height = 30

# --- 데이터 행: 회차 1~10 ---
ROUNDS = 10
data_start = 3

for r in range(ROUNDS):
    row = data_start + r
    fill = alt_fill if r % 2 == 0 else white_fill

    # 회차
    cell_round = ws.cell(row=row, column=1, value=r + 1)
    cell_round.font = Font(name="맑은 고딕", bold=True, size=10)
    cell_round.alignment = center
    cell_round.fill = fill
    cell_round.border = thin_border()

    # 공종/업체명
    c = ws.cell(row=row, column=2, value="")
    c.font = body_font; c.alignment = left; c.fill = fill; c.border = thin_border()

    # 계약금액
    c = ws.cell(row=row, column=3, value=None)
    c.font = body_font; c.alignment = center; c.fill = fill; c.border = thin_border()
    c.number_format = '#,##0'

    # 월별 청구금액 (D~O)
    for m_col in range(4, 16):
        c = ws.cell(row=row, column=m_col, value=None)
        c.font = body_font; c.alignment = center; c.fill = fill; c.border = thin_border()
        c.number_format = '#,##0'

    # 합계 (SUM D~O)
    sum_range = f"D{row}:O{row}"
    c = ws.cell(row=row, column=16, value=f"=SUM({sum_range})")
    c.font = Font(name="맑은 고딕", bold=True, size=10)
    c.alignment = center; c.fill = fill; c.border = thin_border()
    c.number_format = '#,##0'

    # 비고
    c = ws.cell(row=row, column=17, value="")
    c.font = body_font; c.alignment = left; c.fill = fill; c.border = thin_border()

    ws.row_dimensions[row].height = 22

# --- 합계 행 ---
total_row = data_start + ROUNDS
ws.merge_cells(f"A{total_row}:B{total_row}")
c = ws.cell(row=total_row, column=1, value="합 계")
c.font = Font(name="맑은 고딕", bold=True, size=11, color="FFFFFF")
c.alignment = center; c.fill = PatternFill("solid", fgColor="2E75B6"); c.border = thin_border()

# 계약금액 합계
c = ws.cell(row=total_row, column=3, value=f"=SUM(C{data_start}:C{total_row-1})")
c.font = Font(name="맑은 고딕", bold=True, size=10, color="FFFFFF")
c.alignment = center; c.fill = PatternFill("solid", fgColor="2E75B6"); c.border = thin_border()
c.number_format = '#,##0'

for m_col in range(4, 17):
    col_letter = get_column_letter(m_col)
    c = ws.cell(row=total_row, column=m_col,
                value=f"=SUM({col_letter}{data_start}:{col_letter}{total_row-1})")
    c.font = Font(name="맑은 고딕", bold=True, size=10, color="FFFFFF")
    c.alignment = center; c.fill = PatternFill("solid", fgColor="2E75B6"); c.border = thin_border()
    c.number_format = '#,##0'

# 비고 칸
c = ws.cell(row=total_row, column=17, value="")
c.fill = PatternFill("solid", fgColor="2E75B6"); c.border = thin_border()

ws.row_dimensions[total_row].height = 26

# --- 안내 텍스트 ---
note_row = total_row + 2
ws.cell(row=note_row, column=1,
        value="※ 월별 청구금액 칸에 해당 회차 기성청구가 접수된 월의 금액을 입력하세요.")
ws.cell(row=note_row, column=1).font = Font(name="맑은 고딕", italic=True, size=9, color="595959")

# --- 인쇄 설정 ---
ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
ws.page_setup.fitToPage = True
ws.page_setup.fitToWidth = 1
ws.sheet_view.showGridLines = True
ws.freeze_panes = "C3"  # 회차·업체명 고정

output_path = "/home/user/NRB/의왕초평_도급기성청구현황.xlsx"
wb.save(output_path)
print(f"저장 완료: {output_path}")
