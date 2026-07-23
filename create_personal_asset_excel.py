"""
개인 자산관리 양식 생성 스크립트
회사 '자금수지' 양식(구분/세부항목/월별 Cash In-Out/과부족/잔고, 차입금현황, 고정자금수지,
요약자금수지 구조)을 참고하여 개인용 자산관리 워크북을 생성한다.
"""
import datetime

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

FONT_NAME = "맑은 고딕"

# ── 공통 스타일 ──────────────────────────────────────────
thin = Side(style="thin", color="BFBFBF")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
thick_top = Border(left=thin, right=thin, top=Side(style="thin", color="000000"), bottom=thin)

header_fill = PatternFill(fill_type="solid", fgColor="1F3864")      # 진한 남색 (헤더)
category_fill = PatternFill(fill_type="solid", fgColor="D9E1F2")    # 연한 남색 (카테고리 라벨)
total_fill = PatternFill(fill_type="solid", fgColor="D6E4F0")       # 연한 파란색 (소계/합계)
grand_fill = PatternFill(fill_type="solid", fgColor="2E74B5")       # 중간 파란색 (총계)
example_fill = PatternFill(fill_type="solid", fgColor="FFF9E6")     # 연노랑 (예시행)

white_bold = Font(name=FONT_NAME, bold=True, color="FFFFFF", size=10)
title_font = Font(name=FONT_NAME, bold=True, size=13)
note_font = Font(name=FONT_NAME, italic=True, size=9, color="0070C0")
label_font = Font(name=FONT_NAME, bold=True, size=10)
black_font = Font(name=FONT_NAME, size=10)
input_font = Font(name=FONT_NAME, size=10, color="0000FF")          # 입력 셀(예시 포함)
link_font = Font(name=FONT_NAME, size=10, color="008000")           # 타 시트 연동
total_font = Font(name=FONT_NAME, bold=True, size=10, color="1F3864")
grand_font = Font(name=FONT_NAME, bold=True, size=12, color="FFFFFF")

center = Alignment(horizontal="center", vertical="center", wrap_text=True)
left = Alignment(horizontal="left", vertical="center", wrap_text=True)
right = Alignment(horizontal="right", vertical="center")

MONEY_FMT = '#,##0;[Red](#,##0)'
DATE_FMT = "yyyy-mm-dd"
PCT_FMT = "0.00%"

MONTHS = [f"{m}월" for m in range(1, 13)]


def style_header_row(ws, row, first_col, last_col):
    for c in range(first_col, last_col + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = white_bold
        cell.fill = header_fill
        cell.alignment = center
        cell.border = border


def style_range_border(ws, r1, c1, r2, c2):
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            ws.cell(row=r, column=c).border = border


def title_row(ws, text, last_col_letter, row=1):
    ws.merge_cells(f"A{row}:{last_col_letter}{row}")
    c = ws[f"A{row}"]
    c.value = text
    c.font = title_font
    c.alignment = left
    ws.row_dimensions[row].height = 24


def note_row(ws, text, last_col_letter, row=2):
    ws.merge_cells(f"A{row}:{last_col_letter}{row}")
    c = ws[f"A{row}"]
    c.value = text
    c.font = note_font
    c.alignment = left
    ws.row_dimensions[row].height = 16


wb = openpyxl.Workbook()

# ══════════════════════════════════════════════════════════════════
# 1. 요약대시보드
# ══════════════════════════════════════════════════════════════════
ws = wb.active
ws.title = "요약대시보드"
last_col = "N"
title_row(ws, "■ 개인 자산관리 요약 대시보드 (2026년)", last_col)
note_row(ws, "※ 파란색 셀 = 직접 입력 · 검정색 = 자동 계산(수식) · 초록색 = 다른 시트와 연동된 셀", last_col)

ws["A4"].value = "작성기준일"
ws["A4"].font = label_font
ws.merge_cells("B4:C4")
ws["B4"].value = datetime.date(2026, 7, 23)
ws["B4"].font = input_font
ws["B4"].number_format = DATE_FMT
ws["B4"].alignment = center

for col in ["A", "B", "C"]:
    ws.column_dimensions[col].width = 14
for col in "DEFGHIJKLM":
    ws.column_dimensions[col].width = 11
ws.column_dimensions["N"].width = 13

# 순자산 현황
ws.merge_cells("A6:N6")
ws["A6"].value = "순자산 현황"
ws["A6"].font = white_bold
ws["A6"].fill = header_fill
ws["A6"].alignment = center
ws.row_dimensions[6].height = 18

rows_summary = [
    (7, "총자산"),
    (8, "총부채"),
]
for r, label in rows_summary:
    ws.merge_cells(f"A{r}:B{r}")
    ws[f"A{r}"].value = label
    ws[f"A{r}"].font = label_font
    ws[f"A{r}"].alignment = center
    ws.merge_cells(f"C{r}:D{r}")
    ws[f"C{r}"].font = link_font
    ws[f"C{r}"].number_format = MONEY_FMT
    ws[f"C{r}"].alignment = right
    style_range_border(ws, r, 1, r, 4)
dashboard_ws = ws  # 자산현황/부채현황 총계 행 확정 후 C7/C8 수식 채움

ws.merge_cells("A9:B9")
ws["A9"].value = "순자산 (자산-부채)"
ws["A9"].font = total_font
ws["A9"].fill = grand_fill
ws["A9"].alignment = center
ws.merge_cells("C9:D9")
ws["C9"].value = "=C7-C8"
ws["C9"].font = grand_font
ws["C9"].fill = grand_fill
ws["C9"].number_format = MONEY_FMT
ws["C9"].alignment = right
style_range_border(ws, 9, 1, 9, 4)
ws.row_dimensions[9].height = 20

# 월별 현금흐름 요약
ws.merge_cells("A11:N11")
ws["A11"].value = "월별 현금흐름 요약 (단위: 원, '월별자금수지' 시트 연동)"
ws["A11"].font = white_bold
ws["A11"].fill = header_fill
ws["A11"].alignment = center
ws.row_dimensions[11].height = 18

ws["A12"].value = "구분"
for i, m in enumerate(MONTHS):
    col = get_column_letter(2 + i)
    ws[f"{col}12"].value = m
ws["N12"].value = "합계"
style_header_row(ws, 12, 1, 14)

cf_rows = [
    (13, "Cash In", 12, "C"),
    (14, "Cash Out", 22, "C"),
    (15, "과부족", 24, "C"),
    (16, "월말잔고(누적)", 25, "C"),
]
for r, label, src_row, src_total_col in cf_rows:
    ws[f"A{r}"].value = label
    ws[f"A{r}"].font = label_font
    ws[f"A{r}"].alignment = center
    for i in range(12):
        col = get_column_letter(2 + i)          # B..M
        src_col = get_column_letter(4 + i)       # D..O
        cell = ws[f"{col}{r}"]
        cell.value = f"='월별자금수지'!{src_col}{src_row}"
        cell.font = link_font
        cell.number_format = MONEY_FMT
    ws[f"N{r}"].value = f"='월별자금수지'!{src_total_col}{src_row}"
    ws[f"N{r}"].font = link_font
    ws[f"N{r}"].number_format = MONEY_FMT
style_range_border(ws, 12, 1, 16, 14)

ws.freeze_panes = "B13"

# ══════════════════════════════════════════════════════════════════
# 2. 월별자금수지 (메인)
# ══════════════════════════════════════════════════════════════════
ws = wb.create_sheet("월별자금수지")
last_col = "P"
title_row(ws, "■ 개인 월별 자금수지 (2026년, 단위: 원)", last_col)
note_row(ws, "※ 파란색 셀(월별 금액, 기초잔액)에 직접 입력하세요. 합계/과부족/잔고는 자동 계산됩니다.", last_col)

ws.column_dimensions["A"].width = 10
ws.column_dimensions["B"].width = 20
ws.column_dimensions["C"].width = 13
for i in range(12):
    ws.column_dimensions[get_column_letter(4 + i)].width = 10.5
ws.column_dimensions["P"].width = 14

# 헤더
ws["A4"].value = "구분"
ws["B4"].value = "세부항목"
ws["C4"].value = "합계"
for i, m in enumerate(MONTHS):
    ws[f"{get_column_letter(4 + i)}4"].value = m
ws["P4"].value = "비고"
style_header_row(ws, 4, 1, 16)
ws.row_dimensions[4].height = 18

# 기초잔액
ws.merge_cells("A5:B5")
ws["A5"].value = "기초잔액(전년이월)"
ws["A5"].font = label_font
ws["A5"].fill = category_fill
ws["A5"].alignment = center
ws["C5"].value = "=D5"
ws["C5"].font = black_font
ws["C5"].number_format = MONEY_FMT
ws["D5"].value = 0
ws["D5"].font = input_font
ws["D5"].number_format = MONEY_FMT
ws["P5"].value = "직접 입력"
ws["P5"].font = note_font
style_range_border(ws, 5, 1, 5, 16)

# Cash In
cash_in_items = ["근로소득(급여)", "사업소득", "금융소득(이자/배당)", "기타수입", ""]
cash_in_start = 7
for i, item in enumerate(cash_in_items):
    r = cash_in_start + i
    ws[f"B{r}"].value = item
    ws[f"B{r}"].font = black_font
    ws[f"B{r}"].alignment = left
    ws[f"C{r}"].value = f"=SUM(D{r}:O{r})"
    ws[f"C{r}"].font = black_font
    ws[f"C{r}"].number_format = MONEY_FMT
    for c in range(4, 16):
        cell = ws.cell(row=r, column=c)
        cell.number_format = MONEY_FMT
        cell.font = input_font
    style_range_border(ws, r, 1, r, 16)
cash_in_end = cash_in_start + len(cash_in_items) - 1
ws.merge_cells(f"A{cash_in_start}:A{cash_in_end}")
ws[f"A{cash_in_start}"].value = "Cash In\n(수입)"
ws[f"A{cash_in_start}"].font = label_font
ws[f"A{cash_in_start}"].fill = category_fill
ws[f"A{cash_in_start}"].alignment = center

cash_in_total_row = cash_in_end + 1
ws.merge_cells(f"A{cash_in_total_row}:B{cash_in_total_row}")
ws[f"A{cash_in_total_row}"].value = "Cash In 합계"
ws[f"A{cash_in_total_row}"].font = total_font
ws[f"A{cash_in_total_row}"].fill = total_fill
ws[f"A{cash_in_total_row}"].alignment = center
for c in range(3, 16):
    col = get_column_letter(c)
    cell = ws[f"{col}{cash_in_total_row}"]
    cell.value = f"=SUM({col}{cash_in_start}:{col}{cash_in_end})"
    cell.font = total_font
    cell.fill = total_fill
    cell.number_format = MONEY_FMT
style_range_border(ws, cash_in_total_row, 1, cash_in_total_row, 16)

# Cash Out
cash_out_items = [
    "주거비(월세/관리비)", "식비", "교통/통신비", "보험료",
    "대출원리금상환", "저축/투자", "세금/공과금", "기타지출",
]
cash_out_start = cash_in_total_row + 2
for i, item in enumerate(cash_out_items):
    r = cash_out_start + i
    ws[f"B{r}"].value = item
    ws[f"B{r}"].font = black_font
    ws[f"B{r}"].alignment = left
    ws[f"C{r}"].value = f"=SUM(D{r}:O{r})"
    ws[f"C{r}"].font = black_font
    ws[f"C{r}"].number_format = MONEY_FMT
    for c in range(4, 16):
        cell = ws.cell(row=r, column=c)
        cell.number_format = MONEY_FMT
        cell.font = input_font
    style_range_border(ws, r, 1, r, 16)
cash_out_end = cash_out_start + len(cash_out_items) - 1
ws.merge_cells(f"A{cash_out_start}:A{cash_out_end}")
ws[f"A{cash_out_start}"].value = "Cash Out\n(지출)"
ws[f"A{cash_out_start}"].font = label_font
ws[f"A{cash_out_start}"].fill = category_fill
ws[f"A{cash_out_start}"].alignment = center

cash_out_total_row = cash_out_end + 1
ws.merge_cells(f"A{cash_out_total_row}:B{cash_out_total_row}")
ws[f"A{cash_out_total_row}"].value = "Cash Out 합계"
ws[f"A{cash_out_total_row}"].font = total_font
ws[f"A{cash_out_total_row}"].fill = total_fill
ws[f"A{cash_out_total_row}"].alignment = center
for c in range(3, 16):
    col = get_column_letter(c)
    cell = ws[f"{col}{cash_out_total_row}"]
    cell.value = f"=SUM({col}{cash_out_start}:{col}{cash_out_end})"
    cell.font = total_font
    cell.fill = total_fill
    cell.number_format = MONEY_FMT
style_range_border(ws, cash_out_total_row, 1, cash_out_total_row, 16)

# 과부족
surplus_row = cash_out_total_row + 2
ws.merge_cells(f"A{surplus_row}:B{surplus_row}")
ws[f"A{surplus_row}"].value = "과부족 (Cash In - Cash Out)"
ws[f"A{surplus_row}"].font = total_font
ws[f"A{surplus_row}"].fill = category_fill
ws[f"A{surplus_row}"].alignment = center
for c in range(3, 16):
    col = get_column_letter(c)
    cell = ws[f"{col}{surplus_row}"]
    cell.value = f"={col}{cash_in_total_row}-{col}{cash_out_total_row}"
    cell.font = total_font
    cell.number_format = MONEY_FMT
style_range_border(ws, surplus_row, 1, surplus_row, 16)

# 월말잔고(누적)
balance_row = surplus_row + 1
ws.merge_cells(f"A{balance_row}:B{balance_row}")
ws[f"A{balance_row}"].value = "월말잔고(누적)"
ws[f"A{balance_row}"].font = grand_font
ws[f"A{balance_row}"].fill = grand_fill
ws[f"A{balance_row}"].alignment = center
ws[f"C{balance_row}"].value = f"=O{balance_row}"
ws[f"C{balance_row}"].font = grand_font
ws[f"C{balance_row}"].fill = grand_fill
ws[f"C{balance_row}"].number_format = MONEY_FMT
for i in range(12):
    col = get_column_letter(4 + i)
    cell = ws[f"{col}{balance_row}"]
    if i == 0:
        cell.value = f"=D5+D{surplus_row}"
    else:
        prev_col = get_column_letter(4 + i - 1)
        cell.value = f"={prev_col}{balance_row}+{col}{surplus_row}"
    cell.font = grand_font
    cell.fill = grand_fill
    cell.number_format = MONEY_FMT
ws[f"P{balance_row}"].fill = grand_fill
style_range_border(ws, balance_row, 1, balance_row, 16)

ws.freeze_panes = "D5"

# ══════════════════════════════════════════════════════════════════
# 3. 고정지출관리
# ══════════════════════════════════════════════════════════════════
ws = wb.create_sheet("고정지출관리")
last_col = "F"
title_row(ws, "■ 고정 수입/지출 관리 (연간계획, 단위: 원)", last_col)
note_row(ws, "※ 파란색 셀에 매월 고정적으로 발생하는 수입/지출을 입력하세요. 연간합계 = 월금액 × 연횟수", last_col)

ws.column_dimensions["A"].width = 12
ws.column_dimensions["B"].width = 22
ws.column_dimensions["C"].width = 14
ws.column_dimensions["D"].width = 10
ws.column_dimensions["E"].width = 15
ws.column_dimensions["F"].width = 20

ws["A4"].value = "구분"
ws["B4"].value = "항목"
ws["C4"].value = "월금액(원)"
ws["D4"].value = "연횟수"
ws["E4"].value = "연간합계(원)"
ws["F4"].value = "비고"
style_header_row(ws, 4, 1, 6)
ws.row_dimensions[4].height = 18

income_items = [("급여", 3000000, 12, "예시"), ("부수입/용돈", None, 12, None), ("", None, 12, None)]
income_start = 5
for i, (item, amt, cnt, note) in enumerate(income_items):
    r = income_start + i
    is_example = i == 0
    font = input_font if is_example else black_font
    ws[f"B{r}"].value = item
    ws[f"B{r}"].font = font
    ws[f"B{r}"].alignment = left
    ws[f"C{r}"].value = amt
    ws[f"C{r}"].font = input_font
    ws[f"C{r}"].number_format = MONEY_FMT
    ws[f"D{r}"].value = cnt
    ws[f"D{r}"].font = input_font
    ws[f"D{r}"].alignment = center
    ws[f"E{r}"].value = f"=C{r}*D{r}"
    ws[f"E{r}"].font = black_font
    ws[f"E{r}"].number_format = MONEY_FMT
    if note:
        ws[f"F{r}"].value = note
        ws[f"F{r}"].font = note_font
    if is_example:
        for col in "ABCDEF":
            ws[f"{col}{r}"].fill = example_fill
    style_range_border(ws, r, 1, r, 6)
income_end = income_start + len(income_items) - 1
ws.merge_cells(f"A{income_start}:A{income_end}")
ws[f"A{income_start}"].value = "고정수입"
ws[f"A{income_start}"].font = label_font
ws[f"A{income_start}"].fill = category_fill
ws[f"A{income_start}"].alignment = center

income_total_row = income_end + 1
ws.merge_cells(f"A{income_total_row}:D{income_total_row}")
ws[f"A{income_total_row}"].value = "고정수입 소계"
ws[f"A{income_total_row}"].font = total_font
ws[f"A{income_total_row}"].fill = total_fill
ws[f"A{income_total_row}"].alignment = center
ws[f"E{income_total_row}"].value = f"=SUM(E{income_start}:E{income_end})"
ws[f"E{income_total_row}"].font = total_font
ws[f"E{income_total_row}"].fill = total_fill
ws[f"E{income_total_row}"].number_format = MONEY_FMT
style_range_border(ws, income_total_row, 1, income_total_row, 6)

expense_items = [
    ("월세/관리비", 800000, 12, "예시"),
    ("통신비", None, 12, None),
    ("보험료", None, 12, None),
    ("구독료(OTT 등)", None, 12, None),
    ("대출상환액", None, 12, None),
    ("적금/투자", None, 12, None),
    ("기타", None, 12, None),
]
expense_start = income_total_row + 2
for i, (item, amt, cnt, note) in enumerate(expense_items):
    r = expense_start + i
    is_example = i == 0
    ws[f"B{r}"].value = item
    ws[f"B{r}"].font = black_font
    ws[f"B{r}"].alignment = left
    ws[f"C{r}"].value = amt
    ws[f"C{r}"].font = input_font
    ws[f"C{r}"].number_format = MONEY_FMT
    ws[f"D{r}"].value = cnt
    ws[f"D{r}"].font = input_font
    ws[f"D{r}"].alignment = center
    ws[f"E{r}"].value = f"=C{r}*D{r}"
    ws[f"E{r}"].font = black_font
    ws[f"E{r}"].number_format = MONEY_FMT
    if note:
        ws[f"F{r}"].value = note
        ws[f"F{r}"].font = note_font
    if is_example:
        for col in "ABCDEF":
            ws[f"{col}{r}"].fill = example_fill
    style_range_border(ws, r, 1, r, 6)
expense_end = expense_start + len(expense_items) - 1
ws.merge_cells(f"A{expense_start}:A{expense_end}")
ws[f"A{expense_start}"].value = "고정지출"
ws[f"A{expense_start}"].font = label_font
ws[f"A{expense_start}"].fill = category_fill
ws[f"A{expense_start}"].alignment = center

expense_total_row = expense_end + 1
ws.merge_cells(f"A{expense_total_row}:D{expense_total_row}")
ws[f"A{expense_total_row}"].value = "고정지출 소계"
ws[f"A{expense_total_row}"].font = total_font
ws[f"A{expense_total_row}"].fill = total_fill
ws[f"A{expense_total_row}"].alignment = center
ws[f"E{expense_total_row}"].value = f"=SUM(E{expense_start}:E{expense_end})"
ws[f"E{expense_total_row}"].font = total_font
ws[f"E{expense_total_row}"].fill = total_fill
ws[f"E{expense_total_row}"].number_format = MONEY_FMT
style_range_border(ws, expense_total_row, 1, expense_total_row, 6)

net_row = expense_total_row + 2
ws.merge_cells(f"A{net_row}:D{net_row}")
ws[f"A{net_row}"].value = "연간 과부족 (고정수입 - 고정지출)"
ws[f"A{net_row}"].font = grand_font
ws[f"A{net_row}"].fill = grand_fill
ws[f"A{net_row}"].alignment = center
ws[f"E{net_row}"].value = f"=E{income_total_row}-E{expense_total_row}"
ws[f"E{net_row}"].font = grand_font
ws[f"E{net_row}"].fill = grand_fill
ws[f"E{net_row}"].number_format = MONEY_FMT
style_range_border(ws, net_row, 1, net_row, 6)

# ══════════════════════════════════════════════════════════════════
# 4. 자산현황
# ══════════════════════════════════════════════════════════════════
ws = wb.create_sheet("자산현황")
last_col = "F"
title_row(ws, "■ 개인 자산현황 (2026년 기준, 단위: 원)", last_col)
note_row(ws, "※ 파란색 예시행을 참고하여 본인의 자산을 입력하세요. 행이 부족하면 소계 행 위에 행을 추가하세요.", last_col)

ws.column_dimensions["A"].width = 14
ws.column_dimensions["B"].width = 22
ws.column_dimensions["C"].width = 18
ws.column_dimensions["D"].width = 16
ws.column_dimensions["E"].width = 13
ws.column_dimensions["F"].width = 16

ws["A4"].value = "구분"
ws["B4"].value = "항목"
ws["C4"].value = "금융기관/소재지"
ws["D4"].value = "평가금액(원)"
ws["E4"].value = "취득일"
ws["F4"].value = "비고"
style_header_row(ws, 4, 1, 6)
ws.row_dimensions[4].height = 18

asset_categories = [
    ("현금성자산\n(예금/적금)", [
        ("급여통장", "OO은행", 5000000, datetime.date(2020, 3, 2), "예시"),
        ("", "", None, None, None),
        ("", "", None, None, None),
        ("", "", None, None, None),
    ]),
    ("투자자산\n(주식/펀드/코인 등)", [
        ("", "", None, None, None),
        ("", "", None, None, None),
        ("", "", None, None, None),
        ("", "", None, None, None),
    ]),
    ("부동산", [
        ("", "", None, None, None),
        ("", "", None, None, None),
        ("", "", None, None, None),
    ]),
    ("차량/기타동산", [
        ("", "", None, None, None),
        ("", "", None, None, None),
        ("", "", None, None, None),
    ]),
    ("기타자산", [
        ("", "", None, None, None),
        ("", "", None, None, None),
        ("", "", None, None, None),
    ]),
]

subtotal_rows = []
r = 5
for cat_label, items in asset_categories:
    cat_start = r
    for i, (item, org, amt, adate, note) in enumerate(items):
        is_example = note == "예시"
        ws[f"B{r}"].value = item
        ws[f"B{r}"].font = black_font
        ws[f"B{r}"].alignment = left
        ws[f"C{r}"].value = org
        ws[f"C{r}"].font = black_font
        ws[f"C{r}"].alignment = left
        ws[f"D{r}"].value = amt
        ws[f"D{r}"].font = input_font
        ws[f"D{r}"].number_format = MONEY_FMT
        if adate:
            ws[f"E{r}"].value = adate
            ws[f"E{r}"].number_format = DATE_FMT
        ws[f"E{r}"].font = input_font
        ws[f"E{r}"].alignment = center
        if note:
            ws[f"F{r}"].value = note
            ws[f"F{r}"].font = note_font
        if is_example:
            for col in "ABCDEF":
                ws[f"{col}{r}"].fill = example_fill
                if col in ("B", "C"):
                    ws[f"{col}{r}"].font = input_font
        style_range_border(ws, r, 1, r, 6)
        r += 1
    cat_end = r - 1
    ws.merge_cells(f"A{cat_start}:A{cat_end}")
    ws[f"A{cat_start}"].value = cat_label
    ws[f"A{cat_start}"].font = label_font
    ws[f"A{cat_start}"].fill = category_fill
    ws[f"A{cat_start}"].alignment = center

    ws.merge_cells(f"B{r}:C{r}")
    ws[f"B{r}"].value = f"{cat_label.splitlines()[0]} 소계"
    ws[f"B{r}"].font = total_font
    ws[f"B{r}"].fill = total_fill
    ws[f"B{r}"].alignment = center
    ws[f"A{r}"].fill = total_fill
    ws[f"D{r}"].value = f"=SUM(D{cat_start}:D{cat_end})"
    ws[f"D{r}"].font = total_font
    ws[f"D{r}"].fill = total_fill
    ws[f"D{r}"].number_format = MONEY_FMT
    for col in ("E", "F"):
        ws[f"{col}{r}"].fill = total_fill
    style_range_border(ws, r, 1, r, 6)
    subtotal_rows.append(r)
    r += 2  # blank spacer row between categories

grand_row = r
ws.merge_cells(f"A{grand_row}:C{grand_row}")
ws[f"A{grand_row}"].value = "총자산 합계"
ws[f"A{grand_row}"].font = grand_font
ws[f"A{grand_row}"].fill = grand_fill
ws[f"A{grand_row}"].alignment = center
formula = "=" + "+".join(f"D{sr}" for sr in subtotal_rows)
ws[f"D{grand_row}"].value = formula
ws[f"D{grand_row}"].font = grand_font
ws[f"D{grand_row}"].fill = grand_fill
ws[f"D{grand_row}"].number_format = MONEY_FMT
for col in ("E", "F"):
    ws[f"{col}{grand_row}"].fill = grand_fill
style_range_border(ws, grand_row, 1, grand_row, 6)
asset_grand_row = grand_row

# ══════════════════════════════════════════════════════════════════
# 5. 부채현황
# ══════════════════════════════════════════════════════════════════
ws = wb.create_sheet("부채현황")
last_col = "K"
title_row(ws, "■ 개인 부채(대출)현황 (2026년 기준, 단위: 원)", last_col)
note_row(ws, "※ 파란색 예시행을 참고하여 본인의 대출을 입력하세요.", last_col)

widths = {"A": 12, "B": 12, "C": 16, "D": 11, "E": 11, "F": 10, "G": 8, "H": 14, "I": 14, "J": 13, "K": 16}
for col, w in widths.items():
    ws.column_dimensions[col].width = w

headers = ["구분", "금융기관", "상품명", "대출일", "만기일", "상환일", "금리", "최초대출금(원)", "현재잔액(원)", "월상환액(원)", "비고"]
for i, h in enumerate(headers):
    ws.cell(row=4, column=1 + i).value = h
style_header_row(ws, 4, 1, 11)
ws.row_dimensions[4].height = 18

example = ("신용대출", "OO은행", "개인신용대출", datetime.date(2024, 1, 15),
           datetime.date(2029, 1, 15), "매월 25일", 0.045, 30000000, 22000000, 550000, "예시")
blank_rows = 6
data_start = 5
for i in range(1 + blank_rows):
    r = data_start + i
    if i == 0:
        vals = example
        is_example = True
    else:
        vals = ("", "", "", None, None, "", None, None, None, None, None)
        is_example = False
    (gubun, org, prod, sdate, edate, pdate, rate, orig, cur, pay, note) = vals
    ws[f"A{r}"].value = gubun
    ws[f"B{r}"].value = org
    ws[f"C{r}"].value = prod
    ws[f"D{r}"].value = sdate
    ws[f"D{r}"].number_format = DATE_FMT
    ws[f"E{r}"].value = edate
    ws[f"E{r}"].number_format = DATE_FMT
    ws[f"F{r}"].value = pdate
    ws[f"G{r}"].value = rate
    ws[f"G{r}"].number_format = PCT_FMT
    ws[f"H{r}"].value = orig
    ws[f"H{r}"].number_format = MONEY_FMT
    ws[f"I{r}"].value = cur
    ws[f"I{r}"].number_format = MONEY_FMT
    ws[f"J{r}"].value = pay
    ws[f"J{r}"].number_format = MONEY_FMT
    ws[f"K{r}"].value = note
    for col in "ABCDEFGHIJK":
        cell = ws[f"{col}{r}"]
        cell.font = input_font if col not in ("K",) else note_font
        if col in ("A", "B", "C", "F"):
            cell.alignment = left
        elif col == "K":
            pass
        else:
            cell.alignment = center
        if is_example:
            cell.fill = example_fill
    style_range_border(ws, r, 1, r, 11)
data_end = data_start + blank_rows

total_row = data_end + 1
ws.merge_cells(f"A{total_row}:G{total_row}")
ws[f"A{total_row}"].value = "합계"
ws[f"A{total_row}"].font = grand_font
ws[f"A{total_row}"].fill = grand_fill
ws[f"A{total_row}"].alignment = center
for col in ("H", "I", "J"):
    cell = ws[f"{col}{total_row}"]
    cell.value = f"=SUM({col}{data_start}:{col}{data_end})"
    cell.font = grand_font
    cell.fill = grand_fill
    cell.number_format = MONEY_FMT
ws[f"K{total_row}"].fill = grand_fill
style_range_border(ws, total_row, 1, total_row, 11)
liability_total_row = total_row

# 대시보드 총자산/총부채 셀에 실제 행번호로 수식 연결
dashboard_ws["C7"].value = f"='자산현황'!D{asset_grand_row}"
dashboard_ws["C8"].value = f"='부채현황'!I{liability_total_row}"

# ══════════════════════════════════════════════════════════════════
# 6. 일별가계부
# ══════════════════════════════════════════════════════════════════
ws = wb.create_sheet("일별가계부")
last_col = "H"
title_row(ws, "■ 일별 가계부 (2026년, 단위: 원)", last_col)
note_row(ws, "※ 거래가 발생할 때마다 한 줄씩 입력하세요. G열(잔액)은 자동으로 계산됩니다.", last_col)

widths = {"A": 12, "B": 9, "C": 12, "D": 22, "E": 12, "F": 12, "G": 13, "H": 14}
for col, w in widths.items():
    ws.column_dimensions[col].width = w

ws["A4"].value = "기초잔액"
ws["A4"].font = label_font
ws["A4"].fill = category_fill
ws.merge_cells("B4:F4")
ws["B4"].fill = category_fill
ws["G4"].value = 0
ws["G4"].font = input_font
ws["G4"].number_format = MONEY_FMT
ws["G4"].fill = category_fill
ws["H4"].value = "직접 입력"
ws["H4"].font = note_font
ws["H4"].fill = category_fill
style_range_border(ws, 4, 1, 4, 8)

headers = ["날짜", "구분", "카테고리", "내용", "수입금액(원)", "지출금액(원)", "잔액(원)", "비고"]
for i, h in enumerate(headers):
    ws.cell(row=5, column=1 + i).value = h
style_header_row(ws, 5, 1, 8)
ws.row_dimensions[5].height = 18

n_rows = 30
data_start = 6
example = (datetime.date(2026, 1, 5), "지출", "식비", "점심 식사", None, 12000, "예시")
for i in range(n_rows):
    r = data_start + i
    is_example = i == 0
    if is_example:
        adate, gubun, cat, desc, inc, exp, note = example
    else:
        adate, gubun, cat, desc, inc, exp, note = (None, "", "", "", None, None, None)
    ws[f"A{r}"].value = adate
    ws[f"A{r}"].number_format = DATE_FMT
    ws[f"B{r}"].value = gubun
    ws[f"C{r}"].value = cat
    ws[f"D{r}"].value = desc
    ws[f"E{r}"].value = inc
    ws[f"E{r}"].number_format = MONEY_FMT
    ws[f"F{r}"].value = exp
    ws[f"F{r}"].number_format = MONEY_FMT
    if r == data_start:
        ws[f"G{r}"].value = f"=G4+E{r}-F{r}"
    else:
        ws[f"G{r}"].value = f"=G{r-1}+E{r}-F{r}"
    ws[f"G{r}"].number_format = MONEY_FMT
    if note:
        ws[f"H{r}"].value = note
        ws[f"H{r}"].font = note_font
    for col in "ABCDEFG":
        cell = ws[f"{col}{r}"]
        cell.font = input_font
        if col in ("A", "B", "C"):
            cell.alignment = center
        elif col == "D":
            cell.alignment = left
    if is_example:
        for col in "ABCDEFGH":
            ws[f"{col}{r}"].fill = example_fill
    style_range_border(ws, r, 1, r, 8)

ws.freeze_panes = "A6"

# ══════════════════════════════════════════════════════════════════
for sheet in wb.sheetnames:
    wb[sheet].sheet_view.showGridLines = False

wb.save("개인자산관리_양식.xlsx")
print("saved")
