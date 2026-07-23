"""
개인 자산관리 양식 생성 스크립트
회사 자금수지 워크북의 '요약자금수지', '자금일보', 'Cash out' 3개 시트 구조를
그대로 참고하여 개인용으로 변환한 워크북을 생성한다.
"""
import datetime

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

FONT_NAME = "맑은 고딕"

# ── 공통 스타일 ──────────────────────────────────────────
thin = Side(style="thin", color="BFBFBF")
border = Border(left=thin, right=thin, top=thin, bottom=thin)

header_fill = PatternFill(fill_type="solid", fgColor="1F3864")      # 진한 남색 (헤더)
category_fill = PatternFill(fill_type="solid", fgColor="D9E1F2")    # 연한 남색 (그룹 라벨)
total_fill = PatternFill(fill_type="solid", fgColor="D6E4F0")       # 연한 파란색 (소계)
grand_fill = PatternFill(fill_type="solid", fgColor="2E74B5")       # 중간 파란색 (합계)
example_fill = PatternFill(fill_type="solid", fgColor="FFF9E6")     # 연노랑 (예시행)

white_bold = Font(name=FONT_NAME, bold=True, color="FFFFFF", size=10)
title_font = Font(name=FONT_NAME, bold=True, size=13)
note_font = Font(name=FONT_NAME, italic=True, size=9, color="0070C0")
label_font = Font(name=FONT_NAME, bold=True, size=10)
black_font = Font(name=FONT_NAME, size=10)
input_font = Font(name=FONT_NAME, size=10, color="0000FF")          # 입력 셀(예시 포함)
total_font = Font(name=FONT_NAME, bold=True, size=10, color="1F3864")
grand_font = Font(name=FONT_NAME, bold=True, size=11, color="FFFFFF")

center = Alignment(horizontal="center", vertical="center", wrap_text=True)
left = Alignment(horizontal="left", vertical="center", wrap_text=True)
right = Alignment(horizontal="right", vertical="center")

MONEY_FMT = '#,##0;[Red](#,##0)'
DATE_FMT = "yyyy-mm-dd"
PCT_FMT = "0.00%"


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
# 1. 요약자금수지  (원본 '요약자금수지' = 일자별 cash in/out + 예금잔고 누적표)
# ══════════════════════════════════════════════════════════════════
ws = wb.active
ws.title = "요약자금수지"
last_col = "M"
title_row(ws, "■ 개인 요약자금수지 (일별, 2026년, 단위: 원)", last_col)
note_row(ws, "※ 파란색 셀에 직접 입력하세요. 지출소계(K)·잔액(L)은 자동 계산됩니다.", last_col)

widths = {"A": 2, "B": 12, "C": 13, "D": 14, "E": 11, "F": 11, "G": 11,
          "H": 11, "I": 11, "J": 11, "K": 13, "L": 14, "M": 14}
for col, w in widths.items():
    ws.column_dimensions[col].width = w

# 기초잔액
ws.merge_cells("B4:K4")
ws["B4"].value = "기초잔액(전월이월)"
ws["B4"].font = label_font
ws["B4"].fill = category_fill
ws["B4"].alignment = center
ws["L4"].value = 3000000
ws["L4"].font = input_font
ws["L4"].number_format = MONEY_FMT
ws["M4"].value = "직접 입력"
ws["M4"].font = note_font
style_range_border(ws, 4, 2, 4, 13)

# 2단 헤더 (원본의 구분/cash in(금액,거래처)/cash out(업체...)/소계/예금잔고/비고 구조를 개인용으로 변환)
ws.merge_cells("B5:B6")
ws["B5"].value = "날짜"
ws.merge_cells("C5:D5")
ws["C5"].value = "수입 (Cash In)"
ws.merge_cells("E5:J5")
ws["E5"].value = "지출 (Cash Out)"
ws.merge_cells("K5:K6")
ws["K5"].value = "지출소계"
ws.merge_cells("L5:L6")
ws["L5"].value = "잔액"
ws.merge_cells("M5:M6")
ws["M5"].value = "비고"
sub_headers = {"C6": "금액", "D6": "내역", "E6": "생활비", "F6": "카드대금",
               "G6": "대출상환", "H6": "고정비\n(월세/보험/통신)", "I6": "저축/투자", "J6": "기타"}
for coord, val in sub_headers.items():
    ws[coord].value = val
style_header_row(ws, 5, 2, 13)
style_header_row(ws, 6, 2, 13)
ws.row_dimensions[5].height = 18
ws.row_dimensions[6].height = 28

# 데이터 행
n_rows = 20
data_start = 7
example = (datetime.date(2026, 7, 1), 3000000, "급여", 0, 0, 0, 800000, 0, 0, "예시")
for i in range(n_rows):
    r = data_start + i
    is_example = i == 0
    if is_example:
        adate, cin, cin_desc, e_liv, e_card, e_loan, e_fix, e_save, e_etc, note = example
    else:
        adate, cin, cin_desc, e_liv, e_card, e_loan, e_fix, e_save, e_etc, note = (
            None, None, "", None, None, None, None, None, None, None)
    ws[f"B{r}"].value = adate
    ws[f"B{r}"].number_format = DATE_FMT
    ws[f"C{r}"].value = cin
    ws[f"C{r}"].number_format = MONEY_FMT
    ws[f"D{r}"].value = cin_desc
    for col, v in zip("EFGHIJ", (e_liv, e_card, e_loan, e_fix, e_save, e_etc)):
        cell = ws[f"{col}{r}"]
        cell.value = v
        cell.number_format = MONEY_FMT
    ws[f"K{r}"].value = f"=SUM(E{r}:J{r})"
    ws[f"K{r}"].font = black_font
    ws[f"K{r}"].number_format = MONEY_FMT
    prev = f"L{r-1}" if i > 0 else "L4"
    ws[f"L{r}"].value = f"={prev}+C{r}-K{r}"
    ws[f"L{r}"].font = black_font
    ws[f"L{r}"].number_format = MONEY_FMT
    if note:
        ws[f"M{r}"].value = note
        ws[f"M{r}"].font = note_font
    for col in "BCDEFGHIJ":
        cell = ws[f"{col}{r}"]
        cell.font = input_font
        if col in ("B",):
            cell.alignment = center
        elif col == "D":
            cell.alignment = left
    if is_example:
        for col in "BCDEFGHIJKLM":
            ws[f"{col}{r}"].fill = example_fill
    style_range_border(ws, r, 2, r, 13)
data_end = data_start + n_rows - 1

total_row = data_end + 1
ws["B{}".format(total_row)].value = "합계"
ws[f"B{total_row}"].font = grand_font
ws[f"B{total_row}"].fill = grand_fill
ws[f"B{total_row}"].alignment = center
ws[f"C{total_row}"].value = f"=SUM(C{data_start}:C{data_end})"
ws[f"C{total_row}"].font = grand_font
ws[f"C{total_row}"].fill = grand_fill
ws[f"C{total_row}"].number_format = MONEY_FMT
ws[f"D{total_row}"].fill = grand_fill
for col in "EFGHIJ":
    ws[f"{col}{total_row}"].fill = grand_fill
ws[f"K{total_row}"].value = f"=SUM(K{data_start}:K{data_end})"
ws[f"K{total_row}"].font = grand_font
ws[f"K{total_row}"].fill = grand_fill
ws[f"K{total_row}"].number_format = MONEY_FMT
ws[f"L{total_row}"].value = f"=L{data_end}"
ws[f"L{total_row}"].font = grand_font
ws[f"L{total_row}"].fill = grand_fill
ws[f"L{total_row}"].number_format = MONEY_FMT
ws[f"M{total_row}"].fill = grand_fill
style_range_border(ws, total_row, 2, total_row, 13)

ws.freeze_panes = "B7"
ws.sheet_view.showGridLines = False

# ══════════════════════════════════════════════════════════════════
# 2. 자금일보  (원본 '자금일보' = 계좌 유형별 전일잔액/증가/감소/금일잔액 현황판)
# ══════════════════════════════════════════════════════════════════
ws = wb.create_sheet("자금일보")
last_col = "K"
title_row(ws, "■ 개인 자금일보 (계좌별 잔액현황, 2026년, 단위: 원)", last_col)
note_row(ws, "※ 파란색 예시행을 참고해 입력하세요. 금일잔액 = 전일잔액 + 증가 - 감소 (자동 계산)", last_col)

widths = {"A": 2, "B": 12, "C": 11, "D": 15, "E": 16, "F": 13, "G": 12,
          "H": 12, "I": 13, "J": 9, "K": 14}
for col, w in widths.items():
    ws.column_dimensions[col].width = w

ws["B4"].value = "구분"
ws["C4"].value = "금융기관"
ws["D4"].value = "계좌별칭"
ws["E4"].value = "계좌번호"
ws["F4"].value = "전일잔액"
ws["G4"].value = "증가"
ws["H4"].value = "감소"
ws["I4"].value = "금일잔액"
ws["J4"].value = "금리"
ws["K4"].value = "비고"
style_header_row(ws, 4, 2, 11)
ws.row_dimensions[4].height = 18

groups = [
    ("예금\n(수시입출금)", [
        ("OO은행", "급여통장", "110-123-456789", 2500000, 3000000, 800000, None, "예시"),
        ("", "", "", None, None, None, None, None),
        ("", "", "", None, None, None, None, None),
    ]),
    ("적금/정기예금", [
        ("", "", "", None, None, None, None, None),
        ("", "", "", None, None, None, None, None),
    ]),
    ("증권/투자계좌", [
        ("", "", "", None, None, None, None, None),
        ("", "", "", None, None, None, None, None),
    ]),
    ("기타\n(카드포인트 등)", [
        ("", "", "", None, None, None, None, None),
    ]),
]

subtotal_rows = []
r = 5
for label, items in groups:
    g_start = r
    for org, alias, acct, prev_bal, inc, dec, rate, note in items:
        is_example = note == "예시"
        ws[f"C{r}"].value = org
        ws[f"D{r}"].value = alias
        ws[f"E{r}"].value = acct
        ws[f"F{r}"].value = prev_bal
        ws[f"F{r}"].number_format = MONEY_FMT
        ws[f"G{r}"].value = inc
        ws[f"G{r}"].number_format = MONEY_FMT
        ws[f"H{r}"].value = dec
        ws[f"H{r}"].number_format = MONEY_FMT
        ws[f"I{r}"].value = f"=F{r}+G{r}-H{r}"
        ws[f"I{r}"].font = black_font
        ws[f"I{r}"].number_format = MONEY_FMT
        ws[f"J{r}"].value = rate
        ws[f"J{r}"].number_format = PCT_FMT
        if note:
            ws[f"K{r}"].value = note
            ws[f"K{r}"].font = note_font
        for col in "CDEFGHJ":
            cell = ws[f"{col}{r}"]
            cell.font = input_font
            cell.alignment = center if col in ("F", "G", "H", "J") else left
        if is_example:
            for col in "BCDEFGHIJK":
                ws[f"{col}{r}"].fill = example_fill
        style_range_border(ws, r, 2, r, 11)
        r += 1
    g_end = r - 1
    ws.merge_cells(f"B{g_start}:B{g_end}")
    ws[f"B{g_start}"].value = label
    ws[f"B{g_start}"].font = label_font
    ws[f"B{g_start}"].fill = category_fill
    ws[f"B{g_start}"].alignment = center

    ws.merge_cells(f"C{r}:E{r}")
    ws[f"C{r}"].value = f"{label.splitlines()[0]} 소계"
    ws[f"C{r}"].font = total_font
    ws[f"C{r}"].fill = total_fill
    ws[f"C{r}"].alignment = center
    ws[f"B{r}"].fill = total_fill
    for col, src in (("F", "F"), ("G", "G"), ("H", "H"), ("I", "I")):
        cell = ws[f"{col}{r}"]
        cell.value = f"=SUM({src}{g_start}:{src}{g_end})"
        cell.font = total_font
        cell.fill = total_fill
        cell.number_format = MONEY_FMT
    for col in ("J", "K"):
        ws[f"{col}{r}"].fill = total_fill
    style_range_border(ws, r, 2, r, 11)
    subtotal_rows.append(r)
    r += 2

grand_row = r
ws.merge_cells(f"B{grand_row}:E{grand_row}")
ws[f"B{grand_row}"].value = "총 자금 합계"
ws[f"B{grand_row}"].font = grand_font
ws[f"B{grand_row}"].fill = grand_fill
ws[f"B{grand_row}"].alignment = center
for col in ("F", "G", "H", "I"):
    formula = "=" + "+".join(f"{col}{sr}" for sr in subtotal_rows)
    cell = ws[f"{col}{grand_row}"]
    cell.value = formula
    cell.font = grand_font
    cell.fill = grand_fill
    cell.number_format = MONEY_FMT
for col in ("J", "K"):
    ws[f"{col}{grand_row}"].fill = grand_fill
style_range_border(ws, grand_row, 2, grand_row, 11)

ws.freeze_panes = "B5"
ws.sheet_view.showGridLines = False

# ══════════════════════════════════════════════════════════════════
# 3. Cash out  (원본 'Cash out' = 지급 예정/확정 내역 상세 리스트)
# ══════════════════════════════════════════════════════════════════
ws = wb.create_sheet("Cash out")
last_col = "J"
title_row(ws, "■ 개인 Cash Out (지출 상세내역, 2026년, 단위: 원)", last_col)
note_row(ws, "※ 파란색 예시행을 참고해 지출 발생 시마다 한 줄씩 입력하세요.", last_col)

widths = {"A": 2, "B": 7, "C": 12, "D": 12, "E": 13, "F": 14, "G": 22, "H": 13, "I": 11, "J": 16}
for col, w in widths.items():
    ws.column_dimensions[col].width = w

headers = ["순번", "지급수단", "지출일", "지출과목", "사용처", "내용(적요)", "금액(원)", "분류", "비고"]
for i, h in enumerate(headers):
    ws.cell(row=4, column=2 + i).value = h
style_header_row(ws, 4, 2, 10)
ws.row_dimensions[4].height = 18

n_rows = 30
data_start = 5
example = ("신용카드", datetime.date(2026, 7, 1), "식비", "OO마트", "주간 장보기", 85000, "변동비", "예시")
for i in range(n_rows):
    r = data_start + i
    is_example = i == 0
    ws[f"B{r}"].value = f"=IF(F{r}=\"\",\"\",ROW()-{data_start - 1})"
    ws[f"B{r}"].font = black_font
    ws[f"B{r}"].alignment = center
    if is_example:
        method, sdate, cat, place, memo, amt, cls, note = example
    else:
        method, sdate, cat, place, memo, amt, cls, note = ("", None, "", "", "", None, "", None)
    ws[f"C{r}"].value = method
    ws[f"D{r}"].value = sdate
    ws[f"D{r}"].number_format = DATE_FMT
    ws[f"E{r}"].value = cat
    ws[f"F{r}"].value = place
    ws[f"G{r}"].value = memo
    ws[f"H{r}"].value = amt
    ws[f"H{r}"].number_format = MONEY_FMT
    ws[f"I{r}"].value = cls
    if note:
        ws[f"J{r}"].value = note
        ws[f"J{r}"].font = note_font
    for col in "CDEFGHI":
        cell = ws[f"{col}{r}"]
        cell.font = input_font
        if col in ("C", "D", "E", "I"):
            cell.alignment = center
        else:
            cell.alignment = left
    if is_example:
        for col in "BCDEFGHIJ":
            ws[f"{col}{r}"].fill = example_fill
    style_range_border(ws, r, 2, r, 10)
data_end = data_start + n_rows - 1

total_row = data_end + 1
ws.merge_cells(f"B{total_row}:F{total_row}")
ws[f"B{total_row}"].value = "합계"
ws[f"B{total_row}"].font = grand_font
ws[f"B{total_row}"].fill = grand_fill
ws[f"B{total_row}"].alignment = center
ws[f"H{total_row}"].value = f"=SUM(H{data_start}:H{data_end})"
ws[f"H{total_row}"].font = grand_font
ws[f"H{total_row}"].fill = grand_fill
ws[f"H{total_row}"].number_format = MONEY_FMT
for col in ("G", "I", "J"):
    ws[f"{col}{total_row}"].fill = grand_fill
style_range_border(ws, total_row, 2, total_row, 10)

ws.freeze_panes = "B5"
ws.sheet_view.showGridLines = False

wb.save("개인자산관리_양식.xlsx")
print("saved")
