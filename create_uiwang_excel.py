import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "도급기성청구현황"

# ── 스타일 헬퍼 ─────────────────────────────────────────────────
thin = Side(style="thin")
def bdr(): return Border(left=thin, right=thin, top=thin, bottom=thin)
def fl(h): return PatternFill("solid", fgColor=h)
def ft(bold=False, color="000000", size=10, italic=False):
    return Font(name="맑은 고딕", bold=bold, color=color, size=size, italic=italic)
def al(h="center", v="center", wrap=True):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

NUM = '#,##0'
PCT = '0%'

# 색상 팔레트
C_TITLE  = "1F4E79"
C_HEAD   = "2E75B6"
C_CHG    = "FFFFFF"   # 청구 행
C_FAC    = "DEEAF1"   # 공장 행
C_FLD    = "FCE4D6"   # 현장 행
C_M4     = "E2EFDA"   # 4월 입금
C_M5     = "EBF3FA"   # 5월 입금 (파란 계열)
C_UNP    = "FFE0E0"   # 미수금
C_TOTAL  = "D9E1F2"   # 합계 열
C_GRP    = "2E75B6"   # 회차 셀 (좌측)
C_ALT    = "F2F2F2"   # 교대 행 (짝수 회차)

# ── 열 너비 ─────────────────────────────────────────────────────
widths = {"A":7, "B":10, "C":10,
          "D":18, "E":16, "F":15, "G":16, "H":16,
          "I":17, "J":17}
for col, w in widths.items():
    ws.column_dimensions[col].width = w

# ── 데이터 정의 ─────────────────────────────────────────────────
# 각 회차: (회차명, 기준비율그룹, 입금데이터)
# 청구/공장/현장: (구분, 비율, [극동, HJ, 서한, 대저, 엔알비])
# 입금: (구분, 날짜표시, [극동, HJ, 서한, 대저, 엔알비])
rounds = [
    {
        "name": "1회",
        "base": [
            ("청구",  None, [2_588_760_000, 761_400_000, 710_640_000, 507_600_000, 507_600_000]),
            ("공장",  0.43, [1_103_640_000, 324_600_000, 302_960_000, 216_400_000, 216_400_000]),
            ("현장",  0.57, [1_485_120_000, 436_800_000, 407_680_000, 291_200_000, 291_200_000]),
        ],
        "pay": [],
    },
    {
        "name": "2회",
        "base": [
            ("청구",  None, [318_240_000, 93_600_000, 87_360_000, 62_400_000, 62_400_000]),
            ("공장",  0.28, [ 88_230_000, 25_950_000, 24_220_000, 17_300_000, 17_300_000]),
            ("현장",  0.72, [230_010_000, 67_650_000, 63_140_000, 45_100_000, 45_100_000]),
        ],
        "pay": [],
    },
    {
        "name": "3회",
        "base": [
            ("청구",  None, [836_400_000, 246_000_000, 229_600_000, 164_000_000, 164_000_000]),
            ("공장",  0.62, [516_120_000, 151_800_000, 141_680_000, 101_200_000, 101_200_000]),
            ("현장",  0.38, [320_280_000,  94_200_000,  87_920_000,  62_800_000,  62_800_000]),
        ],
        "pay": [],
    },
    {
        "name": "4회",
        "base": [
            ("청구",  None, [1_494_300_000, 439_500_000, 410_200_000, 293_000_000, 293_000_000]),
            ("공장",  0.71, [1_063_350_000, 312_750_000, 291_900_000, 208_500_000, 208_500_000]),
            ("현장",  0.29, [  430_950_000, 126_750_000, 118_300_000,  84_500_000,  84_500_000]),
        ],
        "pay": [
            # (구분, 날짜표시, [극동, HJ, 서한, 대저, 엔알비])
            ("4월 입금", "4/23·4/24", [None, None, None, 293_000_000, 293_000_000]),
            ("5월 입금", "5/7",       [1_494_300_000, None, 410_200_000, None, None]),
            ("미수금",   None,        [None, 439_500_000, None, None, None]),
        ],
    },
    {
        "name": "5회",
        "base": [
            ("청구",  None, [2_029_800_000, 597_000_000, 557_200_000, 398_000_000, 398_000_000]),
            ("공장",  0.59, [1_196_970_000, 352_050_000, 328_580_000, 234_700_000, 234_700_000]),
            ("현장",  0.41, [  832_830_000, 244_950_000, 228_620_000, 163_300_000, 163_300_000]),
        ],
        "pay": [],
    },
]

COMP_COLS = ["D", "E", "F", "G", "H"]  # 업체 5개

# ── 1행: 제목 ────────────────────────────────────────────────────
ws.merge_cells("A1:J1")
c = ws["A1"]
c.value = "의왕초평 A-4BL  도급 기성청구 현황"
c.font = Font(name="맑은 고딕", bold=True, size=15, color="FFFFFF")
c.alignment = al(); c.fill = fl(C_TITLE); c.border = bdr()
ws.row_dimensions[1].height = 38

# ── 2행: 대분류 헤더 ─────────────────────────────────────────────
grp_headers = [
    ("A", "B",  "구분",              C_HEAD),
    ("C", "C",  "비율/날짜",         C_HEAD),
    ("D", "H",  "업체별 금액 (단위: 원, VAT포함)", C_HEAD),
    ("I", "I",  "합 계",             C_HEAD),
    ("J", "J",  "비고",              C_HEAD),
]
for start_col, end_col, label, color in grp_headers:
    if start_col == end_col:
        c = ws[f"{start_col}2"]
    else:
        ws.merge_cells(f"{start_col}2:{end_col}2")
        c = ws[f"{start_col}2"]
    c.value = label
    c.font = Font(name="맑은 고딕", bold=True, size=10, color="FFFFFF")
    c.alignment = al(); c.fill = fl(color); c.border = bdr()
ws.row_dimensions[2].height = 22

# ── 3행: 소분류 헤더 ─────────────────────────────────────────────
sub_headers = [
    ("A", "회차"),
    ("B", "구분"),
    ("C", "비율/날짜"),
    ("D", "극동건설(주)\n51%"),
    ("E", "(주)HJ중공업\n15%"),
    ("F", "(주)서한\n14%"),
    ("G", "(주)대저건설\n10%"),
    ("H", "(주)엔알비\n10%"),
    ("I", "합 계"),
    ("J", "비고"),
]
for col_letter, label in sub_headers:
    c = ws[f"{col_letter}3"]
    c.value = label
    c.font = Font(name="맑은 고딕", bold=True, size=9, color="FFFFFF")
    c.alignment = al(); c.fill = fl(C_HEAD); c.border = bdr()
ws.row_dimensions[3].height = 36

# ── 데이터 행 작성 ───────────────────────────────────────────────
cur_row = 4

ROW_STYLE = {
    "청구":    (C_CHG, "000000", False),
    "공장":    (C_FAC, "1F4E79", False),
    "현장":    (C_FLD, "843C0C", False),
    "4월 입금":(C_M4,  "375623", True),
    "5월 입금":(C_M5,  "1F4E79", True),
    "미수금":  (C_UNP, "C00000", True),
}

for rnd_idx, rnd in enumerate(rounds):
    name       = rnd["name"]
    base_rows  = rnd["base"]
    pay_rows   = rnd["pay"]
    all_rows   = base_rows + [(p[0], p[1], p[2]) for p in pay_rows]
    n_rows     = len(all_rows)
    start_row  = cur_row

    # 회차 셀 (A열 병합)
    if n_rows > 1:
        ws.merge_cells(f"A{start_row}:A{start_row + n_rows - 1}")
    c = ws[f"A{start_row}"]
    c.value     = name
    c.font      = Font(name="맑은 고딕", bold=True, size=12, color="FFFFFF")
    c.alignment = al()
    c.fill      = fl(C_GRP)
    c.border    = bdr()

    for sub_idx, (구분, extra, amounts) in enumerate(all_rows):
        row       = start_row + sub_idx
        bg, tc, is_pay = ROW_STYLE[구분]

        # B: 구분
        c = ws[f"B{row}"]
        c.value     = 구분
        c.font      = ft(bold=True, size=9, color=tc)
        c.alignment = al()
        c.fill      = fl(bg); c.border = bdr()

        # C: 비율 or 날짜
        c = ws[f"C{row}"]
        if isinstance(extra, float):
            c.value          = extra
            c.number_format  = PCT
        else:
            c.value = extra or ("—" if is_pay else "—")
        c.font      = ft(size=9, color=tc)
        c.alignment = al()
        c.fill      = fl(bg); c.border = bdr()

        # D-H: 업체별 금액
        total = 0
        for col_letter, amt in zip(COMP_COLS, amounts):
            c = ws[f"{col_letter}{row}"]
            c.value     = amt
            c.font      = ft(size=9, color=tc if amt else "BBBBBB",
                             bold=(is_pay and amt is not None))
            c.alignment = al(h="right")
            c.fill      = fl(bg); c.border = bdr()
            if amt is not None:
                c.number_format = NUM
                total += amt

        # I: 합계
        c = ws[f"I{row}"]
        c.value         = total if total else None
        c.font          = ft(bold=True, size=9,
                             color=("C00000" if 구분=="미수금" else
                                    ("375623" if "입금" in 구분 else tc)))
        c.alignment     = al(h="right")
        c.fill          = fl(C_TOTAL); c.border = bdr()
        if total:
            c.number_format = NUM

        # J: 비고
        c = ws[f"J{row}"]
        note = ""
        if 구분 == "4월 입금":
            note = "대저건설(4/24) · 엔알비(4/23)"
        elif 구분 == "5월 입금":
            note = "극동건설(5/7) · 서한(5/7)"
        elif 구분 == "미수금":
            note = "HJ중공업 미입금"
        c.value     = note
        c.font      = ft(italic=True, size=8, color="595959")
        c.alignment = al(h="left")
        c.fill      = fl(bg); c.border = bdr()

        ws.row_dimensions[row].height = 21

    cur_row += n_rows

# ── 합계 행 ─────────────────────────────────────────────────────
ws.merge_cells(f"A{cur_row}:C{cur_row}")
c = ws[f"A{cur_row}"]
c.value     = "합   계"
c.font      = Font(name="맑은 고딕", bold=True, size=11, color="FFFFFF")
c.alignment = al(); c.fill = fl(C_TITLE); c.border = bdr()

# 청구 합계만 (공장/현장/입금 제외)
chg_totals = [
    sum(rnd["base"][0][2][i] for rnd in rounds)
    for i in range(5)
]
grand = sum(chg_totals)
for col_letter, val in zip(COMP_COLS, chg_totals):
    c = ws[f"{col_letter}{cur_row}"]
    c.value = val
    c.font = Font(name="맑은 고딕", bold=True, size=10, color="FFFFFF")
    c.alignment = al(h="right"); c.fill = fl(C_TITLE); c.border = bdr()
    c.number_format = NUM

c = ws[f"I{cur_row}"]
c.value = grand
c.font = Font(name="맑은 고딕", bold=True, size=10, color="FFFFFF")
c.alignment = al(h="right"); c.fill = fl(C_TITLE); c.border = bdr()
c.number_format = NUM

c = ws[f"J{cur_row}"]
c.value = "청구금액 기준"
c.font = ft(italic=True, size=9, color="AAAAAA")
c.alignment = al(); c.fill = fl(C_TITLE); c.border = bdr()
ws.row_dimensions[cur_row].height = 26

# ── 범례 ─────────────────────────────────────────────────────────
leg_row = cur_row + 2
legends = [
    ("청구", C_CHG, "000000"),
    ("공장", C_FAC, "1F4E79"),
    ("현장", C_FLD, "843C0C"),
    ("4월 입금", C_M4, "375623"),
    ("5월 입금", C_M5, "1F4E79"),
    ("미수금", C_UNP, "C00000"),
]
ws.cell(row=leg_row, column=1, value="[범례]").font = ft(bold=True, size=9)
for k, (label, bg, tc) in enumerate(legends):
    col = 2 + k
    c = ws.cell(row=leg_row, column=col, value=label)
    c.font = ft(bold=True, size=9, color=tc)
    c.alignment = al()
    c.fill = fl(bg); c.border = bdr()
ws.row_dimensions[leg_row].height = 18

# ── 인쇄/고정 설정 ───────────────────────────────────────────────
ws.freeze_panes = "D4"
ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
ws.page_setup.fitToPage   = True
ws.page_setup.fitToWidth  = 1

out = "/home/user/NRB/의왕초평_도급기성청구현황.xlsx"
wb.save(out)
print("저장 완료:", out)
