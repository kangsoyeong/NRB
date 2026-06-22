# -*- coding: utf-8 -*-
"""
기업은행 원금 상환 스케줄 생성 (이자 제외, 원금만)
기준: '1-1. 차입금현황' 2026-05-31 잔액 → 2026-06부터 만기까지
양식: 예정일/확정일/자금과목/거래처명/사업자(주민)번호/은행/계좌번호/예금주/실제예금/적요/금액
"""
import datetime, calendar
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

BANK = "기업은행"

# ---------------------------------------------------------------------------
# 기업은행 차입금 (2026-05-31 잔액 기준)
#   type: bullet(만기일시) / revolving(수시상환·회전) / installment(원금균등분할)
#   installment: freq_months=상환월 리스트, pay_day, start=(y,m) 개시
# ---------------------------------------------------------------------------
def D(y, m, d): return datetime.date(y, m, d)

LOANS = [
    # --- 만기일시상환 (bullet) ---
    {"acct": "584-029240-32-00033", "bal": 100000000,    "pay_day": 20, "mat": D(2026,6,19),  "type": "bullet", "note": "만기일시"},
    {"acct": "584-029240-32-00034", "bal": 100000000,    "pay_day": 5,  "mat": D(2027,6,24),  "type": "bullet", "note": "만기일시"},
    {"acct": "584-029240-34-00031", "bal": 100000000,    "pay_day": 25, "mat": D(2028,12,22), "type": "bullet", "note": "만기일시"},
    {"acct": "584-029240-34-00032", "bal": 300000000,    "pay_day": 25, "mat": D(2028,12,22), "type": "bullet", "note": "만기일시"},
    {"acct": "584-029240-32-00036", "bal": 5000000000,   "pay_day": 19, "mat": D(2027,5,18),  "type": "bullet", "note": "만기일시"},
    {"acct": "584-029240-32-00037", "bal": 4700000000,   "pay_day": 19, "mat": D(2027,5,18),  "type": "bullet", "note": "만기일시"},
    {"acct": "584-029240-32-00038", "bal": 500000000,    "pay_day": 3,  "mat": D(2026,12,31), "type": "bullet", "note": "만기일시"},
    {"acct": "584-029240-32-00039", "bal": 200000000,    "pay_day": 21, "mat": D(2027,2,19),  "type": "bullet", "note": "만기일시"},
    {"acct": "584-029240-32-00043", "bal": 1000000000,   "pay_day": 21, "mat": D(2027,2,19),  "type": "bullet", "note": "만기일시"},
    {"acct": "584-029240-32-00051", "bal": 2000000000,   "pay_day": 8,  "mat": D(2026,11,7),  "type": "bullet", "note": "만기일시"},
    {"acct": "584-029240-32-00035", "bal": 400000000,    "pay_day": 6,  "mat": D(2027,4,23),  "type": "bullet", "note": "만기일시"},
    {"acct": "584-029240-32-00064", "bal": 20000000000,  "pay_day": 19, "mat": D(2029,1,19),  "type": "bullet", "note": "만기일시"},
    {"acct": "584-029240-32-00070", "bal": 5000000000,   "pay_day": 19, "mat": D(2029,1,19),  "type": "bullet", "note": "만기일시"},
    {"acct": "584-029240-32-00071", "bal": 5000000000,   "pay_day": 19, "mat": D(2029,1,19),  "type": "bullet", "note": "만기일시"},
    # --- 수시상환(회전대출) → 만기 전액상환 가정 ---
    {"acct": "584-029240-31-00031", "bal": 399723203,    "pay_day": 6,  "mat": D(2026,9,11),  "type": "revolving", "note": "수시상환·회전(만기가정)"},
    {"acct": "584-029240-31-00032", "bal": 599882836,    "pay_day": 6,  "mat": D(2027,3,12),  "type": "revolving", "note": "수시상환·회전(만기가정)"},
    {"acct": "584-029240-31-00033", "bal": 800000000,    "pay_day": 6,  "mat": D(2027,2,19),  "type": "revolving", "note": "수시상환·회전(만기가정)"},
    # --- 원금균등분할 ---
    # 32-00050: 분기(2,5,8,11) 균등, 잔액 2,625,000,000 → 6회 × 437,500,000
    {"acct": "584-029240-32-00050", "bal": 2625000000, "pay_day": 10, "mat": D(2027,11,10),
     "type": "installment", "months": [2,5,8,11], "start": (2026,6), "note": "분기 원금균등"},
    # 34-00033: 분기(3,6,9,12) 균등(25년부터), 잔액 1,000,000,000 → 잔여 균등
    {"acct": "584-029240-34-00033", "bal": 1000000000, "pay_day": 25, "mat": D(2031,3,25),
     "type": "installment", "months": [3,6,9,12], "start": (2026,6), "note": "분기 원금균등"},
    # 34-00034: 반기(5,11) 균등, 26년 11월부터, 잔액 4,000,000,000
    {"acct": "584-029240-34-00034", "bal": 4000000000, "pay_day": 16, "mat": D(2032,5,16),
     "type": "installment", "months": [5,11], "start": (2026,11), "note": "반기 원금균등"},
    # 34-00037: 월 균등, 28년 9월부터, 잔액 470,000,000
    {"acct": "584-029240-34-00037", "bal": 470000000, "pay_day": 27, "mat": D(2034,3,27),
     "type": "installment", "months": list(range(1,13)), "start": (2028,9), "note": "월 원금균등"},
]


def clamp_day(y, m, d):
    return datetime.date(y, m, min(d, calendar.monthrange(y, m)[1]))


def installment_dates(loan):
    """개시월~만기월, 지정 상환월에 해당하는 지급일 리스트"""
    dates = []
    sy, sm = loan["start"]
    y, m = sy, sm
    while datetime.date(y, m, 1) <= loan["mat"].replace(day=1):
        if m in loan["months"]:
            dates.append(clamp_day(y, m, loan["pay_day"]))
        m += 1
        if m > 12:
            m = 1; y += 1
    # 만기일은 반드시 마지막 상환일에 포함
    if not dates or dates[-1] != loan["mat"]:
        # 만기월이 상환월 목록에 없으면 만기일을 마지막으로 추가
        if loan["mat"] not in dates:
            dates.append(loan["mat"])
    return dates


def build_rows(loan):
    rows = []
    if loan["type"] in ("bullet", "revolving"):
        rows.append((loan["mat"], loan["bal"]))
    else:  # installment
        dates = installment_dates(loan)
        n = len(dates)
        each = loan["bal"] // n
        # 원 단위 균등, 마지막 회차에 잔액 보정
        remaining = loan["bal"]
        for i, dt in enumerate(dates):
            amt = each if i < n - 1 else remaining
            rows.append((dt, amt))
            remaining -= amt
    return rows


def short_id(acct):
    return "-".join(acct.split("-")[2:])  # 예: 32-00033


# ---------------------------------------------------------------------------
records = []
for loan in LOANS:
    for dt, amt in build_rows(loan):
        records.append({
            "date": dt,
            "client": f"기업 {short_id(loan['acct'])}",
            "account": loan["acct"],
            "summary": f"{loan['acct']} 원금 ({loan['note']})",
            "amount": amt,
        })

records.sort(key=lambda r: (r["date"], r["account"]))

# ---------------------------------------------------------------------------
wb = Workbook()
ws = wb.active
ws.title = "기업은행원금상환스케줄"

headers = ["예정일", "확정일", "자금과목", "거래처명", "사업자(주민)번호",
           "은행", "계좌번호", "예금주", "실제예금", "적요", "금액"]

hdr_fill = PatternFill("solid", fgColor="BDD7EE")
hdr_font = Font(bold=True, size=10)
thin = Side(style="thin", color="BFBFBF")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
center = Alignment(horizontal="center", vertical="center")
right = Alignment(horizontal="right", vertical="center")
won_green = PatternFill("solid", fgColor="E2EFDA")

ws.cell(1, 1, "■ NRB 기업은행 원금 상환 스케줄 — 2026.06 이후 예상 (이자 제외)").font = Font(bold=True, size=12)
ws.cell(2, 1, "※ 2026-05-31 잔액 기준 / 만기일시·수시상환=만기 전액 / 원금균등=비고 주기로 잔액 균등분할").font = Font(size=9, italic=True, color="808080")

hr = 4
for c, h in enumerate(headers, 1):
    cell = ws.cell(hr, c, h)
    cell.fill = hdr_fill; cell.font = hdr_font; cell.alignment = center; cell.border = border

r = hr + 1
for rec in records:
    vals = [rec["date"], rec["date"], "원금", rec["client"], "",
            BANK, rec["account"], "", "", rec["summary"], rec["amount"]]
    for c, v in enumerate(vals, 1):
        cell = ws.cell(r, c, v)
        cell.border = border; cell.font = Font(size=10); cell.fill = won_green
        if c in (1, 2):
            cell.number_format = "yyyy-mm-dd"; cell.alignment = center
        elif c == 3:
            cell.alignment = center
        elif c == 11:
            cell.number_format = "#,##0"; cell.alignment = right
        else:
            cell.alignment = Alignment(vertical="center")
    r += 1

last = r
ws.cell(last, 10, "합 계").font = Font(bold=True)
ws.cell(last, 10).alignment = right
tc = ws.cell(last, 11, f"=SUM(K{hr+1}:K{last-1})")
tc.number_format = "#,##0"; tc.font = Font(bold=True); tc.alignment = right
for c in range(1, 12):
    ws.cell(last, c).fill = PatternFill("solid", fgColor="FFF2CC"); ws.cell(last, c).border = border

widths = [12, 12, 9, 13, 14, 10, 22, 9, 11, 34, 16]
for i, w in enumerate(widths, 1):
    ws.column_dimensions[get_column_letter(i)].width = w
ws.freeze_panes = "A5"

out = "차입금_원금상환스케줄_기업은행.xlsx"
wb.save(out)
print("saved:", out, "rows:", len(records))
total = sum(x["amount"] for x in records)
print(f"총 원금 거래 행: {len(records)},  총 원금액: {total:,}")
# 차입금별 합계 검증
chk = {}
for l in LOANS:
    s = sum(a for _, a in build_rows(l))
    chk[short_id(l['acct'])] = (s, l['bal'], s == l['bal'])
for k, v in chk.items():
    flag = "OK" if v[2] else "MISMATCH"
    print(f"  {k}: 스케줄합 {v[0]:,} vs 잔액 {v[1]:,} [{flag}]")
