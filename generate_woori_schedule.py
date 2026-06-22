# -*- coding: utf-8 -*-
"""
우리은행 이자/원금 상환 스케줄 (2025-12 ~ 만기, 일자별)
기준: '1-1. 차입금현황'
  - 1221-600-517848 : 4.3%, 지급일20, 만기2028-06-20, 이자월납 + 원금분기(3,6,9,12) 5억
  - 1240-200-212936 : 3.22%, 지급일27, 만기2028-11-27, 이자월납 + 만기일시상환 60억
  - 1249-200-013901 : 1.5%, 지급일말일, 만기2036-03-31, 3년거치(29.7~원금균등), 신규실행
"""
import datetime, calendar
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

BANK = "우리은행"
DAY_BASIS = 365


def D(y, m, d): return datetime.date(y, m, d)
def clamp(y, m, d): return datetime.date(y, m, min(d, calendar.monthrange(y, m)[1]))


def month_seq(sy, sm, ey, em):
    y, m = sy, sm
    while (y, m) <= (ey, em):
        yield y, m
        m += 1
        if m > 12:
            m = 1; y += 1


LOANS = [
    {
        "acct": "1221-600-517848", "pay_day": 20, "rate": 0.043,
        "start": (2025, 12), "prev": D(2025, 11, 20),
        "open_bal": 5_500_000_000, "mat": D(2028, 6, 20),
        "principal": {"kind": "periodic", "months": [3, 6, 9, 12], "amt": 500_000_000},
        "note": "원금 분기상환(3·6·9·12월)",
        "actual_int": {(2026, 1): 18260273, (2026, 2): 18260273, (2026, 3): 16493150,
                       (2026, 4): 16434246, (2026, 5): 15904109},
    },
    {
        "acct": "1240-200-212936", "pay_day": 27, "rate": 0.0322,
        "start": (2025, 12), "prev": D(2025, 11, 27),
        "open_bal": 6_000_000_000, "mat": D(2028, 11, 27),
        "principal": {"kind": "bullet"},
        "note": "만기일시상환",
        "actual_int": {(2026, 1): 16204931, (2026, 2): 16204931, (2026, 3): 14820821,
                       (2026, 4): 16408767, (2026, 5): 15879452},
    },
    {
        "acct": "1249-200-013901", "pay_day": 31, "rate": 0.015,
        "start": (2026, 5), "prev": D(2026, 5, 28),     # 2026-05-28 신규실행
        "open_bal": 35_500_000, "mat": D(2036, 3, 31),
        "principal": {"kind": "grace_equal", "from": (2029, 7)},
        "note": "3년거치 후 원금균등(29.7~)·신규실행",
        "actual_int": {(2026, 5): 4376},
    },
]


def principal_for(loan, y, m, is_mat, balance):
    p = loan["principal"]
    if p["kind"] == "periodic":
        if is_mat:
            return balance                      # 만기월: 잔액 전액
        return p["amt"] if m in p["months"] else 0
    if p["kind"] == "bullet":
        return balance if is_mat else 0
    if p["kind"] == "grace_equal":
        fy, fm = p["from"]
        if (y, m) < (fy, fm):
            return 0
        if is_mat:
            return balance
        # 거치종료~만기까지 월수로 균등
        n = sum(1 for _ in month_seq(fy, fm, loan["mat"].year, loan["mat"].month))
        return round(loan["open_bal"] / n)
    return 0


def build(loan):
    rows = []
    balance = loan["open_bal"]
    prev = loan["prev"]
    sy, sm = loan["start"]
    ey, em = loan["mat"].year, loan["mat"].month
    for y, m in month_seq(sy, sm, ey, em):
        cur = clamp(y, m, loan["pay_day"])
        is_mat = (y, m) == (ey, em)
        days = (cur - prev).days
        if (y, m) in loan["actual_int"]:
            interest = loan["actual_int"][(y, m)]
        else:
            interest = round(balance * loan["rate"] * days / DAY_BASIS)
        principal = principal_for(loan, y, m, is_mat, balance)
        principal = min(principal, balance)
        rows.append({"date": cur, "interest": interest, "principal": principal})
        balance -= principal
        prev = cur
        if balance <= 0 and not is_mat:
            # 조기 소진 시 종료
            break
    return rows


def tail(acct): return acct.split("-")[-1]


records = []
for loan in LOANS:
    for s in build(loan):
        base = {"account": loan["acct"], "client": f"우리 {tail(loan['acct'])}"}
        records.append({**base, "date": s["date"], "fund": "이자",
                        "summary": f"{loan['acct']} 이자", "amount": s["interest"]})
        if s["principal"] > 0:
            records.append({**base, "date": s["date"], "fund": "원금",
                            "summary": f"{loan['acct']} 원금", "amount": s["principal"]})

order = {"이자": 0, "원금": 1}
records.sort(key=lambda r: (r["date"], r["account"], order[r["fund"]]))

# ---------------------------------------------------------------------------
wb = Workbook(); ws = wb.active; ws.title = "우리은행상환스케줄"
headers = ["예정일", "확정일", "자금과목", "거래처명", "사업자(주민)번호",
           "은행", "계좌번호", "예금주", "실제예금", "적요", "금액"]
hdr_fill = PatternFill("solid", fgColor="BDD7EE"); hdr_font = Font(bold=True, size=10)
thin = Side(style="thin", color="BFBFBF"); border = Border(thin, thin, thin, thin)
center = Alignment(horizontal="center", vertical="center")
right = Alignment(horizontal="right", vertical="center")
won_green = PatternFill("solid", fgColor="E2EFDA")

ws.cell(1, 1, "■ NRB 우리은행 이자/원금 상환 스케줄 — 2025.12 ~ 만기 (일자별)").font = Font(bold=True, size=12)
ws.cell(2, 1, "※ 2025-12~2026-05 이자는 실적, 그 외 산정(잔액×이율×경과일수÷365) / 원금은 비고 방식").font = Font(size=9, italic=True, color="808080")

hr = 4
for c, h in enumerate(headers, 1):
    cc = ws.cell(hr, c, h); cc.fill = hdr_fill; cc.font = hdr_font; cc.alignment = center; cc.border = border

r = hr + 1
for rec in records:
    is_won = rec["fund"] == "원금"
    vals = [rec["date"], rec["date"], rec["fund"], rec["client"], "",
            BANK, rec["account"], "", "", rec["summary"], rec["amount"]]
    for c, v in enumerate(vals, 1):
        cell = ws.cell(r, c, v); cell.border = border; cell.font = Font(size=10)
        if c in (1, 2): cell.number_format = "yyyy-mm-dd"; cell.alignment = center
        elif c == 3: cell.alignment = center
        elif c == 11: cell.number_format = "#,##0"; cell.alignment = right
        else: cell.alignment = Alignment(vertical="center")
        if is_won: cell.fill = won_green
    r += 1

last = r
ws.cell(last, 10, "합 계").font = Font(bold=True); ws.cell(last, 10).alignment = right
tc = ws.cell(last, 11, f"=SUM(K{hr+1}:K{last-1})"); tc.number_format = "#,##0"
tc.font = Font(bold=True); tc.alignment = right
for c in range(1, 12):
    ws.cell(last, c).fill = PatternFill("solid", fgColor="FFF2CC"); ws.cell(last, c).border = border

for i, w in enumerate([12, 12, 9, 12, 14, 10, 20, 9, 11, 28, 16], 1):
    ws.column_dimensions[get_column_letter(i)].width = w
ws.freeze_panes = "A5"

out = "차입금_이자원금_상환스케줄_우리은행.xlsx"
wb.save(out)
print("saved:", out, "rows:", len(records))

# 검증
for loan in LOANS:
    rs = build(loan)
    psum = sum(x["principal"] for x in rs)
    print(f"  {tail(loan['acct'])}: 원금합 {psum:,} (개시잔액 {loan['open_bal']:,})  이자건수 {len(rs)}")
tot_i = sum(x["amount"] for x in records if x["fund"] == "이자")
tot_p = sum(x["amount"] for x in records if x["fund"] == "원금")
print(f"  총 이자 {tot_i:,} / 총 원금 {tot_p:,} / 행 {len(records)}")
