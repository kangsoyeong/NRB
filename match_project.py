"""
입출금 내역의 '품의내역명'과 자금수지의 '적요'를 80% 유사도 기준으로 매칭하여
자금수지의 '프로젝트'를 입출금 내역에 채워넣는 스크립트.
"""

import os
import sys
import pandas as pd
from rapidfuzz import fuzz, process

# ─── 설정 ──────────────────────────────────────────────────────────────
INPUT_FILE     = "8595a992-4_______.xlsx"
CASHFLOW_FILE  = "8542feab-_________.xlsx"

INPUT_DIR      = "/root/.claude/uploads/0d0eb156-6c70-465c-9460-f5055014f83c"
CASHFLOW_DIR   = "/root/.claude/uploads/4a48c333-2339-407e-bd3f-a433abc945d8"

INPUT_SHEET    = "입,출금 상세(4월)"
CASHFLOW_SHEET = 0

INPUT_HEADER_ROW    = 0
CASHFLOW_HEADER_ROW = 0

INPUT_DESC_COL       = "품의내역명"
CASHFLOW_DESC_COL    = "적요"
CASHFLOW_PROJECT_COL = "프로젝트"
INPUT_PROJECT_COL    = "프로젝트"

THRESHOLD = 80   # 유사도 80% 이상이면 매칭

OUTPUT_FILE = "/home/user/NRB/입출금내역_4월_프로젝트매칭.xlsx"
# ────────────────────────────────────────────────────────────────────────


def score(a: str, b: str, **kwargs) -> float:
    """
    두 문자열의 유사도(0~100).
    - 기본: fuzz.ratio (전체 문자열 유사도)
    - 짧은 쪽이 긴 쪽에 완전히 포함되고 최소 10자 이상이면 partial_ratio도 고려
    """
    r = fuzz.ratio(a, b)
    shorter, longer = (a, b) if len(a) <= len(b) else (b, a)
    if len(shorter) >= 10 and shorter in longer:
        r = max(r, fuzz.partial_ratio(a, b))
    return r


def main():
    input_path    = os.path.join(INPUT_DIR, INPUT_FILE)
    cashflow_path = os.path.join(CASHFLOW_DIR, CASHFLOW_FILE)

    for path, label in [(input_path, "입출금 내역"), (cashflow_path, "자금수지")]:
        if not os.path.exists(path):
            print(f"[오류] {label} 파일 없음: {path}")
            sys.exit(1)

    # ── 자금수지 읽기 ──────────────────────────────────────────────────
    print(f"자금수지 읽는 중: {CASHFLOW_FILE}")
    cf = pd.read_excel(cashflow_path, sheet_name=CASHFLOW_SHEET, header=CASHFLOW_HEADER_ROW)
    cf.columns = cf.columns.astype(str).str.strip()
    cf = cf.dropna(subset=[CASHFLOW_DESC_COL])
    cf[CASHFLOW_DESC_COL] = cf[CASHFLOW_DESC_COL].astype(str).str.strip()
    print(f"  → {len(cf)}행 / 고유 적요: {cf[CASHFLOW_DESC_COL].nunique()}개")

    # 적요 → 프로젝트 dict (중복 적요는 첫 번째 프로젝트 사용)
    desc_to_project = {}
    for _, row in cf.iterrows():
        d = row[CASHFLOW_DESC_COL]
        if d and d not in desc_to_project:
            desc_to_project[d] = str(row[CASHFLOW_PROJECT_COL]) if pd.notna(row[CASHFLOW_PROJECT_COL]) else ""

    cf_descs = list(desc_to_project.keys())

    # ── 입출금 내역 읽기 ───────────────────────────────────────────────
    print(f"\n입출금 내역 읽는 중: {INPUT_FILE}")
    tx = pd.read_excel(input_path, sheet_name=INPUT_SHEET, header=INPUT_HEADER_ROW)
    tx.columns = tx.columns.astype(str).str.strip()
    print(f"  → {len(tx)}행")

    if INPUT_DESC_COL not in tx.columns:
        print(f"[오류] '{INPUT_DESC_COL}' 컬럼 없음. 실제 컬럼: {list(tx.columns)}")
        sys.exit(1)

    if INPUT_PROJECT_COL not in tx.columns:
        tx[INPUT_PROJECT_COL] = ""

    # ── 매칭 ──────────────────────────────────────────────────────────
    print(f"\n매칭 중 (임계값 {THRESHOLD}%)...")
    matched_count = 0
    unmatched_rows = []
    match_log = []

    for idx, row in tx.iterrows():
        desc = str(row[INPUT_DESC_COL]).strip() if pd.notna(row[INPUT_DESC_COL]) else ""

        if not desc or desc in ("nan", ""):
            continue

        result = process.extractOne(desc, cf_descs, scorer=score, score_cutoff=THRESHOLD)

        if result:
            best_desc, best_score, _ = result
            project = desc_to_project[best_desc]
            tx.at[idx, INPUT_PROJECT_COL] = project
            matched_count += 1
            match_log.append({
                "행": idx + 2,
                "품의내역명": desc,
                "매칭 적요": best_desc,
                "유사도": best_score,
                "프로젝트": project,
            })
        else:
            unmatched_rows.append((idx + 2, desc))

    # ── 결과 저장 ─────────────────────────────────────────────────────
    tx.to_excel(OUTPUT_FILE, index=False)
    print(f"\n결과 저장: {OUTPUT_FILE}")

    # ── 요약 ──────────────────────────────────────────────────────────
    total = len([r for r in tx[INPUT_DESC_COL] if pd.notna(r) and str(r).strip() not in ("", "nan")])
    print(f"\n{'='*60}")
    print(f"전체 대상: {total}건")
    print(f"매칭 성공: {matched_count}건")
    print(f"미매칭:   {len(unmatched_rows)}건")
    print(f"{'='*60}")

    if match_log:
        print("\n[매칭 결과 샘플 (상위 20개)]")
        print(f"{'행':>4}  {'유사도':>5}  {'품의내역명':<35}  {'매칭 적요':<35}  프로젝트")
        print("-" * 130)
        for m in match_log[:20]:
            print(f"{m['행']:>4}  {m['유사도']:>5.1f}  {m['품의내역명'][:35]:<35}  {m['매칭 적요'][:35]:<35}  {m['프로젝트']}")

    if unmatched_rows:
        print(f"\n[미매칭 항목 (상위 30개)]")
        print(f"{'행':>4}  품의내역명")
        print("-" * 70)
        for row_num, desc in unmatched_rows[:30]:
            print(f"{row_num:>4}  {desc[:65]}")
        if len(unmatched_rows) > 30:
            print(f"      ... 외 {len(unmatched_rows) - 30}건")

    print("\n완료!")


if __name__ == "__main__":
    main()
