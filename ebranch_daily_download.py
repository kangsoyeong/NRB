"""
e-branch 일별시재조회 자동 엑셀 다운로드
- 스크린샷 분석 기반 좌표 자동화 (이미지 캡처 불필요)
- 매일 오전 9시 실행 or 직접 실행
"""

import os
import sys
import json
import time
import shutil
import logging
import subprocess
from datetime import datetime, timedelta
from pathlib import Path

try:
    import pyautogui
    import pygetwindow as gw
except ImportError:
    print("pip install pyautogui pygetwindow 를 먼저 실행하세요.")
    sys.exit(1)

# ─── 로깅 ─────────────────────────────────────────────────────────────────────

LOG_DIR = Path(__file__).parent / "logs"
LOG_DIR.mkdir(exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(
            LOG_DIR / f"ebranch_{datetime.now():%Y%m}.log", encoding="utf-8"
        ),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger(__name__)

# ─── 설정 ─────────────────────────────────────────────────────────────────────

def load_config() -> dict:
    config_path = Path(__file__).parent / "config.json"
    with open(config_path, encoding="utf-8") as f:
        cfg = json.load(f)
    cfg["download"]["save_folder"] = os.path.expandvars(
        cfg["download"]["save_folder"]
    )
    return cfg


# ─── UI 좌표 (창 좌상단 기준 상대 좌표, 1366×768 스크린샷 분석값) ────────────
#   실제 화면 해상도가 다르면 config.json의 ui_coords 항목에서 조정하세요.

DEFAULT_UI = {
    # 좌측 메뉴 - 일별시재조회 항목
    "menu_daily_inventory": [82, 349],
    # 조회일 날짜 입력 필드 (우측 상단)
    "date_field": [955, 240],
    # Q 조회 버튼
    "search_btn": [793, 353],
    # 엑셀 다운로드 아이콘 (우측 상단 4개 아이콘 중 3번째)
    "excel_btn": [1327, 175],
}

# ─── 유틸 ─────────────────────────────────────────────────────────────────────

def yesterday() -> str:
    return (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")

def yesterday_yyyymmdd() -> str:
    return (datetime.now() - timedelta(days=1)).strftime("%Y%m%d")

def ensure_folder(folder: str) -> Path:
    p = Path(folder)
    p.mkdir(parents=True, exist_ok=True)
    return p

def get_window(title_kw: str = "e-branch") -> gw.Win32Window | None:
    wins = gw.getWindowsWithTitle(title_kw)
    return wins[0] if wins else None

def focus_window(title_kw: str = "e-branch"):
    w = get_window(title_kw)
    if not w:
        return False
    if w.isMinimized:
        w.restore()
    w.activate()
    time.sleep(0.6)
    return True

def abs_coord(rel_x: int, rel_y: int, win: gw.Win32Window) -> tuple[int, int]:
    """창 상대 좌표 → 화면 절대 좌표 변환"""
    return win.left + rel_x, win.top + rel_y

def click(rel_x: int, rel_y: int, win: gw.Win32Window, wait: float = 0.4):
    ax, ay = abs_coord(rel_x, rel_y, win)
    pyautogui.click(ax, ay)
    log.info(f"클릭: 상대({rel_x},{rel_y}) → 절대({ax},{ay})")
    time.sleep(wait)

# ─── 단계별 자동화 ────────────────────────────────────────────────────────────

def launch_app(cfg: dict):
    app_path = cfg["ebranch"]["app_path"]
    if not Path(app_path).exists():
        log.error(f"앱을 찾을 수 없음: {app_path}")
        sys.exit(1)
    log.info("e-branch 실행 중...")
    subprocess.Popen([app_path])
    time.sleep(8)


def wait_for_app(timeout: int = 90) -> gw.Win32Window:
    log.info("e-branch 창 대기 중...")
    deadline = time.time() + timeout
    while time.time() < deadline:
        w = get_window("e-branch")
        if w:
            log.info(f"창 감지: {w.title}")
            return w
        time.sleep(1)
    log.error("e-branch 창을 찾을 수 없습니다.")
    sys.exit(1)


def handle_login(cfg: dict):
    """
    공인인증서 로그인.
    - 자동화가 어려운 경우 사용자가 수동으로 로그인 후 Enter.
    """
    pw = cfg["certificate"].get("password", "")
    if pw:
        log.info("인증서 비밀번호 자동 입력 시도...")
        time.sleep(3)
        pyautogui.typewrite(pw, interval=0.06)
        pyautogui.press("enter")
        time.sleep(5)
    else:
        log.warning("인증서 비밀번호가 설정되지 않았습니다.")
        input(">>> e-branch에 직접 로그인 후 Enter를 누르세요: ")


def navigate_to_inventory(ui: dict, win: gw.Win32Window):
    """좌측 메뉴에서 일별시재조회 클릭"""
    log.info("일별시재조회 메뉴 클릭")
    mx, my = ui["menu_daily_inventory"]
    click(mx, my, win, wait=2)


def set_date_and_search(ui: dict, win: gw.Win32Window):
    """조회일을 전날로 변경 후 조회"""
    date_str = yesterday_yyyymmdd()       # 예: 20260514
    display  = yesterday()                # 예: 2026-05-14
    log.info(f"조회 날짜 설정: {display}")

    # 날짜 필드 클릭 → 전체 선택 → 새 날짜 입력
    dx, dy = ui["date_field"]
    click(dx, dy, win, wait=0.3)
    pyautogui.hotkey("ctrl", "a")
    time.sleep(0.1)
    pyautogui.typewrite(date_str, interval=0.05)
    pyautogui.press("tab")
    time.sleep(0.5)

    # 조회 버튼 클릭
    log.info("조회 버튼 클릭")
    sx, sy = ui["search_btn"]
    click(sx, sy, win, wait=4)          # 데이터 로딩 대기


def click_excel_download(ui: dict, win: gw.Win32Window):
    """우측 상단 엑셀 아이콘 클릭"""
    log.info("엑셀 다운로드 버튼 클릭")
    ex, ey = ui["excel_btn"]
    click(ex, ey, win, wait=3)


def save_excel(cfg: dict) -> bool:
    """
    e-branch 다운로드 위치 → 바탕화면/강소영 폴더로 이동
    e-branch는 보통 다운로드 창(저장 대화상자)을 띄우거나
    기본 다운로드 폴더에 저장합니다.
    """
    wait_sec = cfg["download"]["wait_seconds"]
    log.info(f"파일 저장 대기 ({wait_sec}초)...")
    time.sleep(wait_sec)

    save_dir  = ensure_folder(cfg["download"]["save_folder"])
    prefix    = cfg["download"]["file_prefix"]
    date_tag  = yesterday_yyyymmdd()
    dl_folder = Path(os.path.expanduser("~")) / "Downloads"

    for ext in ("xlsx", "xls"):
        candidates = sorted(
            dl_folder.glob(f"*.{ext}"),
            key=lambda p: p.stat().st_mtime,
            reverse=True,
        )
        for f in candidates:
            # 2분 이내에 생성된 파일만 대상
            if time.time() - f.stat().st_mtime < 120:
                dest = save_dir / f"{prefix}{date_tag}.{ext}"
                shutil.move(str(f), str(dest))
                log.info(f"저장 완료: {dest}")
                return True

    # 저장 대화상자가 열렸을 경우: 경로 직접 입력 후 저장
    dest_path = str(save_dir / f"{prefix}{date_tag}.xlsx")
    log.info("저장 대화상자에 경로 직접 입력 시도")
    time.sleep(1)
    pyautogui.hotkey("ctrl", "l")          # 주소창 열기 (파일탐색기 다이얼로그)
    time.sleep(0.5)
    pyautogui.hotkey("ctrl", "a")
    pyautogui.typewrite(dest_path, interval=0.04)
    pyautogui.press("enter")
    time.sleep(2)
    pyautogui.press("enter")               # 저장 확인
    time.sleep(1)

    if Path(dest_path).exists():
        log.info(f"저장 완료 (대화상자): {dest_path}")
        return True

    log.error("파일을 찾을 수 없습니다. 수동으로 확인하세요.")
    return False


# ─── 메인 ─────────────────────────────────────────────────────────────────────

def main():
    log.info("=" * 55)
    log.info(f"일별시재조회 자동 다운로드  |  대상: {yesterday()}")
    log.info("=" * 55)

    cfg = load_config()
    pyautogui.FAILSAFE = True

    # UI 좌표: config.json에 ui_coords 있으면 덮어쓰기
    ui = {**DEFAULT_UI, **cfg.get("ui_coords", {})}

    # 1. 앱 실행 (이미 실행 중이면 건너뜀)
    win = get_window("e-branch")
    if not win:
        launch_app(cfg)
        win = wait_for_app()
        focus_window()
        handle_login(cfg)
        # 로그인 후 메인화면 로딩 대기
        time.sleep(cfg["ebranch"]["page_load_timeout"])
        win = get_window("e-branch")
    else:
        log.info("e-branch가 이미 실행 중입니다.")
        focus_window()

    if not win:
        log.error("e-branch 창을 찾을 수 없습니다.")
        sys.exit(1)

    # 2. 일별시재조회 메뉴 클릭
    navigate_to_inventory(ui, win)

    # 3. 전날 날짜로 조회
    set_date_and_search(ui, win)

    # 4. 엑셀 다운로드
    click_excel_download(ui, win)

    # 5. 파일 저장
    ok = save_excel(cfg)

    if ok:
        log.info("완료! 바탕화면 > 강소영 폴더를 확인하세요.")
    else:
        log.error("저장 실패. logs 폴더를 확인하세요.")
        sys.exit(1)


if __name__ == "__main__":
    main()
