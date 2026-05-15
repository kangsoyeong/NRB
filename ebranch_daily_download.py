"""
e-branch 일별시재조회 자동 엑셀 다운로드 스크립트
- 매일 오전 9시 실행 (Windows 작업 스케줄러 등록 필요)
- 전날 날짜 기준으로 일별시재 엑셀 다운로드
"""

import json
import os
import sys
import time
import logging
import subprocess
import shutil
from datetime import datetime, timedelta
from pathlib import Path

try:
    import pyautogui
    import pygetwindow as gw
except ImportError:
    print("필수 패키지를 설치해주세요: pip install pyautogui pygetwindow")
    sys.exit(1)

# ─── 로깅 설정 ───────────────────────────────────────────────────────────────

LOG_DIR = Path(__file__).parent / "logs"
LOG_DIR.mkdir(exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_DIR / f"ebranch_{datetime.now():%Y%m}.log", encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger(__name__)

# ─── 설정 로드 ───────────────────────────────────────────────────────────────

def load_config() -> dict:
    config_path = Path(__file__).parent / "config.json"
    if not config_path.exists():
        log.error("config.json 파일이 없습니다.")
        sys.exit(1)
    with open(config_path, encoding="utf-8") as f:
        cfg = json.load(f)
    # 환경변수 치환 (%USERNAME% 등)
    save_folder = cfg["download"]["save_folder"]
    cfg["download"]["save_folder"] = os.path.expandvars(save_folder)
    return cfg


# ─── 유틸 ────────────────────────────────────────────────────────────────────

def yesterday() -> str:
    """전날 날짜 반환 (YYYY-MM-DD)"""
    return (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")


def yesterday_display() -> str:
    """전날 날짜 반환 (YYYYMMDD) - 파일명용"""
    return (datetime.now() - timedelta(days=1)).strftime("%Y%m%d")


def ensure_save_folder(folder: str) -> Path:
    path = Path(folder)
    path.mkdir(parents=True, exist_ok=True)
    return path


def wait_for_window(title_keyword: str, timeout: int = 60) -> bool:
    """지정 키워드를 포함한 창이 열릴 때까지 대기"""
    log.info(f"창 대기 중: '{title_keyword}' (최대 {timeout}초)")
    deadline = time.time() + timeout
    while time.time() < deadline:
        windows = gw.getWindowsWithTitle(title_keyword)
        if windows:
            log.info(f"창 발견: {windows[0].title}")
            time.sleep(1)
            return True
        time.sleep(1)
    log.error(f"창을 찾을 수 없음: '{title_keyword}'")
    return False


def bring_window_to_front(title_keyword: str):
    windows = gw.getWindowsWithTitle(title_keyword)
    if windows:
        w = windows[0]
        if w.isMinimized:
            w.restore()
        w.activate()
        time.sleep(0.5)


def find_and_click_image(image_path: str, confidence: float = 0.8, timeout: int = 30) -> bool:
    """화면에서 이미지를 찾아 클릭 (이미지 파일이 없으면 건너뜀)"""
    if not Path(image_path).exists():
        log.warning(f"이미지 파일 없음: {image_path} (건너뜀)")
        return False
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            loc = pyautogui.locateOnScreen(image_path, confidence=confidence)
            if loc:
                pyautogui.click(pyautogui.center(loc))
                log.info(f"이미지 클릭 완료: {image_path}")
                return True
        except Exception:
            pass
        time.sleep(1)
    log.error(f"이미지를 화면에서 찾을 수 없음: {image_path}")
    return False


# ─── 핵심 자동화 단계 ─────────────────────────────────────────────────────────

def launch_app(app_path: str):
    """e-branch 앱 실행"""
    if not Path(app_path).exists():
        log.error(f"앱 경로가 잘못됨: {app_path}")
        sys.exit(1)
    log.info(f"e-branch 실행: {app_path}")
    subprocess.Popen([app_path])
    time.sleep(5)


def handle_certificate_login(cfg: dict):
    """
    공인인증서 로그인 처리.
    - images/ 폴더에 인증서 관련 버튼 이미지를 저장하면 자동 클릭.
    - 인증서 비밀번호는 config.json의 certificate.password 에 입력.
    """
    images_dir = Path(__file__).parent / "images"

    # 인증서 선택 버튼 클릭 (이미지 파일 준비 필요)
    cert_btn = str(images_dir / "cert_button.png")
    if find_and_click_image(cert_btn, timeout=cfg["ebranch"]["login_timeout"]):
        time.sleep(1)

    # 인증서 비밀번호 입력
    password = cfg["certificate"]["password"]
    if password:
        log.info("인증서 비밀번호 입력")
        pyautogui.typewrite(password, interval=0.05)
        pyautogui.press("enter")
        time.sleep(3)
    else:
        log.warning(
            "인증서 비밀번호가 config.json에 설정되지 않았습니다. "
            "수동으로 로그인 후 Enter를 누르세요."
        )
        input("로그인 완료 후 Enter를 누르세요: ")


def navigate_to_daily_inventory(cfg: dict):
    """
    일별시재조회 메뉴로 이동.
    - images/menu_daily_inventory.png 파일을 캡처해 두면 자동 탐색.
    - 없으면 수동 안내 메시지 출력.
    """
    images_dir = Path(__file__).parent / "images"
    menu_img = str(images_dir / "menu_daily_inventory.png")

    if not find_and_click_image(menu_img, timeout=cfg["ebranch"]["page_load_timeout"]):
        log.warning(
            "메뉴 이미지(images/menu_daily_inventory.png)가 없습니다. "
            "수동으로 일별시재조회 메뉴로 이동 후 Enter를 누르세요."
        )
        input("일별시재조회 화면 진입 후 Enter를 누르세요: ")


def set_yesterday_date(cfg: dict):
    """
    조회 날짜를 전날로 설정.
    - images/date_field.png 를 캡처해두면 자동 설정.
    - 없으면 수동 안내.
    """
    images_dir = Path(__file__).parent / "images"
    date_field_img = str(images_dir / "date_field.png")
    date_str = yesterday()

    if find_and_click_image(date_field_img, timeout=cfg["ebranch"]["page_load_timeout"]):
        pyautogui.hotkey("ctrl", "a")
        pyautogui.typewrite(date_str.replace("-", ""), interval=0.05)
        time.sleep(0.5)

        # 시작일과 종료일이 별도인 경우
        end_date_img = str(images_dir / "date_field_end.png")
        if find_and_click_image(end_date_img, timeout=5):
            pyautogui.hotkey("ctrl", "a")
            pyautogui.typewrite(date_str.replace("-", ""), interval=0.05)
            time.sleep(0.5)

        # 조회 버튼 클릭
        search_img = str(images_dir / "btn_search.png")
        find_and_click_image(search_img, timeout=10)
        time.sleep(cfg["ebranch"]["page_load_timeout"])
    else:
        log.warning(
            "날짜 필드 이미지(images/date_field.png)가 없습니다. "
            f"날짜를 {date_str} 로 수동 설정 후 조회하고 Enter를 누르세요."
        )
        input("날짜 설정 및 조회 완료 후 Enter를 누르세요: ")


def click_excel_download(cfg: dict) -> bool:
    """엑셀 다운로드 버튼 클릭"""
    images_dir = Path(__file__).parent / "images"
    excel_btn_img = str(images_dir / "btn_excel.png")

    if not find_and_click_image(excel_btn_img, timeout=15):
        log.warning(
            "엑셀 버튼 이미지(images/btn_excel.png)가 없습니다. "
            "수동으로 엑셀 다운로드 버튼을 클릭 후 Enter를 누르세요."
        )
        input("엑셀 다운로드 버튼 클릭 후 Enter를 누르세요: ")
    return True


def collect_downloaded_file(cfg: dict) -> bool:
    """
    브라우저/앱의 기본 다운로드 폴더에서 파일을 찾아
    지정 폴더로 이동 후 날짜 기반 파일명으로 변경.
    """
    wait_sec = cfg["download"]["wait_seconds"]
    log.info(f"다운로드 완료 대기 ({wait_sec}초)...")
    time.sleep(wait_sec)

    # 기본 다운로드 폴더
    default_dl = Path(os.path.expanduser("~")) / "Downloads"
    save_folder = ensure_save_folder(cfg["download"]["save_folder"])
    prefix = cfg["download"]["file_prefix"]
    date_str = yesterday_display()
    target_name = f"{prefix}{date_str}.xlsx"
    target_path = save_folder / target_name

    # 최근 변경된 xlsx 파일 탐색 (30초 이내)
    candidates = sorted(
        default_dl.glob("*.xlsx"),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    now = time.time()
    for f in candidates:
        if now - f.stat().st_mtime < 120:
            shutil.move(str(f), str(target_path))
            log.info(f"파일 저장 완료: {target_path}")
            return True

    # xls 확장자도 확인
    candidates_xls = sorted(
        default_dl.glob("*.xls"),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    for f in candidates_xls:
        if now - f.stat().st_mtime < 120:
            dest = save_folder / f"{prefix}{date_str}.xls"
            shutil.move(str(f), str(dest))
            log.info(f"파일 저장 완료: {dest}")
            return True

    log.error("다운로드된 엑셀 파일을 찾을 수 없습니다.")
    return False


# ─── 메인 ────────────────────────────────────────────────────────────────────

def main():
    log.info("=" * 60)
    log.info("e-branch 일별시재조회 자동 다운로드 시작")
    log.info(f"대상 날짜: {yesterday()}")
    log.info("=" * 60)

    cfg = load_config()

    pyautogui.FAILSAFE = True  # 마우스를 좌상단으로 이동 시 중단

    # 1. 앱 실행
    launch_app(cfg["ebranch"]["app_path"])

    # 2. 창 열릴 때까지 대기
    if not wait_for_window(cfg["ebranch"]["window_title"], timeout=cfg["ebranch"]["login_timeout"]):
        log.error("e-branch 창을 열 수 없습니다. 종료합니다.")
        sys.exit(1)

    bring_window_to_front(cfg["ebranch"]["window_title"])

    # 3. 인증서 로그인
    handle_certificate_login(cfg)

    # 4. 앱이 로드될 때까지 대기
    time.sleep(cfg["ebranch"]["page_load_timeout"])
    bring_window_to_front(cfg["ebranch"]["window_title"])

    # 5. 일별시재조회 메뉴로 이동
    navigate_to_daily_inventory(cfg)

    # 6. 전날 날짜 설정 및 조회
    set_yesterday_date(cfg)

    # 7. 엑셀 다운로드 클릭
    click_excel_download(cfg)

    # 8. 파일을 지정 폴더로 이동
    success = collect_downloaded_file(cfg)

    if success:
        log.info("자동 다운로드 완료!")
    else:
        log.error("파일 수집 실패. 로그를 확인하세요.")
        sys.exit(1)


if __name__ == "__main__":
    main()
