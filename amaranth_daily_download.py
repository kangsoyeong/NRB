"""
아마란스(Amaranth) 일별시재조회 자동 엑셀 다운로드
- Selenium 기반 웹 자동화
- 매일 오전 9시 실행 or 직접 실행
"""

import os
import sys
import json
import time
import shutil
import logging
from datetime import datetime, timedelta
from pathlib import Path

try:
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.chrome.service import Service
    from selenium.webdriver.chrome.options import Options
    from webdriver_manager.chrome import ChromeDriverManager
except ImportError:
    print("pip install selenium webdriver-manager 를 먼저 실행하세요.")
    sys.exit(1)

# ─── 로깅 ─────────────────────────────────────────────────────────────────────

LOG_DIR = Path(__file__).parent / "logs"
LOG_DIR.mkdir(exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(
            LOG_DIR / f"amaranth_{datetime.now():%Y%m}.log", encoding="utf-8"
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

# ─── 날짜 ─────────────────────────────────────────────────────────────────────

def yesterday() -> str:
    return (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")

def yesterday_yyyymmdd() -> str:
    return (datetime.now() - timedelta(days=1)).strftime("%Y%m%d")

# ─── 브라우저 설정 ─────────────────────────────────────────────────────────────

def build_driver(save_folder: str) -> webdriver.Chrome:
    options = Options()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-popup-blocking")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_experimental_option("excludeSwitches", ["enable-logging"])
    options.add_experimental_option("prefs", {
        "download.default_directory": str(Path(save_folder).resolve()),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
    })
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    driver.implicitly_wait(10)
    return driver

# ─── 로그인 ───────────────────────────────────────────────────────────────────

def login(driver: webdriver.Chrome, cfg: dict):
    cred = cfg["amaranth"]
    url  = cred["url"]
    log.info(f"접속 중: {url}")
    driver.get(url)
    time.sleep(3)

    wait = WebDriverWait(driver, 30)

    # 아이디 입력
    log.info("아이디 입력")
    id_field = wait.until(
        EC.presence_of_element_located((By.CSS_SELECTOR,
            "input[type='text'], input[name*='id'], input[id*='id'], input[placeholder*='아이디']"
        ))
    )
    id_field.clear()
    id_field.send_keys(cred["user_id"])

    # 비밀번호 입력
    log.info("비밀번호 입력")
    pw_field = driver.find_element(By.CSS_SELECTOR,
        "input[type='password']"
    )
    pw_field.clear()
    pw_field.send_keys(cred["password"])
    pw_field.send_keys(Keys.RETURN)

    time.sleep(5)
    log.info(f"로그인 완료 - 현재 URL: {driver.current_url}")

# ─── 일별시재조회 이동 ─────────────────────────────────────────────────────────

def navigate_to_inventory(driver: webdriver.Chrome, cfg: dict):
    url = cfg["amaranth"]["url"].rstrip("/") + "/#/A/ACQ2060/ACQ2060"
    log.info(f"일별시재조회 이동: {url}")
    driver.get(url)
    time.sleep(4)

# ─── 날짜 설정 및 조회 ─────────────────────────────────────────────────────────

def set_date_and_search(driver: webdriver.Chrome):
    date_str = yesterday()          # 2026-05-14
    date_num = yesterday_yyyymmdd() # 20260514
    log.info(f"조회일 설정: {date_str}")

    wait = WebDriverWait(driver, 20)

    # 날짜 필드 찾기 (조회일)
    date_selectors = [
        "input[id*='date'], input[id*='Date']",
        "input[class*='date'], input[class*='Date']",
        "input[placeholder*='날짜'], input[placeholder*='일자']",
        "input[type='text']",
    ]

    date_field = None
    for sel in date_selectors:
        try:
            fields = driver.find_elements(By.CSS_SELECTOR, sel)
            for f in fields:
                val = f.get_attribute("value") or ""
                # 날짜 형태의 값이 있는 필드 찾기
                if any(c.isdigit() for c in val) and len(val) >= 8:
                    date_field = f
                    break
            if date_field:
                break
        except Exception:
            continue

    if date_field:
        date_field.triple_click() if hasattr(date_field, 'triple_click') else (
            date_field.click(),
            date_field.send_keys(Keys.CONTROL, "a"),
        )
        date_field.clear()
        date_field.send_keys(date_str)
        date_field.send_keys(Keys.TAB)
        time.sleep(0.5)
    else:
        log.warning("날짜 필드를 자동으로 찾지 못했습니다.")

    # 조회 버튼 클릭
    log.info("조회 버튼 클릭")
    search_selectors = [
        "button[id*='search'], button[id*='Search']",
        "button[class*='search'], button[class*='Search']",
        "a[id*='search']",
        "//button[contains(text(),'조회')]",
        "//a[contains(text(),'조회')]",
        "//span[contains(text(),'조회')]",
    ]

    clicked = False
    for sel in search_selectors:
        try:
            if sel.startswith("//"):
                btn = driver.find_element(By.XPATH, sel)
            else:
                btn = driver.find_element(By.CSS_SELECTOR, sel)
            btn.click()
            clicked = True
            log.info(f"조회 클릭: {sel}")
            break
        except Exception:
            continue

    if not clicked:
        log.warning("조회 버튼을 찾지 못했습니다. 화면을 확인하세요.")

    time.sleep(4)

# ─── 엑셀 다운로드 ─────────────────────────────────────────────────────────────

def click_excel(driver: webdriver.Chrome):
    log.info("엑셀 다운로드 클릭")

    excel_selectors = [
        "//button[contains(text(),'엑셀')]",
        "//a[contains(text(),'엑셀')]",
        "//button[contains(text(),'Excel')]",
        "//img[@title='엑셀']",
        "//img[@alt='엑셀']",
        "//button[@title='엑셀']",
        "//a[@title='엑셀']",
        "button[id*='excel'], button[id*='Excel']",
        "a[id*='excel'], a[id*='Excel']",
        "img[src*='excel'], img[src*='Excel']",
    ]

    for sel in excel_selectors:
        try:
            if sel.startswith("//"):
                btn = driver.find_element(By.XPATH, sel)
            else:
                btn = driver.find_element(By.CSS_SELECTOR, sel)
            btn.click()
            log.info(f"엑셀 버튼 클릭 완료: {sel}")
            time.sleep(5)
            return
        except Exception:
            continue

    log.warning("엑셀 버튼을 자동으로 찾지 못했습니다. 스크린샷을 확인하세요.")
    driver.save_screenshot(str(LOG_DIR / f"screenshot_{datetime.now():%H%M%S}.png"))

# ─── 파일 정리 ─────────────────────────────────────────────────────────────────

def collect_file(cfg: dict) -> bool:
    save_dir  = Path(cfg["download"]["save_folder"])
    save_dir.mkdir(parents=True, exist_ok=True)
    prefix    = cfg["download"]["file_prefix"]
    date_tag  = yesterday_yyyymmdd()

    # 다운로드 폴더 = save_folder 로 설정했으므로 바로 확인
    deadline = time.time() + 30
    while time.time() < deadline:
        for ext in ("xlsx", "xls"):
            candidates = sorted(
                save_dir.glob(f"*.{ext}"),
                key=lambda p: p.stat().st_mtime,
                reverse=True,
            )
            for f in candidates:
                if time.time() - f.stat().st_mtime < 60:
                    # 파일명에 날짜가 없으면 이름 변경
                    if date_tag not in f.name:
                        dest = save_dir / f"{prefix}{date_tag}.{ext}"
                        f.rename(dest)
                        log.info(f"저장 완료: {dest}")
                    else:
                        log.info(f"저장 완료: {f}")
                    return True
        time.sleep(2)

    log.error("다운로드 파일을 찾지 못했습니다.")
    return False

# ─── 메인 ─────────────────────────────────────────────────────────────────────

def main():
    log.info("=" * 55)
    log.info(f"아마란스 일별시재조회 자동 다운로드 | 대상: {yesterday()}")
    log.info("=" * 55)

    cfg = load_config()

    save_folder = cfg["download"]["save_folder"]
    Path(save_folder).mkdir(parents=True, exist_ok=True)

    driver = build_driver(save_folder)
    try:
        login(driver, cfg)
        navigate_to_inventory(driver, cfg)
        set_date_and_search(driver)
        click_excel(driver)
        ok = collect_file(cfg)
    finally:
        time.sleep(2)
        driver.quit()

    if ok:
        log.info(f"완료! 바탕화면 > 강소영 폴더를 확인하세요.")
    else:
        log.error("저장 실패. logs 폴더의 스크린샷을 확인하세요.")
        sys.exit(1)


if __name__ == "__main__":
    main()
