"""
e-branch UI 이미지 캡처 도구
- 자동화에 필요한 버튼/필드 이미지를 캡처합니다
- 이 스크립트를 먼저 실행해 각 이미지를 저장한 뒤 자동화를 실행하세요
"""

import time
from pathlib import Path

try:
    import pyautogui
    from PIL import ImageGrab
except ImportError:
    print("pip install pyautogui pillow 를 먼저 실행하세요.")
    raise

IMAGES_DIR = Path(__file__).parent / "images"
IMAGES_DIR.mkdir(exist_ok=True)

CAPTURES = [
    ("cert_button.png",          "인증서 선택/로그인 버튼"),
    ("menu_daily_inventory.png", "일별시재조회 메뉴 항목"),
    ("date_field.png",           "조회 시작 날짜 입력 필드"),
    ("date_field_end.png",       "조회 종료 날짜 입력 필드 (있는 경우)"),
    ("btn_search.png",           "조회 버튼"),
    ("btn_excel.png",            "엑셀 다운로드 버튼"),
]


def capture_region(filename: str, description: str):
    """5초 카운트다운 후 마우스 위치 중심으로 80x30 영역 캡처"""
    out_path = IMAGES_DIR / filename
    print(f"\n[캡처] {description}")
    print(f"  → 저장 위치: {out_path}")
    print("  → e-branch 앱에서 해당 요소 위에 마우스를 올리세요.")
    for i in range(5, 0, -1):
        print(f"     {i}초 후 캡처...", end="\r")
        time.sleep(1)

    x, y = pyautogui.position()
    # 요소를 넉넉히 포함하도록 160x50 영역 캡처
    left, top = x - 80, y - 25
    region = (left, top, left + 160, top + 50)
    img = ImageGrab.grab(bbox=region)
    img.save(str(out_path))
    print(f"  ✓ 캡처 완료: ({x}, {y}) 기준 160x50px → {filename}     ")


def main():
    print("=" * 55)
    print("  e-branch UI 이미지 캡처 도구")
    print("=" * 55)
    print("각 항목마다 5초 카운트다운이 시작되면")
    print("e-branch 화면에서 해당 버튼/필드 위에 마우스를 올려두세요.\n")

    for filename, description in CAPTURES:
        answer = input(f"'{description}' 캡처를 진행할까요? [Y/n] ").strip().lower()
        if answer == "n":
            print("  건너뜀.")
            continue
        capture_region(filename, description)

    print("\n모든 이미지 캡처 완료!")
    print(f"저장 위치: {IMAGES_DIR}")
    print("\n이제 ebranch_daily_download.py 를 실행하거나")
    print("register_task_scheduler.bat 으로 자동화를 등록하세요.")


if __name__ == "__main__":
    main()
