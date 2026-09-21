"""메가MGC커피 메뉴 크롤러.

메뉴 페이지(/menu/)는 menu.php에서 목록 HTML 조각을 받아 그린다.
대분류(음료/푸드/상품) + 소분류 체크박스 값으로 menu.php를 페이지별로 호출해 수집한다.
단독 실행: python mega_crawler.py
"""
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime

import requests
from bs4 import BeautifulSoup

BASE_URL = "https://www.mega-mgccoffee.com"
MENU_URL = f"{BASE_URL}/menu/"
LIST_URL = f"{BASE_URL}/menu/menu.php"
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/130 Safari/537.36",
           "Referer": MENU_URL}
TIMEOUT = 15
MAX_PAGES = 20

# (섹션명, 대분류, 소분류 체크박스 값) — 신상품을 맨 앞에 배치
CATEGORIES = [
    ("🆕 신상품(음료)", 1, 9), ("🆕 신상품(푸드)", 2, 10),
    ("커피", 1, 1), ("티", 1, 2), ("에이드&주스", 1, 3), ("스무디&프라페", 1, 4),
    ("디카페인", 1, 5), ("음료", 1, 6), ("디저트", 2, 7), ("MD", 3, 8),
]


def _text(el) -> str:
    return " ".join(el.get_text(" ", strip=True).split()) if el else ""


def _parse_items(html: str) -> list[dict]:
    soup = BeautifulSoup(html, "html.parser")
    items = []
    for li in soup.select("#menu_list > li"):
        name = _text(li.select_one(".cont_gallery_list_box .cont_text_title b"))
        img = li.select_one(".cont_gallery_list_img img")
        if not name or not img:
            continue
        modal = li.select_one(".inner_modal")
        label = _text(li.select_one(".cont_gallery_list_label"))  # HOT / ICE
        info = [_text(x) for x in modal.select(".cont_text_inner")] if modal else []
        kcal = next((x.replace("1회 제공량", "").strip() for x in info if "kcal" in x), "")
        items.append({
            "name": name,
            "image": img.get("src", ""),
            "desc_text": _text(li.select_one(".cont_gallery_list_box .text2")),
            "extra": " · ".join(x for x in [label, kcal] if x),
        })
    return items


def _fetch_category(cat: tuple) -> list[dict]:
    _, cat1, cat2 = cat
    rows = []
    for page in range(1, MAX_PAGES + 1):
        params = {"page": page, "menu_category1": cat1, "menu_category2": cat1, "category": cat2}
        try:
            resp = requests.get(LIST_URL, params=params, headers=HEADERS, timeout=TIMEOUT)
            resp.raise_for_status()
        except requests.RequestException:
            break
        items = _parse_items(resp.text)
        if not items:
            break
        rows += items
        if f"data-page='{page + 1}'" not in resp.text:  # 다음 페이지 없음
            break
    return rows


def crawl_mega(with_detail: bool = True) -> list[dict]:
    crawled_at = datetime.now().strftime("%Y-%m-%d %H:%M")
    with ThreadPoolExecutor(max_workers=5) as pool:
        results = list(pool.map(_fetch_category, CATEGORIES))

    products = []
    for (section, cat1, cat2), items in zip(CATEGORIES, results):
        seen = set()
        for it in items:
            if it["name"] in seen:
                continue
            seen.add(it["name"])
            products.append({
                "section": section,
                "name": it["name"],
                "image": it["image"],
                "menu_cd": "",
                "link": f"{MENU_URL}?menu_category1={cat1}&menu_category2={cat1}",
                "crawled_at": crawled_at,
                "desc_images": [],
                "desc_text": it["desc_text"] if with_detail else "",
                "extra": it["extra"],
                "is_new": section.startswith("🆕"),
            })
    return products


if __name__ == "__main__":
    for p in crawl_mega():
        if p["is_new"]:
            print(f"[{p['section']}] {p['name']} ({p['extra']})\n   {p['image']}")
