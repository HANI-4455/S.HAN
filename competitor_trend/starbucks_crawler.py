"""스타벅스 코리아 메뉴 크롤러.

스타벅스 메뉴 페이지는 카테고리별 JSON(/upload/json/menu/<카테고리코드>.js)을 불러와 그린다.
이 JSON을 직접 받아 제품명·이미지·설명·출시일·NEW 여부를 수집한다.
단독 실행: python starbucks_crawler.py
"""
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime

import requests

BASE_URL = "https://www.starbucks.co.kr"
IMAGE_HOST = "https://image.istarbucks.co.kr"
MENU_URL = f"{BASE_URL}/menu/drink_list.do"
JSON_URL = f"{BASE_URL}/upload/json/menu/{{}}.js"
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/130 Safari/537.36",
           "Referer": MENU_URL}
TIMEOUT = 15
NEW_SECTION = "🆕 신제품(NEW)"

# (카테고리코드, 구분) — drink_list.do / food_list.do 스크립트에 정의된 코드
CATEGORIES = [
    ("W0000171", "drink"), ("W0000060", "drink"), ("W0000003", "drink"), ("W0000004", "drink"),
    ("W0000005", "drink"), ("W0000422", "drink"), ("W0000061", "drink"), ("W0000075", "drink"),
    ("W0000053", "drink"), ("W0000062", "drink"),
    ("W0000013", "food"), ("W0000032", "food"), ("W0000074", "food"), ("W0000033", "food"),
    ("W0000054", "food"), ("W0000055", "food"), ("W0000056", "food"), ("W0000064", "food"),
]


def _fetch_category(code: str) -> list[dict]:
    try:
        resp = requests.post(JSON_URL.format(code), headers=HEADERS, timeout=TIMEOUT)
        resp.raise_for_status()
        return resp.json().get("list", [])
    except (requests.RequestException, ValueError):
        return []


def _fmt_date(yyyymmdd: str) -> str:
    return f"{yyyymmdd[:4]}-{yyyymmdd[4:6]}-{yyyymmdd[6:8]}" if len(yyyymmdd) == 8 else ""


def crawl_starbucks(with_detail: bool = True) -> list[dict]:
    """카테고리별 제품 목록. NEW 아이콘이 붙은 제품은 '신제품' 섹션에도 넣는다."""
    crawled_at = datetime.now().strftime("%Y-%m-%d %H:%M")
    with ThreadPoolExecutor(max_workers=6) as pool:
        results = list(pool.map(_fetch_category, [c for c, _ in CATEGORIES]))

    products, new_items, seen = [], [], set()
    for (_, kind), items in zip(CATEGORIES, results):
        for it in items:
            code = it.get("product_CD", "")
            if not code or code in seen:
                continue
            seen.add(code)
            released = _fmt_date(it.get("new_SDATE", ""))
            p = {
                "section": it.get("cate_NAME") or "기타",
                "name": it.get("product_NM", "").strip(),
                "image": IMAGE_HOST + it["file_PATH"] if it.get("file_PATH") else "",
                "menu_cd": code,
                "link": f"{BASE_URL}/menu/{kind}_view.do?product_cd={code}",
                "crawled_at": crawled_at,
                "desc_images": [],
                "desc_text": (it.get("content") or "").strip() if with_detail else "",
                "extra": " · ".join(x for x in [
                    f"출시 {released}" if released else "",
                    f"{it['kcal']}kcal" if it.get("kcal") else "",
                ] if x),
                "is_new": it.get("newicon") == "Y",
            }
            products.append(p)
            if p["is_new"]:
                new_items.append({**p, "section": NEW_SECTION})

    # 신제품은 출시일 최신순으로 맨 앞 섹션에 배치
    new_items.sort(key=lambda p: p["extra"], reverse=True)
    return new_items + products


if __name__ == "__main__":
    for p in crawl_starbucks():
        if p["section"] == NEW_SECTION:
            print(f"[{p['section']}] {p['name']} ({p['extra']})\n   {p['image']}")
