"""투썸플레이스 모바일 메인(mo.twosome.co.kr) 제품 크롤러.

메인 페이지의 각 섹션(새로 나왔어요 / 홀케이크 인기메뉴 등)에서
제품명·제품 이미지를 수집하고, 상세 페이지에서 제품 설명(이미지 또는 텍스트)을 가져온다.
단독 실행: python twosome_crawler.py
"""
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from urllib.parse import urljoin, parse_qs, urlparse

import requests
from bs4 import BeautifulSoup

BASE_URL = "https://mo.twosome.co.kr"
MAIN_URL = f"{BASE_URL}/main.do"
DETAIL_URL = f"{BASE_URL}/mn/menuInfoDetail.do?menuCd={{}}"
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (iPhone; CPU iPhone OS 17_0 like Mac OS X) "
        "AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Mobile/15E148 Safari/604.1"
    ),
    "Accept-Language": "ko-KR,ko;q=0.9",
}
TIMEOUT = 15


def _get(url: str) -> str:
    resp = requests.get(url, headers=HEADERS, timeout=TIMEOUT)
    resp.raise_for_status()
    return resp.text


def fetch_detail(menu_cd: str) -> dict:
    """상세 페이지에서 제품 설명 이미지(alt='메뉴설명')와 설명 텍스트를 수집."""
    try:
        soup = BeautifulSoup(_get(DETAIL_URL.format(menu_cd)), "html.parser")
    except requests.RequestException:
        return {"desc_images": [], "desc_text": ""}

    desc_images = [
        urljoin(BASE_URL, img["src"])
        for img in soup.select('img[alt="메뉴설명"]')
        if img.get("src")
    ]
    desc_el = soup.select_one(".menu-detail-info-title dd p.desc")
    desc_text = desc_el.get_text(" ", strip=True) if desc_el else ""
    return {"desc_images": desc_images, "desc_text": desc_text}


def crawl_twosome(with_detail: bool = True) -> list[dict]:
    """메인 페이지 전체 섹션의 제품 목록을 반환."""
    soup = BeautifulSoup(_get(MAIN_URL), "html.parser")
    crawled_at = datetime.now().strftime("%Y-%m-%d %H:%M")
    products, seen = [], set()

    for section in soup.select("section"):
        items = section.select("dl.product_item")
        if not items:
            continue
        title_el = section.select_one("h2 span")
        section_name = title_el.get_text(strip=True) if title_el else "기타"

        for item in items:
            name_el = item.select_one("dd span")
            img_el = item.select_one("dt img")
            link_el = item.select_one("dt a")
            if not name_el or not img_el:
                continue
            href = link_el.get("href", "") if link_el else ""
            menu_cd = parse_qs(urlparse(href).query).get("menuCd", [""])[0]
            key = (section_name, menu_cd or name_el.get_text(strip=True))
            if key in seen:
                continue
            seen.add(key)
            products.append({
                "section": section_name,
                "name": name_el.get_text(strip=True),
                "image": urljoin(BASE_URL, img_el.get("src", "")),
                "menu_cd": menu_cd,
                "link": DETAIL_URL.format(menu_cd) if menu_cd else urljoin(BASE_URL, href),
                "crawled_at": crawled_at,
                "desc_images": [],
                "desc_text": "",
                "extra": "",
                "is_new": "새로" in section_name,  # '새로 나왔어요!' 섹션
            })

    if with_detail:
        codes = sorted({p["menu_cd"] for p in products if p["menu_cd"]})
        with ThreadPoolExecutor(max_workers=6) as pool:
            details = dict(zip(codes, pool.map(fetch_detail, codes)))
        for p in products:
            p.update(details.get(p["menu_cd"], {"desc_images": [], "desc_text": ""}))

    return products


if __name__ == "__main__":
    for p in crawl_twosome():
        print(f"[{p['section']}] {p['name']}")
        print(f"   제품 이미지 : {p['image']}")
        for src in p["desc_images"]:
            print(f"   제품 설명   : {src}")
        if p["desc_text"]:
            print(f"   설명 텍스트 : {p['desc_text']}")
