"""정적 웹페이지(competitor_trend.html)가 읽을 JSON 데이터를 생성한다.

GitHub Actions가 하루 한 번 실행해 competitor_trend/data/*.json 을 갱신한다.
로컬에서 직접 실행해도 된다: python collect.py
"""
import json
import os
from datetime import datetime, timedelta, timezone

import mega_crawler
import news_trend
import starbucks_crawler
import twosome_crawler

KST = timezone(timedelta(hours=9))
DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")

BRANDS = {
    "투썸": ("투썸플레이스", twosome_crawler.MAIN_URL, twosome_crawler.crawl_twosome),
    "스타벅스": ("스타벅스", starbucks_crawler.MENU_URL, starbucks_crawler.crawl_starbucks),
    "메가커피": ("메가MGC커피", mega_crawler.MENU_URL, mega_crawler.crawl_mega),
}
NEWS_KEYWORDS = ["투썸플레이스", "스타벅스 신제품", "메가MGC커피", "카페 신메뉴"]
NEWS_LIMIT = 40


def collect_products() -> dict:
    brands = {}
    for key, (title, url, crawl) in BRANDS.items():
        try:
            items = crawl(True)
            print(f"  {key}: {len(items)}개")
        except Exception as e:  # 한 브랜드가 실패해도 나머지는 저장
            print(f"  {key}: 실패 - {e}")
            items = []
        brands[key] = {"title": title, "source": url, "products": items}
    return brands


def collect_news() -> dict:
    result = {}
    for kw in NEWS_KEYWORDS:
        rows = []
        for name, fetch in (("구글", news_trend.search_google), ("네이버", news_trend.search_naver)):
            try:
                rows += fetch(kw, NEWS_LIMIT)
            except Exception as e:
                print(f"  {kw}/{name}: 실패 - {e}")
        df = news_trend.to_dataframe(rows)
        df["published"] = df["published"].dt.strftime("%Y-%m-%d %H:%M").fillna("")
        result[kw] = df.to_dict("records")
        print(f"  {kw}: {len(df)}건")
    return result


def main():
    os.makedirs(DATA_DIR, exist_ok=True)
    now = datetime.now(KST).strftime("%Y-%m-%d %H:%M")

    print("제품 수집")
    products = collect_products()
    print("뉴스 수집")
    news = collect_news()

    payload = {
        "collected_at": now,
        "brands": products,
        "news": news,
        "news_keywords": NEWS_KEYWORDS,
    }
    path = os.path.join(DATA_DIR, "competitor_data.json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, separators=(",", ":"))
    print(f"저장 완료: {path} ({os.path.getsize(path) / 1024:.0f}KB, 수집 시각 {now})")


if __name__ == "__main__":
    main()
