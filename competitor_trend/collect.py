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
# 네이버 검색 Open API 키 (GitHub Actions Secrets 또는 환경변수).
# 없으면 네이버 뉴스 검색 페이지 파싱으로 자동 대체된다.
NAVER_ID = os.getenv("NAVER_CLIENT_ID", "")
NAVER_SECRET = os.getenv("NAVER_CLIENT_SECRET", "")


def load_previous() -> dict:
    """직전 수집 결과 — 수집 실패한 브랜드는 이 값을 그대로 유지한다."""
    try:
        with open(os.path.join(DATA_DIR, "competitor_data.json"), encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return {}


def collect_products(previous: dict, now: str) -> dict:
    prev_brands = previous.get("brands", {})
    brands = {}
    for key, (title, url, crawl) in BRANDS.items():
        items, error = [], ""
        for attempt in (1, 2, 3):  # 해외 IP 차단·일시 오류 대비 재시도
            try:
                items = crawl(True)
                if items:
                    break
            except Exception as e:
                error = str(e)[:120]
            print(f"  {key}: {attempt}차 시도 실패 - {error or '수집 결과 없음'}")

        prev = prev_brands.get(key, {})
        if items:
            print(f"  {key}: {len(items)}개")
            brands[key] = {"title": title, "source": url, "products": items,
                           "updated_at": now, "stale": False}
        elif prev.get("products"):  # 실패 시 이전 데이터 유지
            print(f"  {key}: 수집 실패 → 이전 데이터 유지 ({prev.get('updated_at', '시각 미상')})")
            brands[key] = {**prev, "title": title, "source": url, "stale": True}
        else:
            brands[key] = {"title": title, "source": url, "products": [],
                           "updated_at": "", "stale": True}
    return brands


def search_naver_with_key(keyword: str, limit: int) -> list:
    return news_trend.search_naver(keyword, limit, NAVER_ID, NAVER_SECRET)


def collect_news(previous: dict) -> dict:
    prev_news = previous.get("news", {})
    result = {}
    for kw in NEWS_KEYWORDS:
        rows = []
        for name, fetch in (("구글", news_trend.search_google), ("네이버", search_naver_with_key)):
            try:
                rows += fetch(kw, NEWS_LIMIT)
            except Exception as e:
                print(f"  {kw}/{name}: 실패 - {e}")
        if not rows:  # 전부 실패하면 이전 뉴스 유지
            result[kw] = prev_news.get(kw, [])
            print(f"  {kw}: 수집 실패 → 이전 데이터 유지 ({len(result[kw])}건)")
            continue
        df = news_trend.to_dataframe(rows)
        df["published"] = df["published"].dt.strftime("%Y-%m-%d %H:%M").fillna("")
        result[kw] = df.to_dict("records")
        print(f"  {kw}: {len(df)}건")
    return result


def main():
    os.makedirs(DATA_DIR, exist_ok=True)
    now = datetime.now(KST).strftime("%Y-%m-%d %H:%M")
    previous = load_previous()

    print("제품 수집")
    products = collect_products(previous, now)
    print("뉴스 수집 (네이버: " + ("Open API" if NAVER_ID and NAVER_SECRET else "검색 페이지 파싱") + ")")
    news = collect_news(previous)

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
