"""구글 / 네이버 뉴스 RSS 검색 + 간단 트렌드 분석.

- 구글: Google News RSS (키 불필요)
- 네이버: 네이버 검색 Open API의 news.xml (RSS 2.0 형식, Client ID/Secret 필요)
          (네이버는 뉴스 검색 RSS를 폐지해 API가 공식 경로)
          키가 없으면 네이버 뉴스 검색 결과 페이지를 파싱해 대체 수집
"""
import html
import re
from collections import Counter
from email.utils import parsedate_to_datetime
from urllib.parse import quote

import feedparser
import pandas as pd
import requests
from bs4 import BeautifulSoup

GOOGLE_RSS = "https://news.google.com/rss/search?q={q}&hl=ko&gl=KR&ceid=KR:ko"
NAVER_API = "https://openapi.naver.com/v1/search/news.xml?query={q}&display={n}&sort=date"
# 네이버 뉴스 탭의 "더보기" 페이징 API (HTML 조각을 JSON으로 반환)
NAVER_WEB = "https://s.search.naver.com/p/newssearch/3/api/tab/more?ssc=tab.news.all&query={q}&sort=1&start={start}"
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130 Safari/537.36"}

STOPWORDS = {
    "있다", "했다", "한다", "위해", "통해", "대한", "및", "등", "더", "첫", "이번", "올해",
    "지난", "오는", "열림", "관련", "기자", "뉴스", "종합", "속보", "단독", "포토", "사진", "영상",
    "the", "and", "for", "with",
}
TAG_RE = re.compile(r"<[^>]+>")
TOKEN_RE = re.compile(r"[가-힣A-Za-z0-9]{2,}")
# 조사 제거용(과도한 형태소 분석 대신 간단한 규칙)
JOSA_RE = re.compile(r"(으로|에서|에게|까지|부터|이다|은|는|이|가|을|를|의|에|와|과|도|로)$")


def _clean(text: str) -> str:
    return html.unescape(TAG_RE.sub("", text or "")).strip()


def _parse_date(value: str):
    try:
        return parsedate_to_datetime(value).astimezone().replace(tzinfo=None)
    except (TypeError, ValueError):
        return pd.NaT


def _rows_from_feed(feed, portal: str) -> list[dict]:
    rows = []
    for e in feed.entries:
        title = _clean(e.get("title", ""))
        source = e.get("source", {}).get("title", "") if portal == "구글" else ""
        # 구글 제목은 "기사 제목 - 언론사" 형식
        if portal == "구글" and " - " in title:
            title, source = title.rsplit(" - ", 1)[0], source or title.rsplit(" - ", 1)[1]
        summary = _clean(e.get("description", ""))
        if portal == "구글":
            summary = ""  # 구글 RSS description은 제목 반복이라 제외
        rows.append({
            "portal": portal,
            "title": title,
            "link": e.get("originallink") or e.get("link", ""),
            "source": source or "네이버뉴스",
            "published": _parse_date(e.get("published", "")),
            "summary": summary,
        })
    return rows


def search_google(keyword: str, limit: int = 50, extra: str = "") -> list[dict]:
    resp = requests.get(GOOGLE_RSS.format(q=quote(f"{keyword} {extra}".strip())), headers=UA, timeout=15)
    resp.raise_for_status()
    return _rows_from_feed(feedparser.parse(resp.content), "구글")[:limit]


def _parse_naver_date(text: str):
    now = pd.Timestamp.now()
    if m := re.match(r"(\d+)(분|시간|일|주) 전", text):
        unit = {"분": "min", "시간": "h", "일": "D", "주": "W"}[m.group(2)]
        return now - pd.Timedelta(int(m.group(1)), unit=unit)
    if m := re.match(r"(\d{4})\.(\d{1,2})\.(\d{1,2})", text):
        return pd.Timestamp(*map(int, m.groups()))
    return pd.NaT


def _search_naver_web(keyword: str, limit: int) -> list[dict]:
    """API 키가 없을 때: 네이버 뉴스 검색 결과 페이지를 파싱(최신순)."""
    rows, seen = [], set()
    for start in range(1, limit + 1, 10):
        resp = requests.get(NAVER_WEB.format(q=quote(keyword), start=start), headers=UA, timeout=15)
        resp.raise_for_status()
        fragment = "".join(c.get("html", "") for c in resp.json().get("collection", []))
        soup = BeautifulSoup(fragment, "html.parser")
        titles = soup.select('a[data-heatmap-target=".tit"]')
        if not titles:
            break
        for a in titles:
            link = a.get("href", "")
            if link in seen:
                continue
            seen.add(link)
            body = soup.select_one(f'a[data-heatmap-target=".body"][href="{link}"]')
            card, source_el = a, None
            for _ in range(12):  # 기사 카드 컨테이너까지 올라가며 언론사 탐색
                card = card.parent
                if card is None:
                    break
                if source_el := card.select_one(".sds-comps-profile-info-title-text"):
                    break
            dates = [s.get_text(strip=True) for s in card.select(".sds-comps-profile-info-subtext")] if card else []
            date = next((d for d in dates if re.match(r"\d+(분|시간|일|주) 전|\d{4}\.", d)), "")
            rows.append({
                "portal": "네이버",
                "title": a.get_text(strip=True).replace("새 창 열림", "").strip(),
                "link": link,
                "source": source_el.get_text(strip=True).replace("새 창 열림", "") if source_el else "네이버뉴스",
                "published": _parse_naver_date(date),
                "summary": body.get_text(" ", strip=True).replace("새 창 열림", "") if body else "",
            })
    return rows[:limit]


def search_naver(keyword: str, limit: int = 50, client_id: str = "", client_secret: str = "") -> list[dict]:
    if not (client_id and client_secret):
        return _search_naver_web(keyword, limit)
    resp = requests.get(
        NAVER_API.format(q=quote(keyword), n=min(limit, 100)),
        headers={"X-Naver-Client-Id": client_id, "X-Naver-Client-Secret": client_secret},
        timeout=15,
    )
    resp.raise_for_status()
    return _rows_from_feed(feedparser.parse(resp.content), "네이버")


def to_dataframe(rows: list[dict]) -> pd.DataFrame:
    df = pd.DataFrame(rows, columns=["portal", "title", "link", "source", "published", "summary"])
    df = df.drop_duplicates(subset="title")
    return df.sort_values("published", ascending=False, na_position="last").reset_index(drop=True)


def top_keywords(df: pd.DataFrame, exclude: str = "", n: int = 20) -> pd.DataFrame:
    """제목+요약에서 자주 등장한 키워드 Top N (검색어 자체는 제외)."""
    exclude_tokens = {t.lower() for t in TOKEN_RE.findall(exclude)}
    counter = Counter()
    for text in (df["title"] + " " + df["summary"].fillna("")):
        # 3글자 이상일 때만 조사 제거(가을→가 같은 오탐 방지), 기사당 1회만 카운트
        tokens = {JOSA_RE.sub("", t) if len(t) >= 3 else t for t in TOKEN_RE.findall(text)}
        counter.update(
            t for t in tokens
            if len(t) >= 2 and t.lower() not in STOPWORDS
            and not any(x in t.lower() for x in exclude_tokens)  # 투썸 → 투썸플레이스도 제외
        )
    return pd.DataFrame(counter.most_common(n), columns=["키워드", "기사 수"])


def daily_counts(df: pd.DataFrame, days: int = 14) -> pd.DataFrame:
    d = df.dropna(subset=["published"]).copy()
    if d.empty:
        return pd.DataFrame()
    d["date"] = d["published"].dt.normalize()
    d = d[d["date"] >= d["date"].max() - pd.Timedelta(days=days - 1)]
    return d.pivot_table(index="date", columns="portal", values="title", aggfunc="count", fill_value=0)
