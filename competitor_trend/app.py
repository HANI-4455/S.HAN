"""경쟁사 제품 동향 & 뉴스 트렌드 대시보드 (Streamlit).

실행: streamlit run app.py
"""
import html
import os

import pandas as pd
import streamlit as st

from news_trend import daily_counts, search_google, search_naver, to_dataframe, top_keywords
import mega_crawler
import starbucks_crawler
import twosome_crawler

st.set_page_config(page_title="경쟁사 동향 모니터", page_icon="☕", layout="wide")

st.markdown("""
<style>
.news-card {border:1px solid rgba(128,128,128,.25); border-radius:14px; padding:16px 18px;
            height:220px; display:flex; flex-direction:column; gap:8px; overflow:hidden;
            background:rgba(128,128,128,.04);}
.news-card:hover {border-color:#e2231a;}
.news-card .meta {font-size:12px; opacity:.7; display:flex; gap:8px; align-items:center;}
.news-card .badge {padding:2px 8px; border-radius:999px; font-weight:600; color:#fff; font-size:11px;}
.news-card .g {background:#4285f4;} .news-card .n {background:#03c75a;}
.news-card a.title {font-weight:700; font-size:15px; line-height:1.4; text-decoration:none; color:inherit;
            display:-webkit-box; -webkit-line-clamp:3; -webkit-box-orient:vertical; overflow:hidden;}
.news-card .summary {font-size:13px; opacity:.8; display:-webkit-box; -webkit-line-clamp:3;
            -webkit-box-orient:vertical; overflow:hidden;}
</style>
""", unsafe_allow_html=True)


# 경쟁사 등록: 이름 → (페이지 제목, 출처 URL, 크롤링 함수, 상세 수집 옵션 설명)
COMPETITORS = {
    "투썸": ("투썸플레이스", twosome_crawler.MAIN_URL, twosome_crawler.crawl_twosome,
            "상세 페이지의 제품 설명까지 수집"),
    "스타벅스": ("스타벅스", starbucks_crawler.MENU_URL, starbucks_crawler.crawl_starbucks,
              "제품 설명 텍스트 포함"),
    "메가커피": ("메가MGC커피", mega_crawler.MENU_URL, mega_crawler.crawl_mega,
              "제품 설명 텍스트 포함"),
}

# ---------------------------------------------------------------- 사이드바
with st.sidebar:
    st.title("☕ 경쟁사 동향 모니터")
    menu = st.radio("기능 선택", ["🏠 신제품 요약(홈)", "경쟁사 제품 동향", "구글·네이버 뉴스 트렌드"])
    if menu == "경쟁사 제품 동향":
        brand = st.selectbox("경쟁사", list(COMPETITORS))
    st.divider()
    st.caption("데이터는 30분간 캐시됩니다. 새로고침하려면 실행 버튼을 다시 누르세요.")


@st.cache_data(ttl=1800, show_spinner=False)
def cached_crawl(brand: str, with_detail: bool):
    return COMPETITORS[brand][2](with_detail)


@st.cache_data(ttl=1800, show_spinner=False)
def cached_news(keyword: str, portals: tuple, limit: int, nid: str, nsecret: str):
    rows = []
    if "구글" in portals:
        rows += search_google(keyword, limit)
    if "네이버" in portals:
        rows += search_naver(keyword, limit, nid, nsecret)
    return to_dataframe(rows)


# ---------------------------------------------------------------- 0. 홈(신제품 요약)
def page_home():
    st.header("🆕 경쟁사 신제품 한눈에 보기")
    st.caption("투썸 · 스타벅스 · 메가MGC커피의 신제품을 한 화면에 모아 봅니다.")

    if st.button("🔄 전체 새로고침", type="primary"):
        cached_crawl.clear()
        st.session_state.ran_home = True
    if not st.session_state.get("ran_home"):
        st.info("‘전체 새로고침’을 누르면 세 브랜드를 한 번에 수집합니다.")
        return

    frames, failed = {}, []
    progress = st.progress(0.0, text="수집 준비 중...")
    for i, brand in enumerate(COMPETITORS, start=1):
        progress.progress(i / len(COMPETITORS), text=f"{brand} 수집 중...")
        try:
            frames[brand] = pd.DataFrame(cached_crawl(brand, True))
        except Exception as e:  # 한 브랜드가 실패해도 나머지는 보여준다
            failed.append(f"{brand}: {e}")
    progress.empty()
    for msg in failed:
        st.error(f"크롤링 실패 — {msg}")
    if not frames:
        return

    cols = st.columns(len(frames))
    for col, (brand, df) in zip(cols, frames.items()):
        new_df = df[df["is_new"]].drop_duplicates("name")
        col.metric(f"{brand} 신제품", len(new_df), f"전체 {df.drop_duplicates('name').shape[0]}개")

    for brand, df in frames.items():
        new_df = df[df["is_new"]].drop_duplicates("name")
        st.subheader(f"{COMPETITORS[brand][0]}  ·  신제품 {len(new_df)}개")
        if new_df.empty:
            st.caption("표시할 신제품이 없습니다.")
            continue
        product_grid(new_df, with_detail=True)

    all_new = pd.concat(
        [df[df["is_new"]].drop_duplicates("name").assign(brand=brand) for brand, df in frames.items()]
    )
    st.download_button("전체 신제품 CSV 다운로드",
                       all_new.to_csv(index=False).encode("utf-8-sig"),
                       "competitor_new_products.csv", "text/csv")


# ---------------------------------------------------------------- 1. 경쟁사 제품 동향
def product_grid(sub: pd.DataFrame, with_detail: bool):
    cols = st.columns(4)
    for i, (_, p) in enumerate(sub.iterrows()):
        with cols[i % 4].container(border=True):
            if p["image"]:
                st.image(p["image"], width="stretch")
            badge = "🆕 " if p["is_new"] else ""
            st.markdown(f"**{badge}[{p['name']}]({p['link']})**")
            if p["extra"]:
                st.caption(p["extra"])
            if with_detail:
                if p["desc_text"]:
                    st.caption(p["desc_text"])
                if p["desc_images"]:
                    with st.expander("제품 설명 보기"):
                        for src in p["desc_images"]:
                            st.image(src, width="stretch")


def show_section(df: pd.DataFrame, section: str, with_detail: bool):
    sub = df[df["section"] == section]
    if sub.empty:
        st.write("검색 결과가 없습니다.")
    else:
        product_grid(sub, with_detail)


def page_competitor(brand: str):
    title, source_url, _, detail_label = COMPETITORS[brand]
    st.header(f"{title} 제품 동향")
    st.caption(f"출처: {source_url}")

    c1, c2 = st.columns([1, 3])
    with_detail = c2.checkbox(detail_label, value=True, key=f"detail_{brand}")
    if c1.button("🔍 웹크롤링 실행", type="primary", width="stretch", key=f"run_{brand}"):
        cached_crawl.clear(brand, with_detail)
        st.session_state[f"ran_{brand}"] = True

    if not st.session_state.get(f"ran_{brand}"):
        st.info("‘웹크롤링 실행’ 버튼을 눌러 최신 제품 정보를 가져오세요.")
        return

    try:
        with st.spinner(f"{title} 사이트에서 제품 정보를 수집하는 중..."):
            products = cached_crawl(brand, with_detail)
    except Exception as e:  # 네트워크/사이트 구조 변경
        st.error(f"크롤링 실패: {e}")
        return
    if not products:
        st.warning("수집된 제품이 없습니다. 사이트 구조가 바뀌었을 수 있습니다.")
        return

    df = pd.DataFrame(products)
    sections = list(dict.fromkeys(df["section"]))
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("수집 제품 수", df.drop_duplicates("name")["name"].size)
    m2.metric("신제품 수", df[df["is_new"]].drop_duplicates("name")["name"].size)
    m3.metric("섹션 수", len(sections))
    m4.metric("수집 시각", df["crawled_at"].iloc[0])

    keyword = st.text_input("제품명 검색", placeholder="예: 케이크, 라떼", key=f"kw_{brand}")
    if keyword:
        df = df[df["name"].str.contains(keyword, case=False, regex=False)]

    if len(sections) <= 8:
        for tab, section in zip(st.tabs(sections), sections):
            with tab:
                show_section(df, section, with_detail)
    else:  # 카테고리가 많으면(스타벅스) 드롭다운으로 선택
        counts = df["section"].value_counts()
        section = st.selectbox("섹션", sections, key=f"sec_{brand}",
                               format_func=lambda s: f"{s} ({counts.get(s, 0)})")
        show_section(df, section, with_detail)

    with st.expander("📋 표로 보기 / 다운로드"):
        table = df.copy()
        table["desc_images"] = table["desc_images"].apply(
            lambda v: v[0] if isinstance(v, list) and v else "")
        st.dataframe(
            table,
            column_config={
                "image": st.column_config.ImageColumn("제품 이미지", width="small"),
                "desc_images": st.column_config.LinkColumn("제품 설명 이미지"),
                "link": st.column_config.LinkColumn("상세 링크"),
            },
            hide_index=True, width="stretch",
        )
        st.download_button("CSV 다운로드", table.to_csv(index=False).encode("utf-8-sig"),
                           f"{brand}_products.csv", "text/csv", key=f"csv_{brand}")


# ---------------------------------------------------------------- 2. 뉴스 트렌드
def news_card(row) -> str:
    cls = "g" if row["portal"] == "구글" else "n"
    date = row["published"].strftime("%Y-%m-%d %H:%M") if pd.notna(row["published"]) else ""
    summary = f'<div class="summary">{html.escape(row["summary"])}</div>' if row["summary"] else ""
    return f"""
    <div class="news-card">
      <div class="meta"><span class="badge {cls}">{row['portal']}</span>
        <span>{html.escape(row['source'])}</span><span>· {date}</span></div>
      <a class="title" href="{html.escape(row['link'])}" target="_blank">{html.escape(row['title'])}</a>
      {summary}
    </div>"""


def page_news():
    st.header("구글·네이버 뉴스 트렌드")

    with st.form("news_form"):
        c1, c2, c3 = st.columns([3, 2, 1])
        keyword = c1.text_input("검색 키워드", value="투썸 신제품")
        portals = c2.multiselect("포털", ["구글", "네이버"], default=["구글", "네이버"])
        limit = c3.number_input("포털별 최대 건수", 10, 100, 50, step=10)
        with st.expander("네이버 검색 API 키 (선택)"):
            st.caption("없으면 네이버 뉴스 검색 페이지를 파싱해 대체 수집합니다. "
                       "키 발급: developers.naver.com → 애플리케이션 등록 → 검색 API")
            nid = st.text_input("Client ID", value=os.getenv("NAVER_CLIENT_ID", ""))
            nsecret = st.text_input("Client Secret", value=os.getenv("NAVER_CLIENT_SECRET", ""),
                                    type="password")
        submitted = st.form_submit_button("🔍 뉴스 검색", type="primary")

    if submitted:
        st.session_state.news_query = (keyword.strip(), tuple(portals), int(limit), nid, nsecret)
    query = st.session_state.get("news_query")
    if not query or not query[0] or not query[1]:
        st.info("키워드와 포털을 선택하고 ‘뉴스 검색’을 누르세요.")
        return

    try:
        with st.spinner("뉴스 RSS를 수집하는 중..."):
            df = cached_news(*query)
    except Exception as e:
        st.error(f"뉴스 수집 실패: {e}")
        return
    if df.empty:
        st.warning("검색 결과가 없습니다.")
        return

    # ---- 트렌드 분석
    st.subheader(f"📈 ‘{query[0]}’ 트렌드 분석")
    week_ago = pd.Timestamp.now() - pd.Timedelta(days=7)
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("전체 기사", len(df))
    m2.metric("최근 7일", int((df["published"] >= week_ago).sum()))
    m3.metric("언론사 수", df["source"].nunique())
    m4.metric("최신 기사", df["published"].max().strftime("%m-%d %H:%M") if df["published"].notna().any() else "-")

    a, b = st.columns(2)
    with a:
        st.markdown("**일자별 기사량 (최근 14일)**")
        daily = daily_counts(df)
        if daily.empty:
            st.caption("날짜 정보가 없습니다.")
        else:
            st.bar_chart(daily)
    with b:
        st.markdown("**연관 키워드 Top 20** (검색어 제외)")
        kw = top_keywords(df, exclude=query[0])
        st.bar_chart(kw.set_index("키워드"), horizontal=True)

    st.markdown("**주요 언론사 Top 10**")
    st.bar_chart(df["source"].value_counts().head(10))

    # ---- 카드 뉴스
    st.subheader("📰 카드 뉴스")
    f1, f2 = st.columns([2, 1])
    filt = f1.text_input("결과 내 검색", placeholder="제목에 포함된 단어")
    portal_filter = f2.multiselect("포털 필터", list(query[1]), default=list(query[1]))
    view = df[df["portal"].isin(portal_filter)]
    if filt:
        view = view[view["title"].str.contains(filt, case=False, regex=False)]

    per_page = 12
    pages = max(1, -(-len(view) // per_page))
    page = st.number_input(f"페이지 (총 {pages})", 1, pages, 1) if pages > 1 else 1
    chunk = view.iloc[(page - 1) * per_page: page * per_page]

    cols = st.columns(3)
    for i, (_, row) in enumerate(chunk.iterrows()):
        cols[i % 3].markdown(news_card(row), unsafe_allow_html=True)
        cols[i % 3].write("")

    st.download_button("CSV 다운로드", view.to_csv(index=False).encode("utf-8-sig"),
                       f"news_{query[0]}.csv", "text/csv")


if menu.startswith("🏠"):
    page_home()
elif menu == "경쟁사 제품 동향":
    page_competitor(brand)
else:
    page_news()
