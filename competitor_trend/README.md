# ☕ 경쟁사 동향 모니터

투썸플레이스 · 스타벅스 · 메가MGC커피의 **신제품 동향**과 **구글·네이버 뉴스 트렌드**를 한 화면에서 확인하는 Streamlit 도구입니다.

## 두 가지 사용 방법

1. **설치 없이 웹에서 보기** — <https://hani-4455.github.io/S.HAN/competitor_trend.html>
   GitHub Actions가 매일 오전 8시에 `collect.py`를 실행해 `data/competitor_data.json`을 갱신하고,
   정적 페이지가 그 JSON을 읽어 카드로 보여준다. 즉시 갱신하려면 저장소 Actions 탭에서
   `경쟁사 동향 수집` 워크플로를 수동 실행한다.
2. **Streamlit 앱으로 직접 실행** — 아래 설치·실행 방법 참고. 수집 시점을 직접 고르고 싶을 때 쓴다.

## 화면 구성

| 메뉴 | 내용 |
|---|---|
| 🏠 신제품 요약(홈) | 세 브랜드 신제품을 한 번에 수집해 카드로 모아 보기, 전체 CSV 다운로드 |
| 경쟁사 제품 동향 | 브랜드별 전체 메뉴를 섹션(카테고리)별로 조회 — 제품명 · 제품 이미지 · 제품 설명 |
| 구글·네이버 뉴스 트렌드 | 키워드 뉴스 검색, 일자별 기사량 · 연관 키워드 · 언론사 분석, 카드 뉴스 |

## 설치 및 실행

```bash
pip install -r requirements.txt
streamlit run app.py
```

브라우저에서 <http://localhost:8501> 이 열립니다.

## 파일

| 파일 | 역할 |
|---|---|
| `app.py` | Streamlit 웹앱(화면 구성) |
| `twosome_crawler.py` | 투썸플레이스 모바일 메인 + 상세 페이지 크롤러 |
| `starbucks_crawler.py` | 스타벅스 카테고리 JSON 수집기 (NEW 아이콘·출시일 포함) |
| `mega_crawler.py` | 메가MGC커피 메뉴 목록 수집기 (신상품 카테고리 포함) |
| `news_trend.py` | 구글 뉴스 RSS · 네이버 뉴스 검색 + 트렌드 분석 |
| `collect.py` | 정적 페이지용 `data/competitor_data.json` 생성 (GitHub Actions가 매일 실행) |

각 크롤러는 단독 실행도 가능합니다. 예: `python starbucks_crawler.py`

## 참고

- 네이버는 뉴스 검색 RSS를 폐지해, **네이버 검색 Open API**(developers.naver.com에서 무료 발급)를 쓰는 것이 안정적입니다.
  키를 앱 화면에 입력하거나 환경변수 `NAVER_CLIENT_ID` / `NAVER_CLIENT_SECRET`로 지정하세요.
  키가 없으면 네이버 뉴스 검색 결과 페이지를 파싱해 대체 수집합니다(사이트 변경 시 영향을 받을 수 있음).
- 수집 결과는 30분간 캐시되며, 각 화면의 실행 버튼을 다시 누르면 새로 수집합니다.
- 경쟁사를 추가하려면 크롤러 파일을 하나 만들고 `app.py`의 `COMPETITORS` 목록에 한 줄만 추가하면 됩니다.
