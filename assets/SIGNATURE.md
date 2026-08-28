# 파리크라상 · 이성한 서명 직인

내가 제작한 모든 결과물에 공통으로 들어가는 서명(직인). **새로 만드는 HTML 도구는 예외 없이 이걸 넣는다.**

## 넣는 법

[`assets/pc-seal.html`](pc-seal.html) 파일 내용을 통째로 복사해서 **`<body>` 여는 태그 바로 뒤**에 붙인다. 그게 전부다 — 외부 파일 참조도, 폰트 로드도 필요 없다.

```html
<body>
<!-- pc-seal:v3 · 파리크라상 이성한 서명 -->
...  ← assets/pc-seal.html 내용 전체
<!-- /pc-seal:v3 -->

  <!-- 기존 페이지 내용 -->
```

하단이 아니라 **최상단**이다. 페이지를 열자마자 스크롤 없이 보여야 하기 때문.

## 규격

| 항목 | 값 |
|---|---|
| 글씨 | 손글씨 캘리그라피 "이성한" — SVG path (폰트 설치 불필요) |
| 색 | `#15237A` (Primary), 캡션 `#2F497D` (Secondary) |
| 크기 | 마크 높이 24px, 캡션 10.5px |
| 정렬 | 우측 상단, 아래 `1px` 구분선 |
| 문구 | `파리크라상 · 이성한` |
| 용량 | 7.6KB (인라인) |

색은 [디자인 가이드](../../.claude/CLAUDE.md)의 Primary/Secondary만 쓴다. 직인에 오렌지(`#F26300`)는 쓰지 않는다 — 오렌지는 섹션 번호와 액센트 라인 전용이다.

## 레이아웃 안전장치 (중요)

`body` 가 `display:flex(row)` 이거나 `display:grid` 인 페이지에서는, 직인이 흐름에 끼어들면서 **기존 콘텐츠를 옆으로 밀어낸다.** 실제로 `조직도.html` 에서 가운데 정렬 카드가 왼쪽으로 밀렸다.

그래서 스니펫 안에 작은 스크립트가 들어 있다. 실행 시점에 `body` 의 computed style 을 읽어서, flex(row)/grid 인 경우에만 `.pc-seal--float` 를 붙여 흐름에서 빼내고 우측 상단에 고정한다. **이 스크립트를 지우면 안 된다.**

상단에 헤더가 있는 페이지는 그 헤더가 `sticky` 인지 `fixed` 인지 확인할 것. `sticky` 는 로드 시점에 흐름에 있으므로 직인이 가려지지 않는다. `fixed` 라면 직인이 헤더 밑에 깔리므로 `.pc-seal--float` 를 강제로 붙이거나 헤더에 `margin-top` 을 줘야 한다.

## 파일

| 파일 | 용도 |
|---|---|
| `pc-seal.html` | **HTML 페이지에 붙여넣는 스니펫** (이것만 쓰면 됨) |
| `signature_mark.svg` | 캘리그라피 단독. `fill:currentColor` 라 색 상속됨 |
| `signature_compact.svg` | 마크 + 세로선 + 파리크라상 / PARIS CROISSANT 1줄 |
| `signature_full.svg` | 풀 락업 — 마크 + 오렌지·블루 구분선 + 캡션 2줄 |
| `signature_source.png` | 원본 손글씨 이미지 (재생성용 마스터) |

PPT·문서·인쇄물에는 `signature_full.svg`, 좁은 자리에는 `signature_compact.svg` 를 쓴다.

## 다시 만들어야 할 때

색이나 굵기만 바꾼다면 SVG 의 `fill` 값만 고치면 된다. 글씨 자체를 다시 뽑아야 하면 `signature_source.png` 에서 시작한다.

```bash
pip install potracer pillow numpy
```

1. PNG 을 그레이스케일로 읽어 `< 165` 인 픽셀을 먹(ink)으로 이진화하고, 잉크 영역으로 크롭
2. `potrace.Bitmap` 에 **반전해서** 넘긴다 — `potracer` 의 `Bitmap.__init__` 이 내부에서 `invert()` 를 호출하므로, 그냥 넣으면 배경이 추적돼 흑백이 뒤집힌다
3. `trace(turdsize=6, alphamax=1.0, opticurve=True, opttolerance=1.2)` — `opttolerance` 가 품질/용량을 가른다. `0.6` 은 11.9KB 로 붓끝 굵기 변화까지 살고, `1.2` 는 6.4KB 로 24px 표시에선 육안 차이가 없다. 인라인 삽입용은 `1.2` 를 쓴다
4. 높이 1000 기준으로 정규화하고 `fill-rule="evenodd"` 로 출력 (획 안쪽 구멍 처리)

## 이력

- **v3** — 위치를 하단에서 최상단으로 이동
- **v2** — flex/grid 페이지용 `--float` 안전장치 추가
- **v1** — 최초 적용 (하단 푸터)

버전을 올릴 땐 `pc-seal:vN` 주석 마커도 같이 올린다. 기존 페이지 일괄 갱신은 이 마커를 기준으로 잡아서 지우고 새로 넣으면 된다.
