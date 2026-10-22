/**
 * ============================================================================
 *  직영점 폐점 분석 자동화 시스템 (Google Apps Script)
 * ============================================================================
 *
 *  [설치 방법 - 비개발자용]
 *   1. 새 Google Sheets 파일을 엽니다.
 *   2. 상단 메뉴 [확장 프로그램 → Apps Script]를 클릭합니다.
 *   3. 기본으로 있는 코드(function myFunction...)를 모두 지우고,
 *      이 파일의 내용 전체를 그대로 붙여넣은 뒤 💾(저장) 버튼을 누릅니다.
 *      (대시보드·챗봇 화면도 이 파일 안에 들어 있으므로 HTML 파일을 따로 만들 필요가 없습니다)
 *   5. Sheets 화면으로 돌아와 새로고침(F5)하면 상단에 [🏪 폐점분석] 메뉴가 생깁니다.
 *   6. [🏪 폐점분석 → 🛠 시트 만들기 (샘플 데이터 포함)]을 처음 한 번 실행합니다.
 *      (최초 실행 시 Google 권한 승인 창이 뜨면 [허용]을 눌러 주세요.)
 *   7. 챗봇을 쓰려면 [🔑 챗봇 API 키 설정 (Gemini)]에서 Gemini API 키를 한 번 등록합니다.
 *
 *  [시트 구성]
 *   - 손익현황   : (입력) 브랜드·점포별 월별 매출/비용 → 영업이익 자동 계산
 *   - 누적손익   : (자동) 1단계 결과 - 점포별 누적 손익 요약
 *   - 자산투자비 : (입력+자동) 2단계 - 투자비·임차 정보 입력 → 감가상각/잔여임차료 자동 계산
 *   - 폐점이력   : (입력+자동) 3단계 - 과거 폐점 원복비 실적 (기존 '실적입력' 시트와 같은 열 순서)
 *   - 폐점기준값 : (자동) 3단계 결과 - 입지 유형 × 철거 범위별 평당 원복비 (적게/보통/많이)
 *   - 폐점검토   : (입력+자동) 3단계 - 폐점 검토 점포의 조건 입력 → 원복비 참고 범위 자동 계산
 *   - 폐점비용실적 : (입력+자동) 3단계 - 실제 폐점 정산 내역 (기존 B&S 파일 [최종금액] 시트와 같은 열 순서)
 *   - 브랜드폐점실적 : (자동) 3단계 결과 - 브랜드별 건당 평균 원상복구비·매몰비
 *   - 폐점요약   : (자동) 최종 결과 - 점포별 예상 총손실 요약 (AI 챗봇 연동용 표)
 *   - 브랜드요약 : (자동) 최종 결과 - 브랜드별 합계
 *   - 폐점시점분석 : (자동) 4단계 - 점포별 최적 폐점 시점 (영업이익 추세 기반)
 *   - 대시보드·챗봇 : 메뉴 [📊 대시보드 열기] / [💬 챗봇 열기] (화면 HTML은 이 파일 맨 아래)
 *   - 설정       : 시트별 금액 단위, 계산 기준일, 최소 사례수 등
 *   - 사용안내   : 사용 방법 안내
 *
 *  [자동 갱신 방식]
 *   - 입력 시트(손익현황/자산투자비/폐점이력/폐점검토/설정)를 수정하면 모든 결과가 자동으로 다시 계산됩니다.
 *   - 감가상각·잔여임차료는 날짜가 지나면 바뀌므로, 메뉴의 [⏰ 매일 자동 갱신 켜기]를
 *     한 번 실행해 두면 매일 오전 6시에 자동으로 다시 계산됩니다.
 *
 *  [금액 단위]
 *   - 결과는 모두 "천원" 단위, 쉼표(#,##0) 형식으로 표시됩니다.
 *   - 입력 단위는 [설정] 시트에서 시트별로 고릅니다. (원 / 천원 / 만원)
 *     기본값: 손익현황·자산투자비 = 천원, 폐점이력·폐점검토 = 만원 (기존 원복비 실적 파일 기준),
 *             폐점비용실적 = 원 (기존 B&S 파일 기준)
 *
 *  [매몰비]
 *   - B&S 정산 내역의 매몰비 = 폐점 시 손실 처리되는 미상각 투자비(장부가)로 보고,
 *     예상 매몰비는 2단계에서 계산한 '잔존자산가치'를 사용합니다.
 * ============================================================================
 */


// ============================================================================
// 0. 공통 설정 (시트 이름, 열 구성, 표시 형식)
//    ※ 시트 이름을 바꾸고 싶으면 아래 따옴표 안의 이름만 바꾸면 됩니다.
// ============================================================================

// 시트 이름 모음
const SHEETS = {
  PL: '손익현황',        // [입력] 월별 손익
  PL_SUM: '누적손익',     // [자동] 1단계 결과
  ASSET: '자산투자비',    // [입력+자동] 2단계
  HIST: '폐점이력',       // [입력+자동] 3단계 과거 폐점 원복비 실적
  BENCH: '폐점기준값',    // [자동] 3단계 평당 원복비 기준표
  REVIEW: '폐점검토',     // [입력+자동] 3단계 폐점 검토 점포 조건
  BS: '폐점비용실적',      // [입력+자동] 3단계 실제 폐점 정산 내역 (B&S)
  BS_SUM: '브랜드폐점실적', // [자동] 3단계 브랜드별 실제 폐점비용 평균
  TIMING: '폐점시점분석',   // [자동] 4단계 점포별 최적 폐점 시점
  SUMMARY: '폐점요약',    // [자동] 최종 점포별 요약 (AI 챗봇 연동용)
  BRAND: '브랜드요약',    // [자동] 최종 브랜드별 합계
  BRAND_MAP: '브랜드매핑', // [입력+자동] 원본 시트 '소속' → 브랜드명 (원본 연동 시)
  CONFIG: '설정',         // 설정값
  GUIDE: '사용안내'       // 안내문
};

// 선택 목록 (회사 기준이 바뀌면 여기만 고치면 됩니다)
const LOCATION_TYPES = ['로드샵', '몰·백화점', '오피스·기타'];          // 입지 유형
const DEMOLITION_SCOPES = ['전체 철거', '부분 철거', '인수인계'];        // 철거 범위
const DATA_SOURCES = ['구매팀', 'SI팀', '인테리어실', '기타'];           // 자료 출처
const YN = ['Y', 'N'];

// 비용을 키우는 주의 항목 6개 (Y 개수로 '판단' 문구를 만듭니다)
const CAUTION_ITEMS = ['주방 배기 덕트', '구조 변경 복구', '설비 원복', '야간·제한 공사', '반출 어려움', '석면'];

// 표시 형식
const FMT_MONEY = '#,##0';                 // 금액: 1,234
const FMT_PL = '#,##0;[Red]-#,##0';        // 손익: 음수는 빨간색 -1,234
const FMT_MONTH = '0.0';                   // 개월 수: 12.5
const FMT_AREA = '0.0';                    // 평수: 30.5
const FMT_DATE = 'yyyy-mm-dd';             // 날짜: 2026-01-31
const FMT_PCT = '0.0%';                    // 비율: 12.3%
const FMT_TEXT = '@';                      // 일반 텍스트 (연월이 날짜로 바뀌지 않도록)

// 열 종류: IN = 사람이 입력하는 열(파란 제목), AUTO = 스크립트가 채우는 열(회색 제목)
const IN = 'in';
const AUTO = 'auto';
const COLOR_IN = '#DCE9F8';
const COLOR_AUTO = '#E7E7E7';
const COLOR_OUT = '#FDEBD0';

// [설정] 시트 항목 이름
const CFG = {
  UNIT_PL: '손익현황 금액 단위',
  UNIT_ASSET: '자산투자비 금액 단위',
  UNIT_HIST: '폐점이력·폐점검토 금액 단위',
  UNIT_BS: '폐점비용실적 금액 단위',
  BASE: '계산 기준일',
  MIN: '기준값 최소 사례수',
  LEVEL: '원복비 적용 수준',
  SCOPE: '폐점요약 대상',
  RENT_RATE: '잔여임차료 반영률(%)',
  TREND: '추세 분석 기간(개월)',
  DEP_IN_PL: '손익에 감가상각비 포함',
  SRC_URL: '원본 시트 주소',
  SRC_PL: '원본 손익 탭',
  SRC_INFO: '원본 점포정보 탭',
  SRC_CLOSE: '원본 폐점비 탭',
  PROFIT: '손익 기준 이익',
  LAST_IMPORT: '마지막 원본 가져오기',
  LAST: '마지막 갱신'
};
const UNITS = { '원': 0.001, '천원': 1, '만원': 10 };   // 천원으로 바꾸는 배수
const LEVELS = { '적게(최소)': 'low', '보통(중앙값)': 'mid', '많이(최대)': 'high' };
const SCOPE_ALL = '전체 점포';
const SCOPE_REVIEW = '폐점검토 점포만';

// ── 1단계: 손익현황(입력) 열 구성 ──
const PL_COLS = [
  { name: '브랜드명', type: IN, width: 100 },
  { name: '점포명', type: IN, width: 110 },
  { name: '연월', type: IN, fmt: FMT_TEXT, note: '예: 2026-01 (YYYY-MM 형식). 점포별로 한 달에 한 행씩 입력합니다.' },
  { name: '매출', type: IN, fmt: FMT_MONEY, note: '해당 월 매출액' },
  { name: '비용', type: IN, fmt: FMT_MONEY, note: '해당 월 영업비용 합계 (매출원가+인건비+임차료+판관비 등)' },
  { name: '영업이익', type: AUTO, fmt: FMT_PL, note: '매출 − 비용으로 자동 계산됩니다. (매출/비용이 비어 있으면 직접 입력한 값을 사용)' },
  { name: '비고', type: IN, width: 160 }
];

// ── 1단계: 누적손익(자동) 열 구성 ──
const PL_SUM_COLS = [
  { name: '브랜드명', width: 100 },
  { name: '점포명', width: 110 },
  { name: '시작월', fmt: FMT_TEXT },
  { name: '최근월', fmt: FMT_TEXT },
  { name: '운영개월수', fmt: '0' },
  { name: '누적매출(천원)', fmt: FMT_MONEY, width: 120 },
  { name: '누적비용(천원)', fmt: FMT_MONEY, width: 120 },
  { name: '누적영업이익(천원)', fmt: FMT_PL, width: 130 },
  { name: '영업이익률', fmt: FMT_PCT },
  { name: '최근3개월 평균영업이익(천원)', fmt: FMT_PL, width: 170 },
  { name: '최근월 영업이익(천원)', fmt: FMT_PL, width: 140 },
  { name: '데이터점검', width: 220 },
  { name: '갱신일시', width: 130 }
];

// ── 2단계: 자산투자비(입력+자동) 열 구성 ──
const ASSET_COLS = [
  { name: '브랜드명', type: IN, width: 100 },
  { name: '점포명', type: IN, width: 110 },
  { name: '평수', type: IN, fmt: FMT_AREA, note: '전용 면적(평). 원복비 = 평당 원복비 × 평수 계산에 사용합니다.' },
  { name: '투자시점', type: IN, fmt: FMT_DATE, note: '오픈(투자) 일자. 이 날짜부터 감가상각이 시작됩니다. 예: 2023-03-01' },
  { name: '인테리어비', type: IN, fmt: FMT_MONEY, note: '감가상각 대상' },
  { name: '집기비', type: IN, fmt: FMT_MONEY, note: '감가상각 대상' },
  { name: '기타투자비', type: IN, fmt: FMT_MONEY, note: '간판·POS·설계비 등 감가상각 대상 기타 투자비' },
  { name: '내용연수(년)', type: IN, fmt: '0', note: '감가상각 기간(년). 예: 5' },
  { name: '보증금', type: IN, fmt: FMT_MONEY, note: '임차보증금. 돌려받는 돈이므로 감가상각·손실 계산에서 제외하고 참고용으로만 표시합니다.' },
  { name: '임차개시일', type: IN, fmt: FMT_DATE },
  { name: '임차종료일', type: IN, fmt: FMT_DATE, note: '임대차 계약 만료일. 이 날짜까지 남은 임차료를 계산합니다.' },
  { name: '월임차료', type: IN, fmt: FMT_MONEY, note: '월 임대료. 관리비도 부담해야 한다면 합산하여 입력하세요.' },
  { name: '계산기준일', type: AUTO, fmt: FMT_TEXT, note: '[폐점검토]의 폐점예정일 → [설정]의 계산 기준일 → 오늘 순서로 적용' },
  { name: '경과개월', type: AUTO, fmt: FMT_MONTH, note: '투자시점부터 계산기준일까지 지난 개월 수' },
  { name: '감가상각누계액(천원)', type: AUTO, fmt: FMT_MONEY, width: 140, note: '정액법: (인테리어+집기+기타) × 경과개월 ÷ (내용연수×12)' },
  { name: '잔존자산가치(천원)', type: AUTO, fmt: FMT_MONEY, width: 130, note: '(인테리어+집기+기타) − 감가상각누계액. 폐점 시 손실 처리되는 금액' },
  { name: '잔여임차개월', type: AUTO, fmt: FMT_MONTH, note: '계산기준일부터 임차종료일까지 남은 개월 수' },
  { name: '잔여임차료(천원)', type: AUTO, fmt: FMT_MONEY, width: 130, note: '월임차료 × 잔여임차개월' },
  { name: '데이터점검', type: AUTO, width: 220 },
  { name: 'POS코드', type: IN, width: 80, note: '원본 시트 연동 시 자동으로 채워집니다. 원본과 점포를 잇는 기준' }
];

// ── 3단계: 폐점이력(입력+자동) 열 구성 ──
//    기존 '폐점 원복비 실적 입력' 파일의 [실적입력] 시트와 열 순서가 같으므로
//    그 시트의 A5:S 범위를 그대로 복사해 2행에 붙여넣으면 됩니다.
const HIST_COLS = [
  { name: 'No', type: IN, width: 45 },
  { name: '브랜드', type: IN, width: 110 },
  { name: '점포명', type: IN, width: 110 },
  { name: '입지 유형', type: IN, list: LOCATION_TYPES },
  { name: '면적(평)', type: IN, fmt: FMT_AREA, note: '전용 면적. 평당 원복비 계산에 꼭 필요합니다.' },
  { name: '폐점 연월', type: IN, fmt: FMT_TEXT, note: '예: 2025-04' },
  { name: '철거 범위', type: IN, list: DEMOLITION_SCOPES },
  { name: '주방 배기 덕트', type: IN, list: YN, width: 80 },
  { name: '구조 변경 복구', type: IN, list: YN, width: 80 },
  { name: '설비 원복', type: IN, list: YN, width: 80 },
  { name: '야간·제한 공사', type: IN, list: YN, width: 80 },
  { name: '반출 어려움', type: IN, list: YN, width: 80 },
  { name: '석면', type: IN, list: YN, width: 60 },
  { name: '원복 공사비', type: IN, fmt: FMT_MONEY, note: '원상복구 공사비 실적 (단위는 [설정]의 폐점이력 금액 단위, 기본 만원)' },
  { name: '기타 폐점비', type: IN, fmt: FMT_MONEY, note: '인력 정리·재고 폐기·간판 등 원복 공사 외 비용 (기본 만원)' },
  { name: '평당 원복비(천원)', type: AUTO, fmt: FMT_MONEY, width: 120, note: '원복 공사비 ÷ 면적 (천원 환산)' },
  { name: '주의 항목 수', type: AUTO, fmt: '0', width: 80, note: '주의 항목 6개 중 Y 개수' },
  { name: '자료 출처', type: IN, list: DATA_SOURCES },
  { name: '비고', type: IN, width: 180 }
];

// ── 3단계: 폐점기준값(자동) 열 구성 ──
const BENCH_COLS = [
  { name: '입지 유형', width: 100 },
  { name: '철거 범위', width: 90 },
  { name: '사례 수', fmt: '0', width: 70 },
  { name: '평당 원복비 적게(최소,천원)', fmt: FMT_MONEY, width: 170 },
  { name: '평당 원복비 보통(중앙값,천원)', fmt: FMT_MONEY, width: 180 },
  { name: '평당 원복비 많이(최대,천원)', fmt: FMT_MONEY, width: 170 },
  { name: '평당 원복비 평균(천원)', fmt: FMT_MONEY, width: 150 },
  { name: '기타 폐점비 평균(천원)', fmt: FMT_MONEY, width: 150 },
  { name: '신뢰도', width: 90 },
  { name: '갱신일시', width: 130 }
];

// ── 3단계: 폐점검토(입력+자동) 열 구성 ──
//    기존 파일의 [참고범위] 계산기를 여러 점포에 한 번에 적용하는 시트입니다.
//    이 시트에 한 줄이라도 있는 점포가 '폐점검토 점포'가 됩니다.
const REVIEW_COLS = [
  { name: '브랜드명', type: IN, width: 100 },
  { name: '점포명', type: IN, width: 110 },
  { name: '폐점예정일', type: IN, fmt: FMT_DATE, note: '입력하면 감가상각·잔여임차료를 이 날짜 기준으로 계산합니다.' },
  { name: '입지 유형', type: IN, list: LOCATION_TYPES },
  { name: '철거 범위', type: IN, list: DEMOLITION_SCOPES },
  { name: '주방 배기 덕트', type: IN, list: YN, width: 80, note: '옥상·외벽까지 이어진 배기 덕트가 있나요?' },
  { name: '구조 변경 복구', type: IN, list: YN, width: 80, note: '바닥 타공·벽체 변경·주방 방수를 되돌려야 하나요?' },
  { name: '설비 원복', type: IN, list: YN, width: 80, note: '전기 증설·급배수·그리스트랩·가스·소방을 원래대로 돌려야 하나요?' },
  { name: '야간·제한 공사', type: IN, list: YN, width: 80, note: '영업시간 외에만 공사할 수 있나요?' },
  { name: '반출 어려움', type: IN, list: YN, width: 80, note: '지하·고층이라 자재 반출이 까다롭나요?' },
  { name: '석면', type: IN, list: YN, width: 60, note: '노후 건물이라 석면 조사·해체가 필요한가요?' },
  { name: '시설 인수 협의', type: IN, list: YN, width: 80, note: '다음 임차인이 시설을 인수하나요? (Y면 인수인계 기준으로 계산)' },
  { name: '원상복구 범위 확인', type: IN, list: YN, width: 90, note: '계약서의 원상복구 조항과 건물주 요구를 확인했나요?' },
  { name: '현장 실사 완료', type: IN, list: YN, width: 80, note: '인테리어실·SI팀이 현장을 봤나요?' },
  { name: '원복비_직접입력', type: IN, fmt: FMT_MONEY, note: '확정 견적이 있으면 입력 (기본 만원). 비우면 폐점기준값으로 자동 계산' },
  { name: '기타폐점비_직접입력', type: IN, fmt: FMT_MONEY, note: '확정 금액이 있으면 입력 (기본 만원). 비우면 과거 평균 적용' },
  { name: '위약금_직접입력', type: IN, fmt: FMT_MONEY, note: '중도 해지 위약금이 있으면 입력 (기본 만원). 예상총손실액에 더해집니다.' },
  { name: '평수', type: AUTO, fmt: FMT_AREA, note: '[자산투자비]의 평수를 가져옵니다.' },
  { name: '적용 유형', type: AUTO, width: 90, note: '시설 인수 협의가 Y면 인수인계 기준' },
  { name: '비슷한 매장 수', type: AUTO, fmt: '0', width: 80 },
  { name: '평당 원복비 보통(천원)', type: AUTO, fmt: FMT_MONEY, width: 140 },
  { name: '원복비 적게(천원)', type: AUTO, fmt: FMT_MONEY, width: 120 },
  { name: '원복비 보통(천원)', type: AUTO, fmt: FMT_MONEY, width: 120 },
  { name: '원복비 많이(천원)', type: AUTO, fmt: FMT_MONEY, width: 120 },
  { name: '기타 폐점비(천원)', type: AUTO, fmt: FMT_MONEY, width: 120 },
  { name: '주의 항목 수', type: AUTO, fmt: '0', width: 80 },
  { name: '판단', type: AUTO, width: 300 },
  { name: '확정도 등급', type: AUTO, width: 300 },
  { name: '데이터점검', type: AUTO, width: 260 }
];

// ── 3단계: 폐점비용실적(입력+자동) 열 구성 ──
//    기존 B&S 파일 [최종금액] 시트의 B~K열과 순서가 같으므로
//    그 시트의 B5:K(마지막 행)를 그대로 복사해 2행에 붙여넣으면 됩니다.
//    소급(정산 차액) 행은 같은 점포코드끼리 자동으로 합산됩니다.
const BS_COLS = [
  { name: '년', type: IN, fmt: '0', width: 55 },
  { name: '월', type: IN, fmt: '0', width: 45 },
  { name: '브랜드', type: IN, width: 80, note: '다른 시트의 브랜드명과 같은 표기를 쓰세요. (예: PB, PAS, PC)' },
  { name: '점포코드', type: IN, width: 75, note: '같은 점포의 소급 행을 묶는 기준입니다. 비우면 브랜드+점포명으로 묶습니다.' },
  { name: '점포명', type: IN, width: 200 },
  { name: '원상복구(예가)', type: IN, fmt: FMT_PL, width: 120, note: '원상복구 비용. 소급 정산 차액은 음수로 입력' },
  { name: '매몰비', type: IN, fmt: FMT_PL, width: 120, note: '폐점 시 손실 처리된 미상각 투자비(장부가)' },
  { name: '합계', type: AUTO, fmt: FMT_PL, width: 120, note: '원상복구 + 매몰비 + 위약금 (자동 계산)' },
  { name: '정산', type: IN, list: ['○', '△'], width: 50, note: '○ : 정산완료, △ : 부분정산, 빈칸 : 미정산' },
  { name: '비고', type: IN, width: 140 },
  { name: '위약금', type: IN, fmt: FMT_PL, width: 100, note: '중도 해지 위약금이 있으면 입력 (선택)' }
];

// ── 3단계: 브랜드폐점실적(자동) 열 구성 ──
const BS_SUM_COLS = [
  { name: '브랜드', width: 90 },
  { name: '폐점 점포수', fmt: '0', width: 80 },
  { name: '원상복구 합계(천원)', fmt: FMT_PL, width: 130 },
  { name: '매몰비 합계(천원)', fmt: FMT_PL, width: 130 },
  { name: '위약금 합계(천원)', fmt: FMT_PL, width: 120 },
  { name: '폐점비용 합계(천원)', fmt: FMT_PL, width: 130 },
  { name: '건당 평균 원상복구(천원)', fmt: FMT_PL, width: 150 },
  { name: '건당 평균 매몰비(천원)', fmt: FMT_PL, width: 150 },
  { name: '건당 평균 폐점비용(천원)', fmt: FMT_PL, width: 150 },
  { name: '매몰비 비중', fmt: FMT_PCT, width: 85 },
  { name: '정산완료 점포수', fmt: '0', width: 100 },
  { name: '부분·미정산 점포수', fmt: '0', width: 110 },
  { name: '신뢰도', width: 90 },
  { name: '갱신일시', width: 130 }
];

// ── 최종: 폐점요약(자동) 열 구성 ── (앞쪽 7개가 요청된 핵심 컬럼)
const SUMMARY_COLS = [
  { name: '브랜드명', width: 100 },
  { name: '점포명', width: 110 },
  { name: '누적손익(천원)', fmt: FMT_PL, width: 120 },
  { name: '잔여임차료(천원)', fmt: FMT_MONEY, width: 120 },
  { name: '예상폐점비용(원복+매몰)(천원)', fmt: FMT_MONEY, width: 180, note: '예상원복비 + 매몰비(=잔존자산가치)' },
  { name: '잔존자산가치(천원)', fmt: FMT_MONEY, width: 130, note: '미상각 투자비 = 예상 매몰비' },
  { name: '예상총손실액(천원)', fmt: FMT_MONEY, width: 130, note: '잔여임차료×반영률([설정], 기본 100%) + 예상폐점비용(원복+매몰) + 기타 폐점비 + 위약금' },
  { name: '폐점 후 최종손익(천원)', fmt: FMT_PL, width: 150 },
  { name: '폐점검토', width: 70 },
  { name: '평수', fmt: FMT_AREA },
  { name: '운영개월수', fmt: '0' },
  { name: '누적매출(천원)', fmt: FMT_MONEY, width: 120 },
  { name: '최근3개월 평균영업이익(천원)', fmt: FMT_PL, width: 170 },
  { name: '잔여임차개월', fmt: FMT_MONTH },
  { name: '예상원복비(천원)', fmt: FMT_MONEY, width: 120 },
  { name: '기타 폐점비(천원)', fmt: FMT_MONEY, width: 120 },
  { name: '위약금(천원)', fmt: FMT_MONEY, width: 100 },
  { name: '예상폐점비용 적게(천원)', fmt: FMT_MONEY, width: 150 },
  { name: '예상폐점비용 많이(천원)', fmt: FMT_MONEY, width: 150 },
  { name: '브랜드 과거 건당 원상복구(천원)', fmt: FMT_PL, width: 190 },
  { name: '브랜드 과거 건당 매몰비(천원)', fmt: FMT_PL, width: 180 },
  { name: '원복비 적용기준', width: 280 },
  { name: '판단', width: 280 },
  { name: '확정도 등급', width: 280 },
  { name: '보증금(천원)', fmt: FMT_MONEY, width: 110 },
  { name: '계산기준일', fmt: FMT_TEXT },
  { name: '데이터점검', width: 240 },
  { name: '갱신일시', width: 130 }
];

// ── 최종: 브랜드요약(자동) 열 구성 ──
const BRAND_COLS = [
  { name: '브랜드명', width: 100 },
  { name: '점포수', fmt: '0' },
  { name: '폐점검토 점포수', fmt: '0', width: 110 },
  { name: '누적적자 점포수', fmt: '0', width: 110 },
  { name: '누적매출(천원)', fmt: FMT_MONEY, width: 120 },
  { name: '누적손익(천원)', fmt: FMT_PL, width: 120 },
  { name: '잔여임차료(천원)', fmt: FMT_MONEY, width: 120 },
  { name: '예상폐점비용(원복+매몰)(천원)', fmt: FMT_MONEY, width: 180 },
  { name: '잔존자산가치(천원)', fmt: FMT_MONEY, width: 130 },
  { name: '예상총손실액(천원)', fmt: FMT_MONEY, width: 130 },
  { name: '폐점 후 최종손익(천원)', fmt: FMT_PL, width: 150 },
  { name: '갱신일시', width: 130 }
];


// ============================================================================
// 메뉴 및 자동 실행(트리거)
// ============================================================================

// 파일을 열 때 상단에 [🏪 폐점분석] 메뉴를 만들어 줍니다.
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏪 폐점분석')
    .addItem('🛠 시트 만들기 (샘플 데이터 포함)', 'setupWithSample')
    .addItem('🛠 시트 만들기 (빈 양식)', 'setupBlank')
    .addSeparator()
    .addItem('🔄 원본 가져오기 + 전체 다시 계산', 'updateAll')
    .addSeparator()
    .addItem('📊 대시보드 열기', 'openDashboard')
    .addItem('🔗 대시보드 새 창으로 열기 (웹앱 링크)', 'openDashboardLink')
    .addItem('💬 챗봇 열기 (사이드바)', 'openChatSidebar')
    .addItem('🔑 챗봇 API 키 설정 (Gemini)', 'setChatApiKey')
    .addSeparator()
    .addItem('⏰ 매일 자동 갱신 켜기 (오전 6시)', 'installDailyTrigger')
    .addItem('⏰ 매일 자동 갱신 끄기', 'removeDailyTrigger')
    .addToUi();
}

// 입력 시트의 셀을 수정할 때마다 자동으로 실행됩니다. (별도 설정 불필요)
function onEdit(e) {
  if (!e || !e.range) return;
  const sheetName = e.range.getSheet().getName();
  const watched = [SHEETS.PL, SHEETS.ASSET, SHEETS.HIST, SHEETS.REVIEW, SHEETS.BS, SHEETS.CONFIG];
  if (watched.indexOf(sheetName) === -1) return;  // 결과 시트 수정은 무시

  try {
    runAll_();
    e.source.toast('폐점요약이 자동 갱신되었습니다.', '🏪 폐점분석', 3);
  } catch (err) {
    e.source.toast('자동 갱신 중 오류: ' + err.message, '🏪 폐점분석', 10);
  }
}

// [메뉴] 전체 다시 계산
function updateAll() {
  try {
    const r = runAll_({ importSource: true });
    SpreadsheetApp.getActiveSpreadsheet().toast(
      (r.imported ? r.imported + '\n' : '') +
      '손익 ' + r.plCount + '개 점포, 자산 ' + r.assetCount + '개 점포, 폐점이력 ' + r.histCount +
      '건, 폐점검토 ' + r.reviewCount + '개 점포를 반영하여 폐점요약을 갱신했습니다.', '🏪 폐점분석', 10);
  } catch (err) {
    showAlert_('계산 중 오류가 발생했습니다.\n\n' + err.message);
  }
}

// 매일 자동 갱신(시간 트리거)에서 호출되는 함수
function scheduledUpdate() {
  runAll_({ importSource: true });
}

// [메뉴] 매일 오전 6시 자동 갱신 켜기
function installDailyTrigger() {
  removeTriggers_('scheduledUpdate');  // 중복 방지를 위해 기존 것 삭제 후 새로 등록
  ScriptApp.newTrigger('scheduledUpdate').timeBased().everyDays(1).atHour(6).create();
  showAlert_('매일 오전 6시에 자동으로 다시 계산되도록 설정했습니다.');
}

// [메뉴] 매일 자동 갱신 끄기
function removeDailyTrigger() {
  const n = removeTriggers_('scheduledUpdate');
  showAlert_(n > 0 ? '매일 자동 갱신을 껐습니다.' : '켜져 있는 자동 갱신이 없습니다.');
}

function removeTriggers_(handlerName) {
  let count = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === handlerName) {
      ScriptApp.deleteTrigger(t);
      count++;
    }
  });
  return count;
}


// ============================================================================
// 전체 실행 순서: 1단계 → 2단계 → 3단계(기준값 → 검토점포) → 최종 요약
// ============================================================================

//   opts.importSource = true 이면 계산 전에 원본 시트에서 데이터를 먼저 가져옵니다. (0단계)
function runAll_(opts) {
  opts = opts || {};
  // 여러 사람이 동시에 입력해도 계산이 겹치지 않도록 잠금을 겁니다.
  let lock = null;
  try {
    lock = LockService.getDocumentLock();
    lock.waitLock(20000);
  } catch (err) {
    lock = null;  // 잠금을 못 얻으면 그냥 진행
  }

  try {
    const imported = opts.importSource ? step0_importSource_(makeContext_()) : null;   // 0단계: 원본 연동
    const r = computeAll_(false);
    step4_writeTimingSheet_(r);                               // 4단계: 점포별 최적 폐점 시점 표
    writeConfigValue_(r.ctx.ss, CFG.LAST, r.ctx.now);
    return { plCount: r.pl.size, assetCount: r.assets.list.length, histCount: r.bench.all.values.length,
      reviewCount: r.review.size, imported: imported };
  } finally {
    if (lock) lock.releaseLock();
  }
}

// 1~3단계 계산을 한 번에 실행합니다.
//   dryRun = true 이면 시트에 쓰지 않고 계산 결과만 돌려줍니다. (대시보드·챗봇용)
let DRY_RUN_ = false;
function computeAll_(dryRun) {
  DRY_RUN_ = !!dryRun;
  try {
    const ctx = makeContext_();
    const pl = step1_updateProfitLoss_(ctx);                  // 1단계: 누적 손익
    const assets = step2_updateAssets_(ctx);                  // 2단계: 감가상각·잔여임차료
    const bench = step3_updateBenchmark_(ctx);                // 3단계: 과거 실적 → 평당 원복비 기준값
    bench.actuals = step3_updateActuals_(ctx);                // 3단계: B&S 정산 내역 → 브랜드별 건당 평균
    const review = step3_updateReview_(ctx, assets, bench);   // 3단계: 폐점검토 점포별 원복비 범위
    const items = step3_buildSummary_(ctx, pl, assets, bench, review);  // 3단계: 최종 폐점요약 생성
    return { ctx: ctx, pl: pl, assets: assets, bench: bench, review: review, items: items };
  } finally {
    DRY_RUN_ = false;
  }
}

// 이 스크립트가 연결된 스프레드시트 (웹앱으로 열 때도 같은 파일을 찾도록 ID를 기억해 둡니다)
function getSpreadsheet_() {
  const props = PropertiesService.getScriptProperties();
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) {
    if (props.getProperty('SPREADSHEET_ID') !== active.getId()) props.setProperty('SPREADSHEET_ID', active.getId());
    return active;
  }
  const id = props.getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('스프레드시트를 찾을 수 없습니다. 시트에서 [🔄 원본 가져오기 + 전체 다시 계산]을 한 번 실행해 주세요.');
  return SpreadsheetApp.openById(id);
}

// [편집기에서 실행] 이 스크립트가 연결된 원본 시트 주소를 '실행 로그'에 보여줍니다.
//   시트를 못 찾을 때: Apps Script 편집기 위쪽 함수 목록에서 showSheetUrl 선택 → ▶ 실행
function showSheetUrl() {
  const ss = getSpreadsheet_();
  Logger.log('원본 시트: ' + ss.getName());
  Logger.log(ss.getUrl());
}

// 계산에 필요한 공통 정보(시간대, 설정값, 기준일)를 한 번에 모읍니다.
function makeContext_() {
  const ss = getSpreadsheet_();
  const tz = ss.getSpreadsheetTimeZone();
  const cfg = readConfig_(ss, tz);
  const today = toParts_(new Date(), tz);
  return {
    ss: ss,
    tz: tz,
    cfg: cfg,
    baseDate: cfg.baseDate || today,                 // 설정 기준일이 없으면 오늘
    now: Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm')
  };
}


// ============================================================================
// 0단계. 원본 시트(직영점 데이터) 연동
//   - [설정]의 '원본 시트 주소'에 직영점 데이터 구글 시트 주소를 넣으면,
//     [🔄 원본 가져오기 + 전체 다시 계산] 메뉴와 매일 자동 갱신 때 아래를 자동으로 채웁니다.
//       · 원본 [시트3]   월별 손익 → [손익현황]  (폐점한 점포는 빼고 가져옴)
//       · 원본 [점포정보] 점포 마스터 → [자산투자비]의 평수·투자시점(오픈일)·POS코드 (빈 칸만 채움)
//       · 원본 [폐점비]   폐점 비용 → [폐점비용실적]
//   - 원본 시트는 읽기만 하고 고치지 않습니다.
//   - 투자비(인테리어·집기)·보증금·임차종료일·월임차료는 원본에 없으므로 [자산투자비]에 직접 입력합니다.
//   - 셀을 고칠 때 자동 실행되는 계산(onEdit)은 다른 파일을 열 권한이 없어 원본을 가져오지 않습니다.
// ============================================================================

const SOURCE_PL_UNIT = 0.001;   // 원본 손익(시트3)은 원 단위 → 천원으로 바꾸는 배수

// 브랜드 코드 추정용 단서 (점포명 앞머리나 소속 이름에 이 글자가 있으면 그 브랜드로 봄)
const BRAND_HINTS = [
  ['PAS', ['PAS', '파스쿠찌']],
  ['PB', ['PB', '파리바게뜨']],
  ['PC', ['PC', '파리크라상']],
  ['CW', ['C@W', 'CW', '커피앳웍스']],
  ['LAG', ['LAG', '라그릴리아']],
  ['STR', ['STR', '스트릿']],
  ['LI', ['리나스']]
];

function step0_importSource_(ctx) {
  const cfg = ctx.cfg;
  const id = sourceId_(cfg.sourceUrl);
  if (!id) return null;                       // 원본 시트 주소가 없으면 연동하지 않음
  const src = SpreadsheetApp.openById(id);
  const msgs = [];

  // (1) 폐점비 탭: 폐점한 점포 목록 + [폐점비용실적] 채우기
  const closedPos = {};
  if (cfg.sourceClosingTab) {
    const t = readSourceTable_(src, cfg.sourceClosingTab, ['년', '월', 'POS코드', '점포명', '원상복구', '매몰비']);
    const toCheon = unitFactor_(t.unitText, 1);   // 제목 위 '(단위 : 천)' 표시를 읽어 천원으로 환산
    const rows = t.rows.filter(function (r) { return num_(t.get(r, '년')) && str_(t.get(r, '점포명')); });
    rows.forEach(function (r) { closedPos[str_(t.get(r, 'POS코드'))] = true; });
    writeClosingActuals_(ctx, rows.map(function (r) {
      return { y: t.get(r, '년'), m: t.get(r, '월'), brand: str_(t.get(r, '브랜드')), code: str_(t.get(r, 'POS코드')),
        name: str_(t.get(r, '점포명')), restore: num0_(t.get(r, '원상복구')) * toCheon, sunk: num0_(t.get(r, '매몰비')) * toCheon };
    }));
    msgs.push('폐점비 ' + rows.length + '건');
  }

  // (2) 점포정보 탭: POS코드 → 평수·오픈일
  const info = {};
  if (cfg.sourceInfoTab) {
    const t = readSourceTable_(src, cfg.sourceInfoTab, ['POS코드', '점포명']);
    t.rows.forEach(function (r) {
      const pos = str_(t.get(r, 'POS코드'));
      if (pos && !info[pos]) info[pos] = { area: num_(t.get(r, '평수')), open: t.get(r, '오픈일') };
    });
  }

  // (3) 손익 탭: 점포별 월 손익 모으기
  const t = readSourceTable_(src, cfg.sourcePlTab, ['년', '월', 'POS코드', '점포명', '순매출액', cfg.profitField]);
  const stores = {};
  let maxYm = '';
  t.rows.forEach(function (r) {
    const pos = str_(t.get(r, 'POS코드'));
    const y = num_(t.get(r, '년')), m = num_(t.get(r, '월'));
    if (!pos || !y || !m) return;               // 합계·빈 행은 건너뜀
    const ym = y + '-' + pad2_(m);
    if (ym > maxYm) maxYm = ym;
    if (!stores[pos]) {
      stores[pos] = { pos: pos, name: str_(t.get(r, '점포명')), org: str_(t.get(r, '소속')),
        org2: str_(t.get(r, '소속 2')), months: [] };
    }
    stores[pos].months.push({ ym: ym, sales: num0_(t.get(r, '순매출액')), profit: num0_(t.get(r, cfg.profitField)) });
  });

  // 폐점한 점포 제외: 폐점비 탭에 있거나, 가장 최근 달 손익이 없는 점포
  const active = [], skipped = [];
  Object.keys(stores).forEach(function (pos) {
    const s = stores[pos];
    s.lastYm = s.months.map(function (x) { return x.ym; }).sort().pop();
    if (closedPos[pos] || s.lastYm < maxYm) skipped.push(s.name); else active.push(s);
  });

  // 브랜드: [브랜드매핑]에 직접 지정한 값 → 없으면 점포명·소속으로 추정
  const brandOf = updateBrandMap_(ctx, active);
  active.forEach(function (s) { s.brand = brandOf(s); });
  active.sort(function (a, b) { return cmp_(a.brand, b.brand) || cmp_(a.name, b.name); });

  const plRows = writeSourceProfitLoss_(ctx, active);
  const assetMsg = upsertSourceAssets_(ctx, active, info);
  const removed = removeSampleStores_(ctx);

  msgs.unshift('손익 ' + active.length + '개 점포·' + plRows + '행 (' + maxYm + '까지, 폐점 ' + skipped.length + '곳 제외)');
  msgs.push(assetMsg);
  if (removed) msgs.push('샘플 점포 ' + removed + '행 삭제');
  const summary = '원본 가져오기 완료: ' + msgs.join(' · ');
  writeConfigValue_(ctx.ss, CFG.LAST_IMPORT, ctx.now + ' · ' + msgs.join(' · '));
  return summary;
}

// 주소(…/d/ID/edit) 또는 ID만 넣어도 ID를 꺼냅니다.
function sourceId_(text) {
  const s = str_(text);
  const m = s.match(/\/d\/([a-zA-Z0-9_-]{20,})/);
  if (m) return m[1];
  return /^[a-zA-Z0-9_-]{20,}$/.test(s) ? s : '';
}

// 원본 탭에서 제목 행(필요한 열 이름이 모두 있는 행)을 찾아 읽습니다. 위쪽 제목·단위 줄은 건너뜀
function readSourceTable_(src, tab, required) {
  const sh = src.getSheetByName(tab);
  if (!sh) throw new Error('원본 시트에 [' + tab + '] 탭이 없습니다. [설정]의 원본 탭 이름을 확인하세요.');
  const values = sh.getDataRange().getValues();
  const norm = function (v) { return str_(v).replace(/\s/g, ''); };
  for (let r = 0; r < Math.min(values.length, 15); r++) {
    const names = values[r].map(norm);
    if (required.every(function (n) { return names.indexOf(norm(n)) !== -1; })) {
      const idx = {};
      names.forEach(function (n, i) { if (n && idx[n] === undefined) idx[n] = i; });
      return {
        rows: values.slice(r + 1),
        unitText: values.slice(0, r).map(function (row) { return row.join(' '); }).join(' '),
        get: function (row, name) { const i = idx[norm(name)]; return i === undefined ? '' : row[i]; }
      };
    }
  }
  throw new Error('원본 [' + tab + '] 탭에서 제목 행(' + required.join(', ') + ')을 찾지 못했습니다.');
}

// '(단위 : 천)' 같은 표시를 읽어 천원으로 바꾸는 배수를 돌려줍니다.
function unitFactor_(text, dflt) {
  const s = str_(text).replace(/\s/g, '');
  if (/단위[:：]?백만/.test(s)) return 1000;
  if (/단위[:：]?천/.test(s)) return 1;
  if (/단위[:：]?만/.test(s)) return 10;
  if (/단위[:：]?원/.test(s)) return 0.001;
  return dflt;
}

// 점포명 앞머리 → 소속 이름 순서로 브랜드 코드를 추정합니다.
function guessBrand_(org, name) {
  const up = function (s) { return str_(s).toUpperCase(); };
  for (let i = 0; i < BRAND_HINTS.length; i++) {
    if (BRAND_HINTS[i][1].some(function (w) { return up(name).indexOf(up(w)) === 0; })) return BRAND_HINTS[i][0];
  }
  for (let i = 0; i < BRAND_HINTS.length; i++) {
    if (BRAND_HINTS[i][1].some(function (w) { return up(org).indexOf(up(w)) !== -1; })) return BRAND_HINTS[i][0];
  }
  return str_(org).replace(/직영|팀|사업부/g, '').trim() || '미분류';
}

// [브랜드매핑] 시트: 원본 '소속'별로 자동 추정 브랜드를 보여 주고, 직접 지정한 브랜드명을 우선 적용합니다.
function updateBrandMap_(ctx, stores) {
  const keyOf = function (s) { return s.org2 && s.org2 !== s.org ? s.org + ' / ' + s.org2 : s.org; };
  let sh = ctx.ss.getSheetByName(SHEETS.BRAND_MAP);
  const manual = {};
  if (sh && sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues().forEach(function (r) {
      if (str_(r[0]) && str_(r[2])) manual[str_(r[0])] = str_(r[2]);
    });
  }
  const groups = {};
  stores.forEach(function (s) {
    const k = keyOf(s);
    if (!groups[k]) groups[k] = { guesses: [], n: 0 };
    groups[k].n++;
    const g = guessBrand_(s.org, s.name);
    if (groups[k].guesses.indexOf(g) === -1) groups[k].guesses.push(g);
  });
  const rows = Object.keys(groups).sort(cmp_).map(function (k) {
    return [k, groups[k].guesses.join(', '), manual[k] || '', groups[k].n];
  });
  if (!sh) {
    sh = ctx.ss.insertSheet(SHEETS.BRAND_MAP);
    sh.getRange(1, 1, 1, 5).setValues([['원본 소속', '자동 추정 브랜드', '브랜드명 (직접 지정)', '점포 수', '설명']])
      .setFontWeight('bold').setBackground(COLOR_AUTO);
    sh.getRange(1, 3).setBackground(COLOR_IN);
    sh.getRange(2, 5).setValue('자동 추정이 틀리면 C열에 브랜드 코드(예: PB, PAS)를 넣고 [🔄 원본 가져오기]를 다시 누르세요. ' +
      'C열이 비어 있으면 점포명·소속으로 추정합니다.');
    sh.setColumnWidth(1, 220); sh.setColumnWidth(2, 140); sh.setColumnWidth(3, 150); sh.setColumnWidth(5, 520);
    sh.setFrozenRows(1);
  }
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 4).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, 4).setValues(rows);
  return function (s) { return manual[keyOf(s)] || guessBrand_(s.org, s.name); };
}

// [손익현황]을 원본 손익으로 다시 채웁니다. (기존 행은 지우고 새로 씀)
function writeSourceProfitLoss_(ctx, stores) {
  const sh = ensureInputSheet_(ctx.ss, SHEETS.PL, PL_COLS);
  const c = readTable_(sh).idx;
  const width = sh.getLastColumn();
  const f = SOURCE_PL_UNIT / ctx.cfg.fPL;       // 원 → [설정]의 손익현황 입력 단위
  const rows = [];
  stores.forEach(function (s) {
    s.months.sort(function (a, b) { return a.ym < b.ym ? -1 : 1; }).forEach(function (m) {
      const row = new Array(width).fill('');
      const sales = Math.round(m.sales * f), profit = Math.round(m.profit * f);
      row[c['브랜드명']] = s.brand;
      row[c['점포명']] = s.name;
      row[c['연월']] = m.ym;
      row[c['매출']] = sales;
      row[c['비용']] = sales - profit;               // 비용 = 매출 − 기준 이익
      row[c['영업이익']] = profit;
      row[c['비고']] = '원본 POS ' + s.pos + ' · ' + ctx.cfg.profitField;
      rows.push(row);
    });
  });
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, width).clearContent();
  if (rows.length) {
    sh.getRange(2, c['연월'] + 1, rows.length, 1).setNumberFormat(FMT_TEXT);
    sh.getRange(2, 1, rows.length, width).setValues(rows);
  }
  return rows.length;
}

// [자산투자비]: 원본 점포가 없으면 행을 추가하고, 평수·투자시점·POS코드는 빈 칸만 채웁니다. (직접 입력한 값은 그대로)
function upsertSourceAssets_(ctx, stores, info) {
  const sh = ensureInputSheet_(ctx.ss, SHEETS.ASSET, ASSET_COLS);
  const t = readTable_(sh);
  const c = t.idx;
  const width = sh.getLastColumn();
  const byPos = {}, byKey = {};
  t.rows.forEach(function (r, i) {
    const pos = str_(r[c['POS코드']]);
    if (pos) byPos[pos] = i;
    byKey[storeKey_(str_(r[c['브랜드명']]), str_(r[c['점포명']]))] = i;
  });
  let added = 0, filled = 0;
  const newRows = [];
  stores.forEach(function (s) {
    const inf = info[s.pos] || {};
    const i = byPos[s.pos] !== undefined ? byPos[s.pos] : byKey[storeKey_(s.brand, s.name)];
    if (i === undefined) {
      const row = new Array(width).fill('');
      row[c['브랜드명']] = s.brand;
      row[c['점포명']] = s.name;
      row[c['평수']] = inf.area || '';
      row[c['투자시점']] = inf.open || '';
      row[c['POS코드']] = s.pos;
      newRows.push(row);
      added++;
      return;
    }
    const r = t.rows[i];
    let touched = false;
    [['평수', inf.area], ['투자시점', inf.open], ['POS코드', s.pos]].forEach(function (x) {
      if (str_(r[c[x[0]]]) === '' && x[1] !== undefined && x[1] !== null && x[1] !== '') {
        sh.getRange(i + 2, c[x[0]] + 1).setValue(x[1]);
        touched = true;
      }
    });
    if (touched) filled++;
  });
  if (newRows.length) sh.getRange(sh.getLastRow() + 1, 1, newRows.length, width).setValues(newRows);
  return '자산투자비 신규 ' + added + '곳·평수/오픈일 보완 ' + filled + '곳';
}

// [폐점비용실적]을 원본 폐점비로 다시 채웁니다. 기존 정산 표시(○/△)는 점포코드로 이어 붙입니다.
function writeClosingActuals_(ctx, items) {
  const sh = ensureInputSheet_(ctx.ss, SHEETS.BS, BS_COLS);
  const t = readTable_(sh);
  const c = t.idx;
  const width = sh.getLastColumn();
  const marks = {};
  t.rows.forEach(function (r) {
    const code = str_(r[c['점포코드']]), mk = str_(r[c['정산']]);
    if (code && mk && marks[code] !== '△') marks[code] = mk;
  });
  const f = 1 / ctx.cfg.fBS;                    // 천원 → [설정]의 폐점비용실적 입력 단위
  const rows = items.map(function (x) {
    const row = new Array(width).fill('');
    row[c['년']] = x.y; row[c['월']] = x.m; row[c['브랜드']] = x.brand; row[c['점포코드']] = x.code;
    row[c['점포명']] = x.name;
    row[c['원상복구(예가)']] = Math.round(x.restore * f);
    row[c['매몰비']] = Math.round(x.sunk * f);
    row[c['정산']] = marks[x.code] || '';
    row[c['비고']] = '원본 폐점비';
    return row;
  });
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, width).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, width).setValues(rows);
}

// 원본을 처음 가져올 때, 시트를 만들 때 넣었던 '샘플 ○○점' 행을 지웁니다.
function removeSampleStores_(ctx) {
  let removed = 0;
  [SHEETS.ASSET, SHEETS.REVIEW].forEach(function (name) {
    const sh = ctx.ss.getSheetByName(name);
    if (!sh || sh.getLastRow() < 2) return;
    const t = readTable_(sh);
    for (let i = t.rows.length - 1; i >= 0; i--) {
      if (/^샘플\s/.test(str_(t.rows[i][t.idx['점포명']]))) { sh.deleteRow(i + 2); removed++; }
    }
  });
  return removed;
}


// ============================================================================
// 1단계. 누적 손익 자동 업데이트
//   - [손익현황] 시트의 월별 데이터를 점포별로 합산합니다.
//   - 매출·비용이 입력되면 영업이익(매출−비용)을 자동으로 채웁니다.
//   - 결과는 [누적손익] 시트에 정리됩니다.
// ============================================================================

function step1_updateProfitLoss_(ctx) {
  const sh = ensureInputSheet_(ctx.ss, SHEETS.PL, PL_COLS);
  const t = readTable_(sh);
  const c = t.idx;
  const f = ctx.cfg.fPL;
  const opColumn = [];          // [손익현황]의 영업이익 열에 다시 써 넣을 값
  const stores = new Map();     // 점포별 월 데이터 모음

  t.rows.forEach(function (r) {
    // (1) 영업이익 자동 계산: 매출과 비용이 둘 다 있으면 매출 − 비용
    const sales = num_(r[c['매출']]);
    const cost = num_(r[c['비용']]);
    let op = num_(r[c['영업이익']]);
    if (sales !== null && cost !== null) op = sales - cost;
    opColumn.push([op === null ? '' : op]);

    // (2) 점포별로 월 데이터 모으기
    const brand = str_(r[c['브랜드명']]);
    const store = str_(r[c['점포명']]);
    if (!brand || !store) return;            // 브랜드/점포명이 없는 행은 건너뜀
    if (sales === null && cost === null && op === null) return;  // 금액이 하나도 없는 행은 건너뜀

    const key = storeKey_(brand, store);
    if (!stores.has(key)) {
      stores.set(key, { brand: brand, store: store, months: new Map(), dup: [], badYm: 0 });
    }
    const s = stores.get(key);
    const ym = toYm_(r[c['연월']], ctx.tz);
    if (!ym) { s.badYm++; return; }          // 연월 형식이 잘못된 행

    if (s.months.has(ym)) s.dup.push(ym);    // 같은 달이 두 번 입력된 경우(합산하되 경고 표시)
    else s.months.set(ym, { sales: 0, cost: 0, op: 0 });
    const m = s.months.get(ym);
    m.sales += (sales || 0) * f;
    m.cost += (cost || 0) * f;
    m.op += (op || 0) * f;
  });

  // (3) 영업이익 열 다시 쓰기
  if (opColumn.length > 0 && !DRY_RUN_) {
    sh.getRange(2, c['영업이익'] + 1, opColumn.length, 1).setValues(opColumn);
  }

  // (4) 점포별 누적 계산
  const result = new Map();
  stores.forEach(function (s, key) {
    const yms = Array.from(s.months.keys()).sort();
    const checks = [];
    let sales = 0, cost = 0, op = 0;
    yms.forEach(function (ym) {
      const m = s.months.get(ym);
      sales += m.sales; cost += m.cost; op += m.op;
    });

    const recent = yms.slice(-3);  // 최근 3개월
    const recent3 = recent.length
      ? recent.reduce(function (sum, ym) { return sum + s.months.get(ym).op; }, 0) / recent.length
      : 0;
    const lastOp = yms.length ? s.months.get(yms[yms.length - 1]).op : 0;

    // 데이터 점검: 중복월, 누락월, 연월 오류
    if (s.dup.length) checks.push('중복월 ' + unique_(s.dup).join(',') + ' (합산됨)');
    if (yms.length >= 2) {
      const expected = ymDiff_(yms[0], yms[yms.length - 1]) + 1;
      if (expected > yms.length) checks.push('누락월 ' + (expected - yms.length) + '개');
    }
    if (s.badYm) checks.push('연월 형식오류 ' + s.badYm + '행');

    result.set(key, {
      brand: s.brand,
      store: s.store,
      firstYm: yms[0] || '',
      lastYm: yms[yms.length - 1] || '',
      months: yms.length,
      sales: sales,
      cost: cost,
      op: op,
      margin: sales > 0 ? op / sales : '',
      recent3: recent3,
      lastOp: lastOp,
      check: checks.join(' / ') || 'OK',
      series: yms.map(function (ym) {   // 월별 실적 (대시보드 추세 분석용)
        const m = s.months.get(ym);
        return { ym: ym, sales: m.sales, cost: m.cost, op: m.op };
      })
    });
  });

  // (5) [누적손익] 시트에 쓰기 (브랜드 → 점포 순 정렬)
  const rows = Array.from(result.values())
    .sort(function (a, b) { return cmp_(a.brand, b.brand) || cmp_(a.store, b.store); })
    .map(function (p) {
      return [p.brand, p.store, p.firstYm, p.lastYm, p.months,
        round_(p.sales), round_(p.cost), round_(p.op), p.margin,
        round_(p.recent3), round_(p.lastOp), p.check, ctx.now];
    });
  writeOutputSheet_(ctx.ss, SHEETS.PL_SUM, PL_SUM_COLS, rows);

  return result;
}


// ============================================================================
// 2단계. 자산 및 투자비 현황 자동 업데이트
//   - [자산투자비] 시트에 입력된 투자비·투자시점·내용연수로 감가상각(정액법)을 계산합니다.
//   - 임차종료일·월임차료로 잔여 임차기간과 잔여 임차료를 계산합니다.
//   - 계산 기준일 = [폐점검토]의 폐점예정일 → [설정] 계산 기준일 → 오늘 순서로 적용합니다.
//   - 결과는 같은 시트의 회색 제목 열(자동 계산 열)에 채워집니다.
// ============================================================================

function step2_updateAssets_(ctx) {
  const sh = ensureInputSheet_(ctx.ss, SHEETS.ASSET, ASSET_COLS);
  const t = readTable_(sh);
  const c = t.idx;
  const f = ctx.cfg.fAsset;
  const closeDates = readClosingDates_(ctx);   // [폐점검토]의 폐점예정일
  const autoRows = [];
  const list = [];
  const seen = {};

  t.rows.forEach(function (r) {
    const brand = str_(r[c['브랜드명']]);
    const store = str_(r[c['점포명']]);
    if (!brand && !store) { autoRows.push({}); return; }   // 빈 행은 자동 계산 열도 비움
    const checks = [];
    if (!brand || !store) checks.push('브랜드명/점포명 누락');
    const key = storeKey_(brand, store);

    // (1) 입력값 읽기 (금액은 천원으로 변환)
    const interior = num0_(r[c['인테리어비']]) * f;
    const fixture = num0_(r[c['집기비']]) * f;
    const etc = num0_(r[c['기타투자비']]) * f;
    const deposit = num0_(r[c['보증금']]) * f;
    const rent = num0_(r[c['월임차료']]) * f;
    const life = num_(r[c['내용연수(년)']]);
    const invDate = toParts_(r[c['투자시점']], ctx.tz);
    const leaseEnd = toParts_(r[c['임차종료일']], ctx.tz);
    const base = closeDates[key] || ctx.baseDate;   // 계산 기준일

    // (2) 감가상각 (정액법, 잔존가치 0 기준)
    //     감가상각누계액 = 투자비 × 경과개월 ÷ (내용연수 × 12)   (내용연수를 넘으면 전액 상각)
    const depBase = interior + fixture + etc;   // 보증금은 돌려받는 돈이므로 제외
    const dep = { depBase: depBase, invDate: invDate, life: life };
    const elapsed = invDate && life > 0 ? Math.max(0, monthsBetween_(invDate, base)) : '';
    if (!(invDate && life > 0) && depBase > 0) {
      checks.push(!invDate ? '투자시점 누락(상각 미반영)' : '내용연수 누락(상각 미반영)');
    }
    const residual = residualAt_(dep, base);    // 잔존자산가치
    const accum = depBase - residual;           // 감가상각누계액

    // (3) 잔여 임차기간·임차료
    //     임차종료일 당일까지 계약기간에 포함되므로 종료일 다음 날까지의 개월 수로 계산
    let remainMonths = 0;
    if (leaseEnd) {
      remainMonths = Math.max(0, monthsBetween_(base, addDays_(leaseEnd, 1)));
    } else if (rent > 0) {
      checks.push('임차종료일 누락');
    }
    if (leaseEnd && rent === 0 && remainMonths > 0) checks.push('월임차료 누락');
    const remainRent = rent * remainMonths;

    // (4) 같은 점포가 두 번 입력된 경우 경고 (요약에는 첫 행만 반영)
    if (seen[key]) checks.push('점포 중복입력(첫 행만 요약 반영)');

    const check = checks.join(' / ') || 'OK';
    autoRows.push({
      '계산기준일': partsToString_(base),
      '경과개월': elapsed === '' ? '' : round1_(elapsed),
      '감가상각누계액(천원)': round_(accum),
      '잔존자산가치(천원)': round_(residual),
      '잔여임차개월': round1_(remainMonths),
      '잔여임차료(천원)': round_(remainRent),
      '데이터점검': check
    });

    if (!seen[key] && brand && store) {
      seen[key] = true;
      list.push({
        key: key, brand: brand, store: store,
        area: num_(r[c['평수']]),
        deposit: deposit,
        residual: residual,
        remainMonths: remainMonths,
        remainRent: remainRent,
        baseDate: partsToString_(base),
        check: check,
        // 대시보드 시뮬레이션용 원본 값
        depBase: depBase, invDate: invDate, life: life, leaseEnd: leaseEnd, rent: rent,
        openDate: partsToString_(invDate), leaseEndDate: partsToString_(leaseEnd)
      });
    }
  });

  // (5) 자동 계산 열에 결과 쓰기 (열 순서가 바뀌어도 제목 이름으로 찾아서 씀)
  writeAutoColumns_(sh, c, ASSET_COLS, autoRows);

  const map = new Map();
  list.forEach(function (a) { map.set(a.key, a); });
  return { list: list, map: map };
}

// 정액법 감가상각 후 남은 금액 (잔존가치 0 기준, 내용연수를 넘으면 0)
//   dep = { depBase: 감가상각 대상 투자비, invDate: 투자시점, life: 내용연수(년) }
//   투자시점이나 내용연수가 없으면 상각하지 않고 투자비 전액을 돌려줍니다.
function residualAt_(dep, date) {
  if (!dep.invDate || !(dep.life > 0)) return dep.depBase;
  const elapsed = Math.max(0, monthsBetween_(dep.invDate, date));
  const totalMonths = dep.life * 12;
  return dep.depBase * (1 - Math.min(elapsed, totalMonths) / totalMonths);
}

// [폐점검토] 시트에서 점포별 폐점예정일만 먼저 읽어 옵니다. (2단계 계산 기준일용)
function readClosingDates_(ctx) {
  const out = {};
  const sh = ctx.ss.getSheetByName(SHEETS.REVIEW);
  if (!sh) return out;
  const t = readTable_(sh);
  const c = t.idx;
  if (c['브랜드명'] === undefined || c['폐점예정일'] === undefined) return out;
  t.rows.forEach(function (r) {
    const d = toParts_(r[c['폐점예정일']], ctx.tz);
    const key = storeKey_(str_(r[c['브랜드명']]), str_(r[c['점포명']]));
    if (d && !out[key]) out[key] = d;
  });
  return out;
}


// ============================================================================
// 3단계-①. 과거 폐점 원복비 실적 → 평당 원복비 기준값 자동 산출
//   - [폐점이력] 시트의 원복 공사비 ÷ 면적 = 평당 원복비를 사례별로 계산합니다.
//   - '입지 유형 × 철거 범위' 조합별로 적게(최소) / 보통(중앙값) / 많이(최대)를 구합니다.
//   - 기타 폐점비는 전체 사례 평균을 사용합니다.
//   - 신뢰도: 사례 5건 이상 '양호', 3~4건 '참고', 3건 미만 '사례 부족' (최소 사례수 3 기준)
//   - 결과는 [폐점기준값] 시트에 정리됩니다.
// ============================================================================

function step3_updateBenchmark_(ctx) {
  const sh = ensureInputSheet_(ctx.ss, SHEETS.HIST, HIST_COLS);
  const t = readTable_(sh);
  const c = t.idx;
  const f = ctx.cfg.fHist;
  const autoRows = [];
  const groups = {};                       // '입지|철거범위' → { values: [평당 원복비...] }
  const all = { values: [], others: [] };  // 전체 사례

  t.rows.forEach(function (r) {
    const brand = str_(r[c['브랜드']]);
    const store = str_(r[c['점포명']]);
    const area = num_(r[c['면적(평)']]);
    const restore = num_(r[c['원복 공사비']]);
    const other = num_(r[c['기타 폐점비']]);
    if (!brand && !store && restore === null && other === null) { autoRows.push({}); return; }

    // 주의 항목 수 = Y 개수
    const cautions = CAUTION_ITEMS.filter(function (name) {
      return str_(r[c[name]]).toUpperCase() === 'Y';
    }).length;

    // 평당 원복비 (천원)
    const perPy = area > 0 && restore !== null ? restore * f / area : null;
    autoRows.push({ '평당 원복비(천원)': perPy === null ? '' : round_(perPy), '주의 항목 수': cautions });

    if (perPy !== null) {
      all.values.push(perPy);
      const key = str_(r[c['입지 유형']]) + '|' + str_(r[c['철거 범위']]);
      if (!groups[key]) groups[key] = { values: [] };
      groups[key].values.push(perPy);
    }
    if (other !== null) all.others.push(other * f);
  });

  writeAutoColumns_(sh, c, HIST_COLS, autoRows);

  // 통계(최소·중앙값·최대·평균) 계산
  Object.keys(groups).forEach(function (k) { groups[k].stats = stats_(groups[k].values); });
  all.stats = stats_(all.values);
  const otherAvg = all.others.length ? sum_(all.others) / all.others.length : 0;

  // [폐점기준값] 시트 쓰기: 기본 9개 조합 + 그 밖에 입력된 조합 + 전체
  const minCases = ctx.cfg.minCases;
  const keys = [];
  LOCATION_TYPES.forEach(function (loc) {
    DEMOLITION_SCOPES.forEach(function (scope) { keys.push(loc + '|' + scope); });
  });
  Object.keys(groups).sort(cmp_).forEach(function (k) { if (keys.indexOf(k) === -1) keys.push(k); });

  const rows = keys.map(function (k) {
    const g = groups[k] || { values: [], stats: null };
    const parts = k.split('|');
    return benchRow_(parts[0] || '(미입력)', parts[1] || '(미입력)', g, '', minCases, ctx.now);
  });
  rows.push(benchRow_('전체', '전체', all, round_(otherAvg), minCases, ctx.now));
  writeOutputSheet_(ctx.ss, SHEETS.BENCH, BENCH_COLS, rows);

  return { groups: groups, all: all, otherAvg: otherAvg };
}

function benchRow_(loc, scope, g, otherAvg, minCases, now) {
  const n = g.values.length;
  const s = g.stats;
  return [loc, scope, n,
    s ? round_(s.min) : '-', s ? round_(s.median) : '-', s ? round_(s.max) : '-', s ? round_(s.mean) : '-',
    otherAvg, reliability_(n, minCases), now];
}

function reliability_(n, minCases) {
  return n >= minCases + 2 ? '양호' : n >= minCases ? '참고' : '사례 부족';
}

// 평당 원복비 참고 범위를 구합니다. (기존 [참고범위] 시트와 같은 규칙)
//   - 시설 인수 협의 = Y 이면 '인수인계' 기준
//   - 같은 '입지 유형 × 철거 범위' 사례가 최소 사례수(기본 3건) 이상이면 그 사례로,
//     부족하면 전체 사례로 계산합니다.
function estimateRestore_(bench, location, scope, takeover, minCases) {
  const applied = takeover === 'Y' ? '인수인계' : scope;
  const key = location + '|' + applied;
  const g = location && applied ? bench.groups[key] : null;
  const similar = g ? g.values.length : 0;

  if (g && similar >= minCases) {
    return { applied: applied, similar: similar, stats: g.stats,
      basis: location + '·' + applied + ' 사례 ' + similar + '건 기준' };
  }
  const why = location && applied
    ? location + '·' + applied + ' 사례 ' + similar + '건(' + minCases + '건 미만)'
    : '입지 유형/철거 범위 미입력';
  if (!bench.all.stats) return { applied: applied, similar: similar, stats: null, basis: '폐점이력 없음(0 적용)' };
  return { applied: applied, similar: similar, stats: bench.all.stats,
    basis: '전체 사례 ' + bench.all.values.length + '건 기준 ※' + why };
}

// 주의 항목 개수로 판단 문구를 만듭니다.
function judgement_(cautions, hasData) {
  if (!hasData) return '실적 데이터 입력 필요';
  if (cautions >= 2) return '주의 항목 ' + cautions + '개 → 보통~많이 쪽 가능성이 높습니다';
  if (cautions === 1) return '주의 항목 1개 → 보통 수준 이상일 수 있습니다';
  return '주의 항목 없음 → 적게~보통 쪽 가능성이 높습니다';
}

// 원상복구 범위 확인·현장 실사 여부로 확정도 등급을 정합니다.
function confidenceGrade_(scopeChecked, siteChecked) {
  if (scopeChecked !== 'Y') return 'C 산정 보류 · 원상복구 범위 미확인 → 내부 참고용만';
  if (siteChecked !== 'Y') return 'B 조건부 · 현장 실사 전 → 범위로만 보고';
  return 'A 산정 가능 · 참고 범위로 보고 가능 (최종 금액은 현장 확인 후 확정)';
}


// ============================================================================
// 3단계-①-2. 실제 폐점 정산 내역(B&S) → 브랜드별 건당 평균 폐점비용
//   - [폐점비용실적] 시트의 원상복구·매몰비·위약금을 점포별로 합산합니다.
//     (같은 점포코드의 소급·정산 차액 행은 하나의 폐점 건으로 묶어 순액으로 계산)
//   - 브랜드별 폐점 점포수, 합계, 건당 평균, 매몰비 비중, 정산 상태를 정리합니다.
//   - 결과는 [브랜드폐점실적] 시트에 정리되고,
//     폐점요약의 '브랜드 과거 건당 원상복구/매몰비' 비교 열과
//     평수가 없는 점포의 원복비 대체값으로 사용됩니다.
// ============================================================================

function step3_updateActuals_(ctx) {
  const sh = ensureInputSheet_(ctx.ss, SHEETS.BS, BS_COLS);
  const t = readTable_(sh);
  const c = t.idx;
  const f = ctx.cfg.fBS;
  const autoRows = [];
  const stores = new Map();   // 점포(폐점 건)별 순액

  t.rows.forEach(function (r) {
    const brand = str_(r[c['브랜드']]);
    const code = str_(r[c['점포코드']]);
    const name = str_(r[c['점포명']]);
    const restore = num_(r[c['원상복구(예가)']]);
    const sunk = num_(r[c['매몰비']]);
    const penalty = num_(r[c['위약금']]);
    if (!brand && !name && restore === null && sunk === null && penalty === null) { autoRows.push({}); return; }

    const total = (restore || 0) + (sunk || 0) + (penalty || 0);
    autoRows.push({ '합계': total });

    // 점포코드가 있으면 코드로, 없으면 브랜드+점포명으로 같은 폐점 건을 묶습니다.
    const key = code ? 'code:' + code : storeKey_(brand, name);
    if (!stores.has(key)) {
      stores.set(key, { brand: brand, restore: 0, sunk: 0, penalty: 0, marks: [] });
    }
    const s = stores.get(key);
    s.restore += (restore || 0) * f;
    s.sunk += (sunk || 0) * f;
    s.penalty += (penalty || 0) * f;
    s.marks.push(str_(r[c['정산']]));
  });

  writeAutoColumns_(sh, c, BS_COLS, autoRows);

  // 브랜드별 집계
  const newGroup = function () {
    return { n: 0, restore: 0, sunk: 0, penalty: 0, settled: 0, open: 0 };
  };
  const byBrand = {};
  const all = newGroup();
  stores.forEach(function (s) {
    // 정산 상태: △가 하나라도 있으면 부분정산, 모두 ○이면 정산완료, 그 밖에는 미정산
    const settled = s.marks.indexOf('△') === -1 &&
      s.marks.every(function (m) { return m === '○'; });
    [byBrand[s.brand] || (byBrand[s.brand] = newGroup()), all].forEach(function (g) {
      g.n++;
      g.restore += s.restore;
      g.sunk += s.sunk;
      g.penalty += s.penalty;
      if (settled) g.settled++; else g.open++;
    });
  });

  const minCases = ctx.cfg.minCases;
  const row = function (label, g) {
    const total = g.restore + g.sunk + g.penalty;
    return [label, g.n, round_(g.restore), round_(g.sunk), round_(g.penalty), round_(total),
      round_(g.restore / g.n), round_(g.sunk / g.n), round_(total / g.n),
      total ? g.sunk / total : '', g.settled, g.open, reliability_(g.n, minCases), ctx.now];
  };
  const rows = Object.keys(byBrand).sort(cmp_).map(function (b) { return row(b, byBrand[b]); });
  if (all.n > 0) rows.push(row('전체', all));
  writeOutputSheet_(ctx.ss, SHEETS.BS_SUM, BS_SUM_COLS, rows);

  return { byBrand: byBrand, all: all };
}

// 브랜드의 과거 건당 평균 원상복구·매몰비 (비교용). 사례가 없으면 null
function brandActual_(actuals, brand) {
  const g = actuals.byBrand[brand];
  if (!g || !g.n) return null;
  return { n: g.n, restore: g.restore / g.n, sunk: g.sunk / g.n };
}

// 평수가 없어 평당 계산을 못 할 때 쓰는 원복비 대체값:
//   브랜드 건당 평균(최소 사례수 이상) → 전체 건당 평균 → 없으면 0
function restoreFallback_(actuals, brand, minCases) {
  const g = actuals.byBrand[brand];
  if (g && g.n >= minCases) {
    return { value: g.restore / g.n, basis: brand + ' 과거 폐점 건당 평균 원상복구(' + g.n + '건) ※평수 미입력' };
  }
  if (actuals.all.n > 0) {
    return { value: actuals.all.restore / actuals.all.n,
      basis: '전체 과거 폐점 건당 평균 원상복구(' + actuals.all.n + '건) ※평수 미입력' };
  }
  return { value: 0, basis: '평수 미입력·폐점비용실적 없음(0 적용)' };
}


// ============================================================================
// 3단계-②. 폐점검토 점포별 원복비 참고 범위 자동 계산
//   - [폐점검토] 시트의 점포 조건(입지 유형, 철거 범위, 주의 항목 등)에
//     [폐점기준값]을 적용해 원복비 적게/보통/많이와 판단·확정도 등급을 채웁니다.
//   - 원복비 = 평당 원복비 × 평수([자산투자비]의 평수)
// ============================================================================

function step3_updateReview_(ctx, assets, bench) {
  const sh = ensureInputSheet_(ctx.ss, SHEETS.REVIEW, REVIEW_COLS);
  const t = readTable_(sh);
  const c = t.idx;
  const f = ctx.cfg.fHist;
  const minCases = ctx.cfg.minCases;
  const autoRows = [];
  const result = new Map();

  t.rows.forEach(function (r) {
    const brand = str_(r[c['브랜드명']]);
    const store = str_(r[c['점포명']]);
    if (!brand && !store) { autoRows.push({}); return; }
    const key = storeKey_(brand, store);
    const yn = function (name) { return str_(r[c[name]]).toUpperCase(); };
    const checks = [];

    const asset = assets.map.get(key);
    const area = asset && asset.area > 0 ? asset.area : null;
    if (!asset) checks.push('자산투자비 미입력');
    else if (!area) checks.push('자산투자비 평수 미입력');

    const location = str_(r[c['입지 유형']]);
    const scope = str_(r[c['철거 범위']]);
    const est = estimateRestore_(bench, location, scope, yn('시설 인수 협의'), minCases);
    const cautions = CAUTION_ITEMS.filter(function (name) { return yn(name) === 'Y'; }).length;

    // 원복비 범위 (천원) = 평당 원복비 × 평수
    const s = est.stats;
    const range = { low: 0, mid: 0, high: 0 };
    let basis = est.basis;
    if (s && area) {
      range.low = s.min * area; range.mid = s.median * area; range.high = s.max * area;
    } else if (!area) {
      // 평수가 없으면 B&S 실적의 건당 평균 원상복구로 대체
      const fb = restoreFallback_(bench.actuals, brand, minCases);
      range.low = fb.value; range.mid = fb.value; range.high = fb.value;
      basis = fb.basis;
    }

    // 직접입력 값이 있으면 그 값을 우선 적용 (적게=보통=많이 모두 같은 값)
    const manualRestore = num_(r[c['원복비_직접입력']]);
    const manualOther = num_(r[c['기타폐점비_직접입력']]);
    const penalty = num0_(r[c['위약금_직접입력']]) * f;
    if (manualRestore !== null) {
      const v = manualRestore * f;
      range.low = v; range.mid = v; range.high = v;
      basis = '원복비 직접입력';
    }
    const other = manualOther !== null ? manualOther * f : bench.otherAvg;

    const judge = manualRestore !== null ? '직접입력 금액 적용' : judgement_(cautions, !!s);
    const grade = confidenceGrade_(yn('원상복구 범위 확인'), yn('현장 실사 완료'));
    if (result.has(key)) checks.push('점포 중복입력(첫 행만 요약 반영)');
    const check = checks.join(' / ') || 'OK';

    autoRows.push({
      '평수': area || '',
      '적용 유형': est.applied || '',
      '비슷한 매장 수': est.similar,
      '평당 원복비 보통(천원)': s ? round_(s.median) : '',
      '원복비 적게(천원)': round_(range.low),
      '원복비 보통(천원)': round_(range.mid),
      '원복비 많이(천원)': round_(range.high),
      '기타 폐점비(천원)': round_(other),
      '주의 항목 수': cautions,
      '판단': judge,
      '확정도 등급': grade,
      '데이터점검': check
    });

    if (!result.has(key)) {
      result.set(key, { brand: brand, store: store, range: range, other: other, penalty: penalty,
        basis: basis, judge: judge, grade: grade, check: check });
    }
  });

  writeAutoColumns_(sh, c, REVIEW_COLS, autoRows);
  return result;
}


// ============================================================================
// 3단계-③. 최종 폐점요약 리스트 자동 생성 (1~3단계 통합)
//   - 점포별: 누적손익, 잔여임차료, 예상폐점비용(원복+매몰), 잔존자산가치, 예상총손실액
//   - 예상폐점비용(원복+매몰) = 예상원복비 + 매몰비(= 2단계 잔존자산가치, 미상각 투자비)
//   - 예상총손실액 = 잔여임차료 + 예상폐점비용(원복+매몰) + 기타 폐점비 + 위약금
//     (잔존자산가치는 매몰비로 이미 포함되어 있으므로 두 번 더하지 않습니다)
//   - 폐점 후 최종손익 = 누적손익 − 예상총손실액 (점포의 전체 기간 성과 참고용)
//   - 예상원복비는 [설정]의 '원복비 적용 수준'(기본: 보통=중앙값)을 사용합니다.
//   - 폐점검토 시트에 없는 점포는 '전체 사례' 기준값 × 평수로 계산합니다.
//   - AI 챗봇 연동을 위해 1행 제목 + 2행부터 데이터, 병합 셀·소계 행 없이 깔끔한 표로 만듭니다.
// ============================================================================

function step3_buildSummary_(ctx, pl, assets, bench, review) {
  const minCases = ctx.cfg.minCases;
  const level = ctx.cfg.level;   // 'low' | 'mid' | 'high'

  // 자산투자비 점포 + 손익현황에만 있는 점포 + 폐점검토에만 있는 점포를 모두 포함
  const keys = assets.list.map(function (a) { return a.key; });
  pl.forEach(function (p, key) { if (keys.indexOf(key) === -1) keys.push(key); });
  review.forEach(function (v, key) { if (keys.indexOf(key) === -1) keys.push(key); });

  let items = keys.map(function (key) {
    const a = assets.map.get(key);
    const p = pl.get(key);
    const rv = review.get(key);
    const brand = a ? a.brand : p ? p.brand : rv.brand;
    const store = a ? a.store : p ? p.store : rv.store;
    const area = a && a.area > 0 ? a.area : null;

    // 데이터 점검 문구 모으기
    const checks = [];
    if (!a) checks.push('자산투자비 미입력');
    else if (a.check !== 'OK') checks.push(a.check);
    if (!p) checks.push('손익현황 미입력');
    else if (p.check !== 'OK') checks.push(p.check);
    if (rv && rv.check !== 'OK') {
      rv.check.split(' / ').forEach(function (x) { if (checks.indexOf(x) === -1) checks.push(x); });
    }

    // 원복비·기타 폐점비: 폐점검토 점포는 검토 조건 반영, 그 외는 전체 사례 기준
    let range, other, penalty, basis, judge, grade;
    if (rv) {
      range = rv.range; other = rv.other; penalty = rv.penalty;
      basis = rv.basis; judge = rv.judge; grade = rv.grade;
    } else {
      const est = estimateRestore_(bench, '', '', '', minCases);
      const s = est.stats;
      basis = est.basis;
      if (s && area) {
        range = { low: s.min * area, mid: s.median * area, high: s.max * area };
      } else {
        const fb = restoreFallback_(bench.actuals, brand, minCases);   // 평수 없으면 B&S 건당 평균
        range = { low: fb.value, mid: fb.value, high: fb.value };
        basis = area ? '폐점이력 없음(0 적용)' : fb.basis;
      }
      other = bench.otherAvg;
      penalty = 0;
      judge = '폐점검토 미입력';
      grade = '';
    }

    const restore = range[level];
    const cumPL = p ? p.op : 0;
    const remainRent = a ? a.remainRent : 0;
    const residual = a ? a.residual : 0;                      // 잔존자산가치 = 예상 매몰비
    const closingCost = restore + residual;                  // 예상폐점비용(원복+매몰)
    // 예상총손실액 (잔여임차료는 [설정]의 반영률만큼만 손실로 봅니다. 기본 100%)
    const totalLoss = remainRent * ctx.cfg.rentRate + closingCost + other + penalty;
    const past = brandActual_(bench.actuals, brand);         // 브랜드 과거 실적 (비교용)

    return {
      key: key, range: range,
      brand: brand, store: store,
      cumPL: cumPL, remainRent: remainRent, closingCost: closingCost,
      residual: residual, totalLoss: totalLoss, finalPL: cumPL - totalLoss,
      review: rv ? 'Y' : 'N',
      area: area, months: p ? p.months : 0, sales: p ? p.sales : 0,
      recent3: p ? p.recent3 : 0,
      remainMonths: a ? a.remainMonths : 0,
      restore: restore, other: other, penalty: penalty,
      costLow: range.low + residual, costHigh: range.high + residual,
      pastRestore: past ? past.restore : '', pastSunk: past ? past.sunk : '',
      basis: basis, judge: judge, grade: grade,
      deposit: a ? a.deposit : 0,
      baseDate: a ? a.baseDate : partsToString_(ctx.baseDate),
      check: checks.join(' / ') || 'OK'
    };
  });

  // 정렬: 브랜드명 → 예상총손실액 큰 순
  items.sort(function (x, y) { return cmp_(x.brand, y.brand) || (y.totalLoss - x.totalLoss); });
  const allItems = items.slice();   // 대시보드·챗봇은 설정과 관계없이 전체 점포를 봅니다

  // [설정]에서 '폐점검토 점포만'을 고르면 검토 대상만 표시
  if (ctx.cfg.scope === SCOPE_REVIEW) {
    items = items.filter(function (x) { return x.review === 'Y'; });
  }

  const rows = items.map(function (x) {
    return [
      x.brand, x.store,
      round_(x.cumPL), round_(x.remainRent), round_(x.closingCost),
      round_(x.residual), round_(x.totalLoss), round_(x.finalPL),
      x.review, x.area === null ? '' : x.area, x.months, round_(x.sales), round_(x.recent3),
      round1_(x.remainMonths), round_(x.restore), round_(x.other), round_(x.penalty),
      round_(x.costLow), round_(x.costHigh), round_(x.pastRestore), round_(x.pastSunk),
      x.basis, x.judge, x.grade,
      round_(x.deposit), x.baseDate, x.check, ctx.now
    ];
  });
  writeOutputSheet_(ctx.ss, SHEETS.SUMMARY, SUMMARY_COLS, rows);

  // [브랜드요약] 시트: 브랜드별 합계
  const brands = new Map();
  items.forEach(function (x) {
    if (!brands.has(x.brand)) {
      brands.set(x.brand, { stores: 0, review: 0, deficit: 0, sales: 0, cumPL: 0,
        remainRent: 0, closingCost: 0, residual: 0, totalLoss: 0, finalPL: 0 });
    }
    const b = brands.get(x.brand);
    b.stores++;
    if (x.review === 'Y') b.review++;
    if (x.cumPL < 0) b.deficit++;
    b.sales += x.sales; b.cumPL += x.cumPL; b.remainRent += x.remainRent;
    b.closingCost += x.closingCost; b.residual += x.residual;
    b.totalLoss += x.totalLoss; b.finalPL += x.finalPL;
  });
  const brandRows = Array.from(brands.keys()).sort(cmp_).map(function (name) {
    const b = brands.get(name);
    return [name, b.stores, b.review, b.deficit, round_(b.sales), round_(b.cumPL),
      round_(b.remainRent), round_(b.closingCost), round_(b.residual),
      round_(b.totalLoss), round_(b.finalPL), ctx.now];
  });
  writeOutputSheet_(ctx.ss, SHEETS.BRAND, BRAND_COLS, brandRows);

  return allItems;
}


// ============================================================================
// 시트 만들기 (처음 한 번 실행)
//   - 없는 시트만 새로 만들고, 이미 있는 시트의 데이터는 건드리지 않습니다.
// ============================================================================

// [메뉴] 샘플 데이터 포함
function setupWithSample() { setup_(true); }

// [메뉴] 빈 양식
function setupBlank() { setup_(false); }

function setup_(withSample) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const created = [];

  // (1) 입력 시트 만들기
  const inputs = [
    [SHEETS.PL, PL_COLS, SAMPLE_PL_()],
    [SHEETS.ASSET, ASSET_COLS, SAMPLE_ASSET],
    [SHEETS.HIST, HIST_COLS, SAMPLE_HIST],
    [SHEETS.REVIEW, REVIEW_COLS, SAMPLE_REVIEW],
    [SHEETS.BS, BS_COLS, SAMPLE_BS]
  ];
  inputs.forEach(function (x) {
    const name = x[0], cols = x[1], sample = x[2];
    const isNew = !ss.getSheetByName(name);
    const sh = ensureInputSheet_(ss, name, cols);
    if (isNew) {
      created.push(name);
      formatInputSheet_(sh, cols);
      if (withSample) writeSample_(sh, cols, sample);
    }
  });

  // (2) 설정 시트, 사용안내 시트 만들기
  if (ensureConfigSheet_(ss)) created.push(SHEETS.CONFIG);
  if (!ss.getSheetByName(SHEETS.GUIDE)) { createGuideSheet_(ss); created.push(SHEETS.GUIDE); }

  // (3) 전체 계산 실행 → 결과 시트 자동 생성 (원본 시트 주소가 있으면 원본 데이터도 가져옴)
  runAll_({ importSource: true });

  // (4) 시트 순서 정리
  const order = [SHEETS.SUMMARY, SHEETS.BRAND, SHEETS.TIMING, SHEETS.PL, SHEETS.PL_SUM, SHEETS.ASSET,
    SHEETS.HIST, SHEETS.BENCH, SHEETS.REVIEW, SHEETS.BS, SHEETS.BS_SUM, SHEETS.CONFIG, SHEETS.GUIDE];
  order.forEach(function (name, i) {
    const sh = ss.getSheetByName(name);
    if (sh) { ss.setActiveSheet(sh); ss.moveActiveSheet(i + 1); }
  });
  // 기본으로 있던 빈 '시트1'은 삭제
  ['시트1', 'Sheet1'].forEach(function (name) {
    const sh = ss.getSheetByName(name);
    if (sh && sh.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(sh);
  });
  ss.setActiveSheet(ss.getSheetByName(SHEETS.SUMMARY));

  showAlert_(
    (created.length ? '새로 만든 시트: ' + created.join(', ') + '\n\n' : '모든 시트가 이미 있어 새로 만든 시트는 없습니다.\n\n') +
    '파란색 제목 열에 데이터를 입력하면 회색 제목 열과 [폐점요약]이 자동으로 계산됩니다.\n' +
    '날짜 경과를 반영하려면 메뉴의 [⏰ 매일 자동 갱신 켜기]도 실행해 주세요.');
}

// 입력 시트 제목 행 꾸미기 (파란색 = 입력, 회색 = 자동 계산) + 선택 목록 설정
function formatInputSheet_(sh, cols) {
  const maxRows = Math.max(sh.getMaxRows(), 1000);
  if (sh.getMaxRows() < maxRows) sh.insertRowsAfter(sh.getMaxRows(), maxRows - sh.getMaxRows());
  cols.forEach(function (col, i) {
    const cell = sh.getRange(1, i + 1);
    cell.setBackground(col.type === AUTO ? COLOR_AUTO : COLOR_IN);
    if (col.note) cell.setNote(col.note);
    if (col.fmt) sh.getRange(2, i + 1, maxRows - 1, 1).setNumberFormat(col.fmt);
    if (col.list) {
      const rule = SpreadsheetApp.newDataValidation().requireValueInList(col.list, true).build();
      sh.getRange(2, i + 1, maxRows - 1, 1).setDataValidation(rule);
    }
    sh.setColumnWidth(i + 1, col.width || 105);
  });
  sh.getRange(1, 1, 1, cols.length).setFontWeight('bold').setHorizontalAlignment('center').setWrap(true);
  sh.setFrozenRows(1);
}

function writeSample_(sh, cols, sample) {
  if (!sample || !sample.length) return;
  const rows = sample.map(function (obj) {
    return cols.map(function (col) { return obj[col.name] === undefined ? '' : obj[col.name]; });
  });
  sh.getRange(2, 1, rows.length, cols.length).setValues(rows);
}

// [설정] 시트 항목 (항목, 기본값, 설명, 선택 목록)
const CONFIG_ROWS = [
  [CFG.UNIT_PL, '천원', '[손익현황] 금액 입력 단위 (원 / 천원 / 만원). 결과는 항상 천원으로 표시됩니다.', Object.keys(UNITS)],
  [CFG.UNIT_ASSET, '천원', '[자산투자비] 금액 입력 단위 (원 / 천원 / 만원)', Object.keys(UNITS)],
  [CFG.UNIT_HIST, '만원', '[폐점이력]·[폐점검토] 금액 입력 단위. 기존 원복비 실적 파일이 만원 단위라 기본값이 만원입니다.', Object.keys(UNITS)],
  [CFG.UNIT_BS, '원', '[폐점비용실적] 금액 입력 단위. 기존 B&S 정산 파일이 원 단위라 기본값이 원입니다.', Object.keys(UNITS)],
  [CFG.BASE, '', '비워 두면 오늘 날짜 기준. 특정 날짜 기준으로 보려면 날짜 입력 (예: 2026-12-31)', null],
  [CFG.MIN, 3, '같은 입지 유형·철거 범위(또는 브랜드) 사례가 이 건수 이상이면 그 사례로, 미만이면 전체 사례로 계산합니다.', null],
  [CFG.LEVEL, '보통(중앙값)', '예상원복비에 쓸 수준 (적게(최소) / 보통(중앙값) / 많이(최대))', Object.keys(LEVELS)],
  [CFG.SCOPE, SCOPE_ALL, '폐점요약에 표시할 점포 범위', [SCOPE_ALL, SCOPE_REVIEW]],
  [CFG.RENT_RATE, 100, '폐점 후에도 내야 하는 잔여 임차료 비율. 계약상 전액 부담이면 100, 위약금으로 대신하면 0 (위약금은 폐점검토에 입력)', null],
  [CFG.TREND, 12, '최적 폐점 시점 계산 시 영업이익 추세를 볼 최근 개월 수', null],
  [CFG.DEP_IN_PL, 'Y', '[손익현황]의 비용에 감가상각비가 포함되어 있으면 Y', YN],
  [CFG.SRC_URL, '', '직영점 데이터 구글 시트 주소. 넣으면 [🔄 원본 가져오기]·매일 자동 갱신 때 손익현황·자산투자비(평수·오픈일)·폐점비용실적을 자동으로 채웁니다. 비우면 연동 안 함', null],
  [CFG.SRC_PL, '시트3', '원본 시트의 월별 손익 탭 이름 (년·월·소속·POS코드·점포명·순매출액·점포이익·사업부이익·영업이익)', null],
  [CFG.SRC_INFO, '점포정보', '원본 시트의 점포 마스터 탭 이름 (POS코드·평수·오픈일). 비우면 가져오지 않음', null],
  [CFG.SRC_CLOSE, '폐점비', '원본 시트의 폐점비 탭 이름 (년·월·브랜드·POS코드·점포명·원상복구·매몰비). 비우면 가져오지 않음', null],
  [CFG.PROFIT, '점포이익', '손익현황에 쓸 이익. 폐점 판단에는 점포이익을 권장합니다 (본부·사업부 배부 비용은 폐점해도 남기 때문)', ['점포이익', '사업부이익', '영업이익']],
  [CFG.LAST_IMPORT, '', '자동 기록 (수정하지 마세요)', null],
  [CFG.LAST, '', '자동 기록 (수정하지 마세요)', null]
];

// [설정] 시트를 만들고, 이미 있으면 빠진 항목만 아래에 추가합니다.
function ensureConfigSheet_(ss) {
  let sh = ss.getSheetByName(SHEETS.CONFIG);
  const isNew = !sh;
  if (isNew) {
    sh = ss.insertSheet(SHEETS.CONFIG);
    sh.getRange(1, 1, 1, 3).setValues([['항목', '값', '설명']])
      .setFontWeight('bold').setBackground(COLOR_IN);
    sh.setColumnWidth(1, 190);
    sh.setColumnWidth(2, 140);
    sh.setColumnWidth(3, 640);
    sh.setFrozenRows(1);
  }
  const labels = sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 1).getValues()
    .map(function (r) { return str_(r[0]); });
  CONFIG_ROWS.forEach(function (row) {
    if (labels.indexOf(row[0]) !== -1) return;
    const r = sh.getLastRow() + 1;
    sh.getRange(r, 1, 1, 3).setValues([[row[0], row[1], row[2]]]);
    if (row[0] !== CFG.LAST && row[0] !== CFG.LAST_IMPORT) sh.getRange(r, 2).setBackground('#FFFDE7');
    if (row[0] === CFG.BASE) sh.getRange(r, 2).setNumberFormat(FMT_DATE);
    if (row[3]) {
      sh.getRange(r, 2).setDataValidation(
        SpreadsheetApp.newDataValidation().requireValueInList(row[3], true).build());
    }
    labels.push(row[0]);
  });
  return isNew;
}

function createGuideSheet_(ss) {
  const sh = ss.insertSheet(SHEETS.GUIDE);
  const lines = [
    ['직영점 폐점 분석 자동화 - 사용 안내'],
    [''],
    ['■ 입력 방법 (파란색 제목 열만 입력하세요. 회색 제목 열은 자동으로 채워집니다.)'],
    ['1단계 [손익현황] 점포별로 매월 한 행씩 브랜드명·점포명·연월(YYYY-MM)·매출·비용 입력 → [누적손익]에 자동 합산'],
    ['2단계 [자산투자비] 점포별 한 행: 평수, 투자시점, 인테리어·집기·기타투자비, 내용연수, 보증금, 임차종료일, 월임차료'],
    ['        → 감가상각누계액, 잔존자산가치, 잔여임차개월, 잔여임차료 자동 계산'],
    ['3단계 [폐점이력] 과거 폐점 원복비 실적 입력 (기존 실적입력 시트의 A5:S 범위를 그대로 붙여넣기 가능)'],
    ['        → [폐점기준값]에 입지 유형 × 철거 범위별 평당 원복비 적게/보통/많이 자동 산출'],
    ['        ※ 과거 사례가 많을수록 정확합니다. 유형별로 최소 3건 이상 입력을 권장합니다.'],
    ['      [폐점검토] 폐점을 검토하는 점포만 한 줄씩: 입지 유형, 철거 범위, 주의 항목 Y/N, 확인 상태 Y/N'],
    ['        → 원복비 적게/보통/많이, 판단, 확정도 등급 자동 계산'],
    ['      [폐점비용실적] 실제 폐점 정산 내역 (기존 B&S 파일 [최종금액] B5:K 범위를 그대로 붙여넣기 가능, 원 단위)'],
    ['        → [브랜드폐점실적]에 브랜드별 건당 평균 원상복구·매몰비 자동 산출 (소급 행은 점포코드로 합산)'],
    [''],
    ['■ 결과 보기'],
    ['[폐점요약] 점포별 누적손익·잔여임차료·예상폐점비용·잔존자산가치·예상총손실액 (AI 챗봇 연동용 표)'],
    ['[브랜드요약] 브랜드별 합계'],
    ['[폐점시점분석] 점포별 최적 폐점 시점 (최근 영업이익 추세 기준) · 즉시 폐점/계약 만료 운영과 비교'],
    ['메뉴 [📊 대시보드 열기] 점포를 고르면 현황, 월 선택 시 예상 폐점 비용, 최적 폐점 시점 차트를 보여줍니다'],
    ['메뉴 [💬 챗봇 열기] 자연어로 질문 (예: "신촌점 언제 닫는 게 좋아?") · 처음 한 번 [🔑 챗봇 API 키 설정] 필요'],
    [''],
    ['■ 계산 기준'],
    ['· 금액은 모두 천원 단위. 입력 단위는 [설정]에서 시트별로 선택 (폐점이력·폐점검토 기본 = 만원)'],
    ['· 감가상각: 정액법, 잔존가치 0. (인테리어+집기+기타) × 경과개월 ÷ (내용연수×12). 보증금은 회수 자산이므로 제외'],
    ['· 계산 기준일: [폐점검토] 폐점예정일 → [설정] 계산 기준일 → 오늘 순서로 적용'],
    ['· 원복비 = 평당 원복비 × 평수. 같은 입지·철거범위 사례 3건 미만이면 전체 사례 기준. 시설 인수 Y면 인수인계 기준'],
    ['· 평수가 없으면 원복비 = [브랜드폐점실적]의 브랜드 건당 평균 원상복구 (3건 미만이면 전체 평균)'],
    ['· 매몰비 = 잔존자산가치(미상각 투자비). B&S 정산의 매몰비와 같은 개념입니다.'],
    ['· 예상폐점비용(원복+매몰) = 예상원복비([설정] 적용 수준, 기본 보통) + 매몰비(잔존자산가치)'],
    ['· 예상총손실액 = 잔여임차료 + 예상폐점비용(원복+매몰) + 기타 폐점비(과거 평균) + 위약금'],
    ['· 폐점 후 최종손익 = 누적손익 − 예상총손실액'],
    ['· 최적 폐점 시점: 이번 달~계약 종료월 각 달에 대해 (그 달까지 예상 영업이익 합계 − 그 달 폐점 비용)이 가장 큰 달'],
    ['  예상 영업이익 = 최근 12개월(설정) 월 영업이익의 직선 추세. 잔여임차료 반영률을 낮추면(위약금 협의 등) 조기 폐점이 유리해질 수 있습니다.'],
    ['· 위약금은 [폐점검토] 위약금_직접입력에, 원복 지연에 따른 추가 비용은 기타폐점비_직접입력에 반영하세요.'],
    [''],
    ['■ 자동 갱신'],
    ['· 입력 시트를 수정하면 즉시 자동 갱신됩니다.'],
    ['· 날짜 경과를 반영하려면 메뉴 [🏪 폐점분석 → ⏰ 매일 자동 갱신 켜기]를 한 번 실행하세요.'],
    ['· 수동 갱신: 메뉴 [🏪 폐점분석 → 🔄 원본 가져오기 + 전체 다시 계산]'],
    ['· 원본 연동: [설정]의 "원본 시트 주소"에 직영점 데이터 시트 주소를 넣으면, 수동 갱신·매일 자동 갱신 때 손익현황·자산투자비·폐점비용실적을 자동으로 채웁니다'],
    ['  (셀을 고칠 때 자동 실행되는 계산에서는 원본을 가져오지 않습니다. 원본이 바뀌었으면 수동 갱신을 눌러 주세요)']
  ];
  sh.getRange(1, 1, lines.length, 1).setValues(lines);
  sh.getRange(1, 1).setFontSize(14).setFontWeight('bold');
  lines.forEach(function (line, i) {   // '■'로 시작하는 제목 줄은 굵게
    if (line[0].charAt(0) === '■') sh.getRange(i + 1, 1).setFontWeight('bold');
  });
  sh.setColumnWidth(1, 900);
}


// ============================================================================
// 공통 도우미 함수 (수정할 필요 없음)
// ============================================================================

// 입력 시트가 없으면 만들고, 필요한 제목 열이 빠져 있으면 오른쪽 끝에 추가합니다.
function ensureInputSheet_(ss, name, cols) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  const lastCol = sh.getLastColumn();
  const headers = lastCol > 0
    ? sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return str_(h); })
    : [];
  let next = headers.filter(function (h) { return h !== ''; }).length === 0 ? 1 : lastCol + 1;
  cols.forEach(function (col) {
    if (headers.indexOf(col.name) === -1) {
      const cell = sh.getRange(1, next);
      cell.setValue(col.name).setFontWeight('bold')
        .setBackground(col.type === AUTO ? COLOR_AUTO : COLOR_IN);
      if (col.note) cell.setNote(col.note);
      if (col.fmt) sh.getRange(2, next, Math.max(sh.getMaxRows() - 1, 1), 1).setNumberFormat(col.fmt);
      next++;
    }
  });
  return sh;
}

// 시트 전체를 읽어서 {제목→열번호, 데이터 행들} 형태로 돌려줍니다.
function readTable_(sh) {
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return { idx: {}, rows: [] };
  const values = sh.getRange(1, 1, lastRow, lastCol).getValues();
  const idx = {};
  values[0].forEach(function (h, i) {
    const name = str_(h);
    if (name && idx[name] === undefined) idx[name] = i;
  });
  return { idx: idx, rows: values.slice(1) };
}

// 입력 시트의 자동 계산 열(회색 제목)에 결과를 씁니다.
//   autoRows: 행마다 { 열이름: 값 } (빈 객체면 그 행의 자동 계산 열을 비움)
function writeAutoColumns_(sh, idx, cols, autoRows) {
  if (!autoRows.length || DRY_RUN_) return;
  cols.filter(function (col) { return col.type === AUTO; }).forEach(function (col) {
    const values = autoRows.map(function (row) {
      const v = row[col.name];
      return [v === undefined || v === null ? '' : v];
    });
    const range = sh.getRange(2, idx[col.name] + 1, values.length, 1);
    range.setNumberFormat(col.fmt || FMT_TEXT);   // 형식을 먼저 지정 후 값 쓰기
    range.setValues(values);
  });
}

// 결과 시트에 표를 씁니다. (1행 제목, 2행부터 데이터 - AI 챗봇이 읽기 쉬운 형태)
function writeOutputSheet_(ss, name, cols, rows) {
  if (DRY_RUN_) return null;
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.clearContents();

  const width = cols.length;
  const header = [cols.map(function (c) { return c.name; })];
  sh.getRange(1, 1, 1, width).setValues(header)
    .setFontWeight('bold').setBackground(COLOR_OUT)
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);

  if (rows.length > 0) {
    // 표시 형식을 먼저 지정한 뒤 값을 씁니다. ("2026-01" 같은 글자가 날짜로 바뀌지 않도록)
    cols.forEach(function (col, i) {
      sh.getRange(2, i + 1, rows.length, 1).setNumberFormat(col.fmt || FMT_TEXT);
    });
    sh.getRange(2, 1, rows.length, width).setValues(rows);
  }
  cols.forEach(function (col, i) {
    sh.setColumnWidth(i + 1, col.width || 95);
    if (col.note) sh.getRange(1, i + 1).setNote(col.note);   // 제목 셀 메모로 계산식 안내
  });
  sh.setFrozenRows(1);
  sh.setFrozenColumns(2);
  return sh;
}

// [설정] 시트 값 읽기
function readConfig_(ss, tz) {
  const cfg = { fPL: 1, fAsset: 1, fHist: UNITS['만원'], fBS: UNITS['원'], baseDate: null, minCases: 3,
    level: 'mid', scope: SCOPE_ALL, rentRate: 1, trendMonths: 12, depInPL: true,
    sourceUrl: '', sourcePlTab: '시트3', sourceInfoTab: '점포정보', sourceClosingTab: '폐점비', profitField: '점포이익' };
  const sh = ss.getSheetByName(SHEETS.CONFIG);
  if (!sh || sh.getLastRow() < 1) return cfg;
  sh.getRange(1, 1, sh.getLastRow(), 2).getValues().forEach(function (r) {
    const k = str_(r[0]);
    const v = r[1];
    const unit = UNITS[str_(v)];
    if (k === CFG.UNIT_PL && unit) cfg.fPL = unit;
    if (k === CFG.UNIT_ASSET && unit) cfg.fAsset = unit;
    if (k === CFG.UNIT_HIST && unit) cfg.fHist = unit;
    if (k === CFG.UNIT_BS && unit) cfg.fBS = unit;
    if (k === CFG.BASE) cfg.baseDate = toParts_(v, tz);
    if (k === CFG.MIN && num_(v) > 0) cfg.minCases = Math.floor(num_(v));
    if (k === CFG.LEVEL && LEVELS[str_(v)]) cfg.level = LEVELS[str_(v)];
    if (k === CFG.SCOPE && str_(v)) cfg.scope = str_(v);
    if (k === CFG.RENT_RATE && num_(v) !== null) cfg.rentRate = Math.min(Math.max(num_(v), 0), 100) / 100;
    if (k === CFG.TREND && num_(v) >= 3) cfg.trendMonths = Math.floor(num_(v));
    if (k === CFG.DEP_IN_PL && str_(v)) cfg.depInPL = str_(v).toUpperCase() !== 'N';
    if (k === CFG.SRC_URL) cfg.sourceUrl = str_(v);
    if (k === CFG.SRC_PL && str_(v)) cfg.sourcePlTab = str_(v);
    if (k === CFG.SRC_INFO) cfg.sourceInfoTab = str_(v);
    if (k === CFG.SRC_CLOSE) cfg.sourceClosingTab = str_(v);
    if (k === CFG.PROFIT && str_(v)) cfg.profitField = str_(v);
  });
  return cfg;
}

function writeConfigValue_(ss, label, value) {
  const sh = ss.getSheetByName(SHEETS.CONFIG);
  if (!sh || sh.getLastRow() < 1) return;
  const labels = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
  for (let i = 0; i < labels.length; i++) {
    if (str_(labels[i][0]) === label) { sh.getRange(i + 1, 2).setValue(value); return; }
  }
}

// 최소·중앙값·최대·평균 계산
function stats_(values) {
  if (!values.length) return null;
  const v = values.slice().sort(function (a, b) { return a - b; });
  const mid = Math.floor(v.length / 2);
  return {
    min: v[0],
    max: v[v.length - 1],
    median: v.length % 2 ? v[mid] : (v[mid - 1] + v[mid]) / 2,
    mean: sum_(v) / v.length
  };
}

function sum_(arr) { return arr.reduce(function (s, x) { return s + x; }, 0); }

// 문자 정리 (앞뒤 공백 제거)
function str_(v) {
  return v === null || v === undefined ? '' : String(v).trim();
}

// 숫자로 변환 ("1,234", "1,234원" 같은 글자도 숫자로). 빈 칸이면 null
function num_(v) {
  if (typeof v === 'number') return isFinite(v) ? v : null;
  if (isDate_(v)) return null;
  const s = str_(v).replace(/[,\s원]/g, '');
  if (s === '') return null;
  const n = Number(s);
  return isFinite(n) ? n : null;
}

// 숫자로 변환하되 빈 칸이면 0
function num0_(v) {
  const n = num_(v);
  return n === null ? 0 : n;
}

// 날짜 → {y, m, d}. 날짜 셀, "2026-01-31", "2026.1.31", "20260131", "2026-01"(1일로 간주) 모두 인식
function toParts_(v, tz) {
  if (isDate_(v)) {
    if (isNaN(v.getTime())) return null;
    const p = Utilities.formatDate(v, tz, 'yyyy-M-d').split('-');
    return { y: Number(p[0]), m: Number(p[1]), d: Number(p[2]) };
  }
  const s = str_(v);
  if (!s) return null;
  const m = s.match(/^(\d{4})\s*[-./년]?\s*(\d{1,2})\s*(?:[-./월]\s*(\d{1,2})|(\d{2}))?/);
  if (!m) return null;
  const y = Number(m[1]);
  const mo = Number(m[2]);
  const d = m[3] ? Number(m[3]) : m[4] ? Number(m[4]) : 1;
  if (mo < 1 || mo > 12 || d < 1 || d > 31) return null;
  return { y: y, m: mo, d: Math.min(d, daysInMonth_(y, mo)) };
}

// 날짜 값인지 확인
function isDate_(v) {
  return Object.prototype.toString.call(v) === '[object Date]';
}

// 연월 → "YYYY-MM"
function toYm_(v, tz) {
  const p = toParts_(v, tz);
  return p ? p.y + '-' + pad2_(p.m) : '';
}

// 두 연월("YYYY-MM") 사이의 개월 차이
function ymDiff_(a, b) {
  const pa = a.split('-'), pb = b.split('-');
  return (Number(pb[0]) - Number(pa[0])) * 12 + (Number(pb[1]) - Number(pa[1]));
}

// 두 날짜 사이의 개월 수 (일 단위는 해당 월 일수로 나누어 소수로 반영)
function monthsBetween_(a, b) {
  return (b.y - a.y) * 12 + (b.m - a.m) + (b.d - a.d) / daysInMonth_(b.y, b.m);
}

function addDays_(p, days) {
  const dt = new Date(Date.UTC(p.y, p.m - 1, p.d + days));
  return { y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate() };
}

function daysInMonth_(y, m) {
  return new Date(Date.UTC(y, m, 0)).getUTCDate();
}

function partsToString_(p) {
  return p ? p.y + '-' + pad2_(p.m) + '-' + pad2_(p.d) : '';
}

function pad2_(n) { return (n < 10 ? '0' : '') + n; }
function round_(n) { return typeof n === 'number' ? Math.round(n) : n; }
function round1_(n) { return typeof n === 'number' ? Math.round(n * 10) / 10 : n; }
function storeKey_(brand, store) { return brand + '||' + store; }
function cmp_(a, b) { return String(a).localeCompare(String(b), 'ko'); }
function unique_(arr) { return arr.filter(function (x, i) { return arr.indexOf(x) === i; }); }

function showAlert_(msg) {
  try {
    SpreadsheetApp.getUi().alert('🏪 폐점분석', msg, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (err) {
    Logger.log(msg);  // 자동 실행(트리거) 중에는 팝업을 띄울 수 없으므로 로그로 남김
  }
}


// ============================================================================
// 샘플 데이터 (🛠 시트 만들기 (샘플 데이터 포함) 실행 시에만 사용)
//   - 실제 데이터를 입력할 때는 샘플 행을 지우고 입력하세요.
//   - 손익현황·자산투자비 = 천원, 폐점이력·폐점검토 = 만원, 폐점비용실적 = 원 단위
//   - 모든 샘플은 가상 데이터입니다. (공개 저장소용 — 실제 데이터는 시트에만 입력)
// ============================================================================

// 4개 샘플 점포 × 24개월 손익을 만들어 줍니다. (브랜드 코드는 실제 B&S 데이터와 같은 표기)
//   매출 = 기준매출 × (1+월성장률)^경과월 × 계절지수,  비용 = 고정비 + 변동비율 × 매출
function SAMPLE_PL_() {
  const stores = [
    // [브랜드, 점포, 기준 월매출, 월 성장률, 월 고정비, 변동비율]
    ['PB', '샘플 강남점', 92000, 0.002, 31000, 0.53],
    ['PB', '샘플 신촌점', 60000, -0.02, 25000, 0.57],      // 매출 감소 추세
    ['PAS', '샘플 판교점', 62000, 0.008, 31500, 0.52],     // 적자지만 개선 추세
    ['PAS', '샘플 서면점', 76000, -0.004, 27000, 0.56]
  ];
  const season = [0.97, 0.95, 1.0, 1.01, 1.03, 1.02, 1.04, 1.05, 1.0, 0.99, 0.97, 1.02];  // 1~12월
  const rows = [];
  stores.forEach(function (s) {
    for (let i = 0; i < 24; i++) {
      const y = 2024 + Math.floor((9 + i) / 12);
      const m = ((9 + i) % 12) + 1;               // 2024-10 ~ 2026-09
      const sales = Math.round(s[2] * Math.pow(1 + s[3], i) * season[m - 1]);
      const cost = Math.round(s[4] + s[5] * sales + (i % 3) * 200);
      rows.push({ '브랜드명': s[0], '점포명': s[1], '연월': y + '-' + pad2_(m), '매출': sales, '비용': cost });
    }
  });
  return rows;
}

const SAMPLE_ASSET = [
  { '브랜드명': 'PB', '점포명': '샘플 강남점', '평수': 35, '투자시점': '2022-04-01',
    '인테리어비': 180000, '집기비': 60000, '기타투자비': 15000, '내용연수(년)': 5, '보증금': 150000,
    '임차개시일': '2022-04-01', '임차종료일': '2027-03-31', '월임차료': 12000 },
  { '브랜드명': 'PB', '점포명': '샘플 신촌점', '평수': 35, '투자시점': '2024-03-01',
    '인테리어비': 130000, '집기비': 45000, '기타투자비': 10000, '내용연수(년)': 5, '보증금': 80000,
    '임차개시일': '2024-03-01', '임차종료일': '2029-02-28', '월임차료': 7500 },
  { '브랜드명': 'PAS', '점포명': '샘플 판교점', '평수': 40, '투자시점': '2023-09-01',
    '인테리어비': 200000, '집기비': 70000, '기타투자비': 20000, '내용연수(년)': 5, '보증금': 120000,
    '임차개시일': '2023-09-01', '임차종료일': '2028-08-31', '월임차료': 9500 },
  { '브랜드명': 'PAS', '점포명': '샘플 서면점', '평수': 32, '투자시점': '2021-11-01',
    '인테리어비': 150000, '집기비': 50000, '기타투자비': 12000, '내용연수(년)': 5, '보증금': 100000,
    '임차개시일': '2021-11-01', '임차종료일': '2027-10-31', '월임차료': 8000 }
];

// 기존 '폐점 원복비 실적 입력 (가상 예시 18건)' 파일의 데이터 (단위: 만원)
const SAMPLE_HIST = [
  [1, 'PB (가상)', '가상 1호점', '로드샵', 30, '2024-02', '전체 철거', 'N', 'N', 'N', 'N', 'N', 'N', 1260, 250, '구매팀', ''],
  [2, 'PB (가상)', '가상 2호점', '로드샵', 35, '2024-05', '전체 철거', 'N', 'N', 'N', 'N', 'N', 'N', 1610, 300, '구매팀', ''],
  [3, 'PB (가상)', '가상 3호점', '로드샵', 28, '2024-09', '전체 철거', 'Y', 'N', 'N', 'N', 'N', 'N', 1400, 250, 'SI팀', ''],
  [4, 'PB (가상)', '가상 4호점', '로드샵', 40, '2025-01', '전체 철거', 'Y', 'N', 'Y', 'N', 'N', 'N', 2120, 350, '구매팀', ''],
  [5, '브랜드A (가상)', '가상 5호점', '로드샵', 32, '2025-04', '전체 철거', 'Y', 'Y', 'Y', 'N', 'N', 'N', 1860, 300, 'SI팀', ''],
  [6, 'PB (가상)', '가상 6호점', '로드샵', 36, '2025-08', '전체 철거', 'Y', 'Y', 'Y', 'N', 'N', 'Y', 3420, 400, '구매팀', '노후 건물 석면 해체 포함'],
  [7, 'PB (가상)', '가상 7호점', '로드샵', 25, '2024-06', '부분 철거', 'N', 'N', 'N', 'N', 'N', 'N', 550, 200, '구매팀', ''],
  [8, 'PB (가상)', '가상 8호점', '로드샵', 30, '2024-11', '부분 철거', 'N', 'N', 'N', 'N', 'N', 'N', 780, 200, 'SI팀', ''],
  [9, '브랜드A (가상)', '가상 9호점', '로드샵', 27, '2025-06', '부분 철거', 'Y', 'N', 'N', 'N', 'N', 'N', 840, 250, '구매팀', ''],
  [10, 'PB (가상)', '가상 10호점', '로드샵', 33, '2025-03', '인수인계', 'N', 'N', 'N', 'N', 'N', 'N', 160, 150, '인테리어실', '다음 임차인 시설 인수'],
  [11, 'PB (가상)', '가상 11호점', '몰·백화점', 38, '2024-03', '전체 철거', 'N', 'N', 'Y', 'Y', 'N', 'N', 2090, 350, '구매팀', ''],
  [12, '브랜드B (가상)', '가상 12호점', '몰·백화점', 42, '2024-08', '전체 철거', 'N', 'N', 'Y', 'Y', 'N', 'N', 2520, 400, '구매팀', ''],
  [13, 'PB (가상)', '가상 13호점', '몰·백화점', 35, '2025-02', '전체 철거', 'Y', 'N', 'Y', 'Y', 'N', 'N', 2240, 350, 'SI팀', ''],
  [14, '브랜드B (가상)', '가상 14호점', '몰·백화점', 45, '2025-05', '전체 철거', 'Y', 'Y', 'Y', 'Y', 'N', 'N', 3150, 450, '구매팀', ''],
  [15, 'PB (가상)', '가상 15호점', '몰·백화점', 40, '2025-09', '전체 철거', 'Y', 'Y', 'Y', 'Y', 'Y', 'N', 3520, 400, '구매팀', '구조 복구 + 반출 제한'],
  [16, 'PB (가상)', '가상 16호점', '몰·백화점', 30, '2024-10', '부분 철거', 'N', 'N', 'N', 'Y', 'N', 'N', 900, 250, 'SI팀', ''],
  [17, '브랜드A (가상)', '가상 17호점', '몰·백화점', 34, '2025-07', '부분 철거', 'Y', 'N', 'N', 'Y', 'N', 'N', 1220, 300, '구매팀', ''],
  [18, 'PB (가상)', '가상 18호점', '오피스·기타', 26, '2025-01', '전체 철거', 'N', 'N', 'N', 'N', 'N', 'N', 1250, 200, '인테리어실', '']
].map(function (r) {
  return { 'No': r[0], '브랜드': r[1], '점포명': r[2], '입지 유형': r[3], '면적(평)': r[4], '폐점 연월': r[5],
    '철거 범위': r[6], '주방 배기 덕트': r[7], '구조 변경 복구': r[8], '설비 원복': r[9],
    '야간·제한 공사': r[10], '반출 어려움': r[11], '석면': r[12], '원복 공사비': r[13],
    '기타 폐점비': r[14], '자료 출처': r[15], '비고': r[16] };
});

// 폐점검토 샘플: 신촌점은 기존 [참고범위] 예시와 같은 조건 (단위: 만원)
const SAMPLE_REVIEW = [
  { '브랜드명': 'PB', '점포명': '샘플 신촌점', '폐점예정일': '2026-12-31', '입지 유형': '로드샵', '철거 범위': '전체 철거',
    '주방 배기 덕트': 'Y', '구조 변경 복구': 'Y', '설비 원복': 'N', '야간·제한 공사': 'Y', '반출 어려움': 'N', '석면': 'Y',
    '시설 인수 협의': 'N', '원상복구 범위 확인': 'N', '현장 실사 완료': 'N' },
  { '브랜드명': 'PAS', '점포명': '샘플 판교점', '입지 유형': '몰·백화점', '철거 범위': '부분 철거',
    '주방 배기 덕트': 'N', '구조 변경 복구': 'N', '설비 원복': 'N', '야간·제한 공사': 'Y', '반출 어려움': 'N', '석면': 'N',
    '시설 인수 협의': 'N', '원상복구 범위 확인': 'Y', '현장 실사 완료': 'N' }
];

// 폐점비용실적 샘플: B&S [최종금액] 시트와 같은 형식의 가상 데이터 (단위: 원)
//   ※ 공개 저장소용 코드라 실제 정산 데이터는 넣지 않았습니다.
//     실제 B&S [최종금액] 시트의 B5:K 범위를 [폐점비용실적] A2에 값만 붙여넣으세요.
//   - 같은 점포코드의 소급 행(음수)은 자동으로 합산되어 한 건으로 계산됩니다.
//   - [년, 월, 브랜드, 점포코드, 점포명, 원상복구(예가), 매몰비, 정산, 비고]
const SAMPLE_BS = [
  [2025, 1, 'PB', 90001, '가상 A점', 85000000, 120500000, '○', ''],
  [2025, 3, 'PB', 90002, '가상 B점', 0, 6200000, '○', '시설 인수'],
  [2025, 5, 'PB', 90003, '가상 C점', 100000000, 79400000, '○', ''],
  [2025, 6, 'PB', 90003, '가상 C점', -9600000, 0, '○', '5월 폐점분 소급'],
  [2025, 4, 'PAS', 90011, '가상 D점', 145000000, 15000000, '○', ''],
  [2025, 7, 'PAS', 90011, '가상 D점', -56700000, 0, '○', '4월 폐점분 소급'],
  [2025, 9, 'PAS', 90012, '가상 E점', 181000000, 23400000, '△', '부분 정산'],
  [2025, 10, 'PAS', 90013, '가상 F점', 49000000, 39800000, '△', '매몰비만 정산']
].map(function (r) {
  return { '년': r[0], '월': r[1], '브랜드': r[2], '점포코드': r[3], '점포명': r[4],
    '원상복구(예가)': r[5], '매몰비': r[6], '정산': r[7], '비고': r[8] };
});


// ============================================================================
// 4단계. 점포 대시보드 - 폐점 시점 시뮬레이션
//   - 점포마다 '이번 달 ~ 계약 종료월'의 각 달 말일에 폐점하는 경우를 모두 계산해 비교합니다.
//   - 순손익(M) = 이번 달부터 M월까지의 예상 영업이익 합계 − M월 폐점 비용
//       M월 폐점 비용 = 남은 임차료 × 반영률 + 잔존자산가치(매몰비) + 원복비 + 기타 폐점비
//                      + 위약금(계약 만료 전에 폐점할 때만)
//   - 순손익이 가장 큰 달 = 최적 폐점 시점
//   - 예상 영업이익은 최근 N개월(설정, 기본 12개월) 실적의 직선 추세(또는 평균)로 계산합니다.
//   - ※ 손익현황의 비용에 감가상각비가 포함되어 있다고 가정합니다. ([설정]에서 변경 가능)
// ============================================================================

const MAX_HORIZON_MONTHS = 120;   // 계약 종료일이 아주 멀어도 최대 10년까지만 계산
const NO_LEASE_HORIZON = 36;      // 임차종료일이 없으면 36개월까지 계산

// [폐점시점분석] 시트 열 구성
const TIMING_COLS = [
  { name: '브랜드명', width: 90 },
  { name: '점포명', width: 120 },
  { name: '폐점검토', width: 70 },
  { name: '최근3개월 평균영업이익(천원)', fmt: FMT_PL, width: 170 },
  { name: '영업이익 추세(천원/월)', fmt: FMT_PL, width: 140, note: '최근 N개월 월 영업이익의 직선 추세 기울기 (+면 개선, −면 악화)' },
  { name: '추세 설명력(R²)', fmt: '0.00', width: 100, note: '1에 가까울수록 추세가 뚜렷합니다. 0.3 미만이면 추세보다 평균을 참고하세요.' },
  { name: '계약종료월', fmt: FMT_TEXT, width: 85 },
  { name: '최적 폐점 시점', fmt: FMT_TEXT, width: 95 },
  { name: '권고', width: 230 },
  { name: '즉시 폐점 시 순손익(천원)', fmt: FMT_PL, width: 160 },
  { name: '최적 시점 순손익(천원)', fmt: FMT_PL, width: 150 },
  { name: '계약만료 운영 시 순손익(천원)', fmt: FMT_PL, width: 180 },
  { name: '즉시 폐점 대비 이득(천원)', fmt: FMT_PL, width: 160 },
  { name: '계약만료 운영 대비 이득(천원)', fmt: FMT_PL, width: 180 },
  { name: '계산 조건', width: 260 },
  { name: '갱신일시', width: 130 }
];

// 1~3단계 계산 결과로 점포별 시뮬레이션 모델을 만듭니다. (시트에 쓰지 않음)
function buildStoreModels_(r) {
  const ctx = r.ctx;
  const cfg = ctx.cfg;
  const startNum = ymNum_(ctx.baseDate.y, ctx.baseDate.m);   // 이번 달
  const closeDates = readClosingDates_(ctx);

  return r.items.map(function (x) {
    const p = r.pl.get(x.key);
    const a = r.assets.map.get(x.key);
    const history = p ? p.series : [];
    const trend = trend_(history, cfg.trendMonths);
    const leaseEnd = a ? a.leaseEnd : null;
    const leaseEndNum = leaseEnd ? ymNum_(leaseEnd.y, leaseEnd.m) : null;
    const horizon = Math.max(0, Math.min(
      leaseEndNum === null ? NO_LEASE_HORIZON : leaseEndNum - startNum, MAX_HORIZON_MONTHS));

    // 폐점 후보 월별 값: 예상 영업이익, 잔존자산가치, 남은 임차료
    const timeline = [];
    for (let i = 0; i <= horizon; i++) {
      const n = startNum + i;
      const y = Math.floor(n / 12);
      const m = n % 12 + 1;
      const monthEnd = { y: y, m: m, d: daysInMonth_(y, m) };
      let opTrend = trend.intercept + trend.slope * n;
      let opFlat = trend.avg;
      if (!cfg.depInPL && a) {
        // 손익에 감가상각비가 없으면, 그 달 감가상각비를 빼서 회계상 손익으로 맞춥니다.
        const dep = residualAt_(a, addDays_({ y: y, m: m, d: 1 }, -1)) - residualAt_(a, monthEnd);
        opTrend -= dep;
        opFlat -= dep;
      }
      timeline.push({
        ym: y + '-' + pad2_(m),
        opTrend: opTrend,
        opFlat: opFlat,
        residual: a ? residualAt_(a, monthEnd) : 0,
        rentAfter: a && leaseEnd
          ? a.rent * Math.max(0, monthsBetween_(addDays_(monthEnd, 1), addDays_(leaseEnd, 1))) : 0,
        leaseEnd: n === leaseEndNum
      });
    }

    const planned = closeDates[x.key];
    return {
      key: x.key, brand: x.brand, store: x.store, review: x.review, area: x.area,
      basis: x.basis, judge: x.judge, grade: x.grade, check: x.check,
      status: {
        baseDate: x.baseDate,
        cumSales: x.sales, cumOP: x.cumPL,
        margin: p && p.sales > 0 ? p.op / p.sales : null,
        months: x.months, firstYm: p ? p.firstYm : '', lastYm: p ? p.lastYm : '',
        recent3: x.recent3, lastOp: p ? p.lastOp : 0,
        deposit: x.deposit, residualNow: x.residual,
        remainMonths: x.remainMonths, remainRent: x.remainRent,
        closingCostNow: x.closingCost, totalLossNow: x.totalLoss,
        rent: a ? a.rent : 0, openDate: a ? a.openDate : '', leaseEndDate: a ? a.leaseEndDate : '',
        pastRestore: x.pastRestore === '' ? null : x.pastRestore,
        pastSunk: x.pastSunk === '' ? null : x.pastSunk
      },
      fixed: { restore: x.range, other: x.other, penalty: x.penalty },
      trend: { slope: trend.slope, avg: trend.avg, n: trend.n, r2: trend.r2 },
      history: history,
      timeline: timeline,
      leaseEndYm: leaseEnd ? leaseEnd.y + '-' + pad2_(leaseEnd.m) : '',
      plannedYm: planned ? planned.y + '-' + pad2_(planned.m) : ''
    };
  });
}

// 최근 n개월 월 영업이익의 직선 추세 (최소제곱법). 3개월 미만이면 평균만 사용
function trend_(history, n) {
  const pts = history.slice(-n).map(function (h) {
    const p = h.ym.split('-');
    return { x: ymNum_(Number(p[0]), Number(p[1])), y: h.op };
  });
  if (!pts.length) return { slope: 0, intercept: 0, avg: 0, n: 0, r2: null };
  const avgY = sum_(pts.map(function (q) { return q.y; })) / pts.length;
  if (pts.length < 3) return { slope: 0, intercept: avgY, avg: avgY, n: pts.length, r2: null };
  const avgX = sum_(pts.map(function (q) { return q.x; })) / pts.length;
  let sxx = 0, sxy = 0, syy = 0;
  pts.forEach(function (q) {
    sxx += (q.x - avgX) * (q.x - avgX);
    sxy += (q.x - avgX) * (q.y - avgY);
    syy += (q.y - avgY) * (q.y - avgY);
  });
  const slope = sxx ? sxy / sxx : 0;
  const r2 = syy ? (sxy * sxy) / (sxx * syy) : null;
  return { slope: slope, intercept: avgY - slope * avgX, avg: avgY, n: pts.length, r2: r2 };
}

function ymNum_(y, m) { return y * 12 + (m - 1); }

// ※ 이 함수는 대시보드 화면(브라우저)에도 그대로 복사되어 쓰입니다.
//   다른 함수를 부르지 않도록 이 안에서만 계산하고, 옛날 자바스크립트 문법(var)으로 작성했습니다.
//   opts: { rate: 잔여임차료 반영률(0~1), level: 'low'|'mid'|'high', method: 'trend'|'flat' }
function evaluateClosing_(model, opts) {
  opts = opts || {};
  var rate = typeof opts.rate === 'number' ? opts.rate : 1;
  var level = opts.level || 'mid';
  var method = opts.method === 'flat' ? 'flat' : 'trend';
  var tl = model.timeline || [];
  var restore = (model.fixed.restore && model.fixed.restore[level]) || 0;
  var other = model.fixed.other || 0;
  var penaltyAmt = model.fixed.penalty || 0;
  var rows = [];
  var cum = 0;
  var best = 0;
  for (var i = 0; i < tl.length; i++) {
    var t = tl[i];
    var op = method === 'flat' ? t.opFlat : t.opTrend;
    cum += op;
    var rent = t.rentAfter * rate;
    var penalty = t.leaseEnd ? 0 : penaltyAmt;          // 계약 만료월 폐점이면 위약금 없음
    var cost = rent + t.residual + restore + other + penalty;
    rows.push({ ym: t.ym, op: op, cumOp: cum, rent: rent, residual: t.residual, restore: restore,
      other: other, penalty: penalty, cost: cost, value: cum - cost });
    if (rows[i].value > rows[best].value + 0.5) best = i;   // 0.5천원 미만 차이는 이른 달 우선
  }
  var last = rows.length - 1;
  var out = { rows: rows, best: best, last: last, rate: rate, level: level, method: method,
    gainVsNow: 0, gainVsLast: 0 };
  if (!rows.length) { out.recommendation = '계산할 기간이 없습니다'; return out; }
  out.gainVsNow = rows[best].value - rows[0].value;
  out.gainVsLast = rows[best].value - rows[last].value;
  if (last === 0) out.recommendation = '이번 달(' + rows[0].ym + ')이 계약 종료월입니다';
  else if (best === 0) out.recommendation = '즉시 폐점(' + rows[0].ym + ')이 가장 유리';
  else if (best === last) {
    out.recommendation = model.leaseEndYm
      ? '계약 만료(' + rows[last].ym + ')까지 운영이 가장 유리'
      : '검토 기간 끝(' + rows[last].ym + ')까지 운영이 유리 (임차종료일 미입력)';
  } else out.recommendation = rows[best].ym + ' 폐점이 가장 유리';
  return out;
}

// [폐점시점분석] 시트: 점포별 최적 폐점 시점 (기본 설정 기준)
function step4_writeTimingSheet_(r) {
  const cfg = r.ctx.cfg;
  const opts = { rate: cfg.rentRate, level: cfg.level, method: 'trend' };
  const levelName = Object.keys(LEVELS).filter(function (k) { return LEVELS[k] === cfg.level; })[0];
  const cond = '잔여임차료 반영 ' + Math.round(cfg.rentRate * 100) + '% · 원복비 ' + levelName +
    ' · 최근 ' + cfg.trendMonths + '개월 추세';
  const rows = buildStoreModels_(r).map(function (m) {
    const ev = evaluateClosing_(m, opts);
    const has = ev.rows.length > 0;
    const at = function (i) { return has ? round_(ev.rows[i].value) : ''; };
    return {
      gain: ev.gainVsLast,
      row: [m.brand, m.store, m.review, round_(m.status.recent3), round_(m.trend.slope),
        m.trend.r2 === null ? '' : Math.round(m.trend.r2 * 100) / 100,
        m.leaseEndYm, has ? ev.rows[ev.best].ym : '', ev.recommendation,
        at(0), at(ev.best), at(ev.last), round_(ev.gainVsNow), round_(ev.gainVsLast), cond, r.ctx.now]
    };
  });
  rows.sort(function (a, b) { return cmp_(a.row[0], b.row[0]) || (b.gain - a.gain); });
  writeOutputSheet_(r.ctx.ss, SHEETS.TIMING, TIMING_COLS, rows.map(function (x) { return x.row; }));
}

// [메뉴] 대시보드 열기 (시트 위에 큰 창으로 표시)
function openDashboard() {
  const html = renderPage_('full').setWidth(1400).setHeight(900);
  SpreadsheetApp.getUi().showModalDialog(html, '📊 직영점 폐점 분석 대시보드');
}

// [메뉴] 챗봇 열기 (오른쪽 사이드바)
function openChatSidebar() {
  SpreadsheetApp.getUi().showSidebar(renderPage_('chat').setTitle('💬 폐점분석 챗봇'));
}

// 웹앱으로 배포하면 브라우저 새 탭(전체 화면)으로 대시보드를 열 수 있습니다.
//   배포 방법: Apps Script 편집기 [배포 → 새 배포 → 유형 선택(⚙) → 웹 앱]
//   링크 끝에 ?mode=chat 을 붙이면 챗봇만 크게 열립니다.
function doGet(e) {
  const mode = e && e.parameter && e.parameter.mode === 'chat' ? 'chat' : 'full';
  return renderPage_(mode)
    .setTitle(mode === 'chat' ? '폐점분석 챗봇' : '직영점 폐점 분석 대시보드')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// [메뉴] 대시보드 새 창으로 열기 (웹앱 링크)
function openDashboardLink() {
  getSpreadsheet_();   // 웹앱에서 이 스프레드시트를 찾을 수 있도록 ID를 기억해 둡니다
  const url = ScriptApp.getService().getUrl();
  let body;
  if (url) {
    body =
      '<p style="margin:0 0 12px">아래 링크를 누르면 대시보드가 새 탭에서 열립니다. 즐겨찾기해 두고 써도 됩니다.</p>' +
      '<p><a href="' + url + '" target="_blank" style="font-size:15px;font-weight:700">📊 대시보드 열기</a>' +
      ' &nbsp;·&nbsp; <a href="' + url + '?mode=chat" target="_blank">💬 챗봇만 열기</a></p>' +
      '<p style="color:#777;font-size:12px;word-break:break-all">' + url + '</p>' +
      '<script>window.open(' + JSON.stringify(url) + ', "_blank");</script>';
  } else {
    body =
      '<p style="margin:0 0 8px"><b>아직 웹앱으로 배포되지 않았습니다.</b> 처음 한 번만 아래처럼 배포해 주세요.</p>' +
      '<ol style="padding-left:18px;line-height:1.7;margin:0">' +
      '<li>[확장 프로그램 → Apps Script] 편집기 오른쪽 위 <b>[배포 → 새 배포]</b></li>' +
      '<li>유형 선택(⚙) → <b>웹 앱</b></li>' +
      '<li>다음 사용자 인증 정보로 실행: <b>나</b> / 액세스 권한: <b>조직 내 모든 사용자</b>(또는 필요한 범위)</li>' +
      '<li>[배포] → 권한 허용 → 나오는 <b>웹 앱 URL</b>이 대시보드 링크입니다</li>' +
      '</ol><p style="color:#777;font-size:12px">배포 후 이 메뉴를 다시 누르면 링크가 바로 열립니다.</p>';
  }
  const html = HtmlService.createHtmlOutput('<div style="font-family:sans-serif;font-size:14px">' + body + '</div>')
    .setWidth(460).setHeight(url ? 190 : 300);
  SpreadsheetApp.getUi().showModalDialog(html, '🔗 대시보드 새 창으로 열기');
}

function renderPage_(mode) {
  // 화면 HTML은 이 파일 맨 아래 DASHBOARD_HTML_ 에 들어 있습니다. (별도 HTML 파일 불필요)
  const t = HtmlService.createTemplate(DASHBOARD_HTML_);
  t.mode = mode;
  t.sharedJs = evaluateClosing_.toString();   // 서버와 화면이 같은 계산식을 쓰도록 복사
  return t.evaluate();
}

// 대시보드 화면이 불러가는 데이터 (google.script.run 으로 호출)
function getDashboardData() {
  const r = computeAll_(true);
  return JSON.stringify({
    asOf: partsToString_(r.ctx.baseDate),
    updated: r.ctx.now,
    sheetUrl: r.ctx.ss.getUrl(),   // [📄 원본 시트 열기] 버튼용 (주소는 코드에 저장하지 않음)
    defaults: { rate: r.ctx.cfg.rentRate, level: r.ctx.cfg.level, method: 'trend',
      trendMonths: r.ctx.cfg.trendMonths, depInPL: r.ctx.cfg.depInPL },
    models: buildStoreModels_(r)
  });
}


// ============================================================================
// 5단계. 폐점분석 챗봇 (Google Gemini API)
//   - [🔑 챗봇 API 키 설정 (Gemini)]에서 Gemini API 키를 한 번 등록하면 됩니다.
//     (허브의 회의록 챗봇·도면 3D 등과 같은 Gemini 키를 써도 됩니다)
//   - 챗봇은 아래 '도구'로 시트 데이터를 직접 조회·계산합니다. (숫자를 지어내지 않도록)
//     대시보드와 같은 계산 함수(evaluateClosing_)를 쓰므로 결과가 항상 일치합니다.
//   - 질문과 조회 결과는 Google Gemini API로 전송됩니다. (사내 공유 데이터 기준)
// ============================================================================

const GEMINI_MODEL = 'gemini-2.5-flash';   // 허브의 다른 도구와 같은 모델
const CHAT_MAX_TOOL_ROUNDS = 8;            // 한 질문에 도구를 최대 몇 번 연달아 쓸지

// [메뉴] Gemini API 키 등록
function setChatApiKey() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('🔑 챗봇 API 키 설정 (Gemini)',
    'Gemini API 키(AIza...)를 붙여넣으세요.\n키는 이 스크립트 속성에만 저장되고 시트에는 표시되지 않습니다.',
    ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const key = res.getResponseText().trim();
  if (!key) return;
  PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', key);
  ui.alert('Gemini API 키를 저장했습니다. [💬 챗봇 열기] 또는 대시보드의 챗봇 탭에서 사용해 보세요.');
}

// 챗봇이 쓸 수 있는 도구 목록
const CHAT_TOOLS = [
  {
    name: 'list_stores',
    description: '전체 점포(또는 한 브랜드)의 현황과 최적 폐점 시점을 요약한 목록을 가져옵니다. ' +
      '점포 비교, 순위, "어느 점포가 ~" 같은 질문에 사용하세요. 금액 단위는 천원입니다.',
    input_schema: {
      type: 'object',
      properties: {
        brand: { type: 'string', description: '브랜드명 (예: PB, PAS). 비우면 전체 브랜드' },
        review_only: { type: 'boolean', description: 'true면 폐점검토 시트에 올라간 점포만' }
      }
    }
  },
  {
    name: 'get_store_detail',
    description: '한 점포의 상세 현황을 가져옵니다: 누적 매출·손익, 최근 월별 실적, 영업이익 추세, ' +
      '잔존자산가치, 잔여 임차기간·임차료, 원복비 범위, 판단, 확정도 등급, 데이터 점검 결과. 금액 단위는 천원입니다.',
    input_schema: {
      type: 'object',
      properties: {
        store: { type: 'string', description: '점포명 (일부만 입력해도 찾습니다)' },
        brand: { type: 'string', description: '브랜드명 (같은 이름의 점포가 여러 개일 때)' }
      },
      required: ['store']
    }
  },
  {
    name: 'simulate_closing',
    description: '특정 월 말일에 폐점할 때의 예상 폐점 비용 내역(잔여임차료, 매몰비, 원복비, 기타, 위약금)과 ' +
      '순손익을 계산하고, 최적 시점·계약만료 운영과 비교합니다. 금액 단위는 천원입니다.',
    input_schema: {
      type: 'object',
      properties: {
        store: { type: 'string', description: '점포명' },
        brand: { type: 'string', description: '브랜드명 (선택)' },
        close_month: { type: 'string', description: '폐점 월 (YYYY-MM)' },
        rent_rate_percent: { type: 'number', description: '폐점 후에도 내야 하는 잔여 임차료 비율 0~100. 비우면 설정값' },
        restore_level: { type: 'string', enum: ['low', 'mid', 'high'], description: '원복비 수준 (적게/보통/많이). 비우면 설정값' },
        forecast: { type: 'string', enum: ['trend', 'flat'], description: '손익 전망 방식: trend=최근 추세 연장, flat=최근 평균 유지' }
      },
      required: ['store', 'close_month']
    }
  },
  {
    name: 'find_optimal_closing',
    description: '이번 달부터 계약 종료월까지 매달 폐점하는 경우를 모두 비교해 순손익이 가장 큰 최적 폐점 시점을 찾습니다. ' +
      '가정(잔여임차료 반영률, 원복비 수준, 손익 전망 방식)을 바꿔 민감도를 볼 수 있습니다. 금액 단위는 천원입니다.',
    input_schema: {
      type: 'object',
      properties: {
        store: { type: 'string', description: '점포명' },
        brand: { type: 'string', description: '브랜드명 (선택)' },
        rent_rate_percent: { type: 'number', description: '잔여 임차료 반영률 0~100. 비우면 설정값' },
        restore_level: { type: 'string', enum: ['low', 'mid', 'high'], description: '원복비 수준. 비우면 설정값' },
        forecast: { type: 'string', enum: ['trend', 'flat'], description: '손익 전망 방식' }
      },
      required: ['store']
    }
  },
  {
    name: 'get_brand_closing_actuals',
    description: '실제 과거 폐점 정산(B&S) 기준 브랜드별 폐점 점포수, 원상복구·매몰비 합계와 건당 평균을 가져옵니다. 금액 단위는 천원입니다.',
    input_schema: {
      type: 'object',
      properties: { brand: { type: 'string', description: '브랜드명. 비우면 전체 브랜드' } }
    }
  },
  {
    name: 'get_restore_benchmark',
    description: '과거 원복비 실적 기준 입지 유형 × 철거 범위별 평당 원복비(적게=최소/보통=중앙값/많이=최대)와 ' +
      '기타 폐점비 평균을 가져옵니다. 금액 단위는 천원입니다.',
    input_schema: { type: 'object', properties: {} }
  }
];

// 대시보드·사이드바의 챗봇이 호출하는 함수
//   messagesJson: 지금까지의 대화 전체 (Gemini contents 형식, 문자열)
//   돌려주는 값: { messages: 새 대화 전체, note: 안내 문구, error: 오류 문구 }
function chatTurn(messagesJson) {
  const contents = JSON.parse(messagesJson);
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!key) {
    return JSON.stringify({ messages: contents,
      error: 'Gemini API 키가 없습니다. 시트 메뉴 [🏪 폐점분석 → 🔑 챗봇 API 키 설정 (Gemini)]에서 먼저 등록해 주세요.' });
  }

  const data = chatData_();
  for (let round = 0; round < CHAT_MAX_TOOL_ROUNDS; round++) {
    const res = callGemini_(key, data.system, contents);
    if (res.error) return JSON.stringify({ messages: contents, error: res.error });

    const cand = res.candidates && res.candidates[0];
    if (!cand || !cand.content || !cand.content.parts) {
      const why = (res.promptFeedback && res.promptFeedback.blockReason) || (cand && cand.finishReason) || '알 수 없음';
      return JSON.stringify({ messages: contents, note: '답변을 만들지 못했습니다 (' + why + '). 질문을 바꿔 다시 시도해 주세요.' });
    }
    // 응답은 손대지 않고 그대로 대화에 붙입니다. (다음 요청에서 그대로 다시 보내야 합니다)
    contents.push(cand.content);

    const calls = cand.content.parts.filter(function (p) { return p.functionCall; });
    if (calls.length) {
      // 요청된 도구를 모두 실행하고, 결과를 한 번에 돌려줍니다.
      contents.push({ role: 'user', parts: calls.map(function (p) {
        const fc = p.functionCall;
        let response;
        try {
          response = { result: roundDeep_(runChatTool_(fc.name, fc.args || {}, data)) };
        } catch (err) {
          response = { error: err.message };
        }
        return { functionResponse: { name: fc.name, response: response } };
      }) });
      continue;
    }

    let note = '';
    if (cand.finishReason === 'MAX_TOKENS') note = '답변이 길어 중간에 잘렸습니다. 범위를 좁혀 다시 물어봐 주세요.';
    if (cand.finishReason === 'SAFETY') note = '안전 필터 때문에 답변이 중단되었습니다. 질문을 바꿔 다시 시도해 주세요.';
    return JSON.stringify({ messages: contents, note: note });
  }
  return JSON.stringify({ messages: contents, note: '조회가 너무 많아 중단했습니다. 질문을 나눠서 해 주세요.' });
}

// Gemini API 호출 (일시적 오류는 최대 3번까지 재시도)
function callGemini_(key, system, contents) {
  const body = {
    systemInstruction: { parts: [{ text: system }] },
    contents: contents,
    tools: [{ functionDeclarations: CHAT_TOOLS.map(toGeminiFunction_) }],
    generationConfig: { temperature: 0.2, maxOutputTokens: 8192 }
  };
  const options = {
    method: 'post',
    contentType: 'application/json',
    muteHttpExceptions: true,
    headers: { 'x-goog-api-key': key },
    payload: JSON.stringify(body)
  };
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + GEMINI_MODEL + ':generateContent';
  for (let attempt = 0; attempt < 3; attempt++) {
    const resp = UrlFetchApp.fetch(url, options);
    const code = resp.getResponseCode();
    let json = null;
    try { json = JSON.parse(resp.getContentText()); } catch (err) { json = null; }
    if (code === 200 && json) return json;

    const msg = json && json.error ? json.error.message : resp.getContentText().slice(0, 300);
    if (code === 429 || code === 503 || code >= 500) {   // 혼잡·일시 오류 → 잠시 후 재시도
      Utilities.sleep(2000 * (attempt + 1));
      continue;
    }
    if (code === 400 && /API key/i.test(msg) || code === 401 || code === 403) {
      return { error: 'Gemini API 키가 올바르지 않거나 권한이 없습니다. [🔑 챗봇 API 키 설정 (Gemini)]에서 다시 등록해 주세요. (' + msg + ')' };
    }
    return { error: 'Gemini API 오류 (' + code + '): ' + msg };
  }
  return { error: 'Gemini API가 혼잡합니다. 잠시 후 다시 시도해 주세요.' };
}

// 도구 정의를 Gemini 형식(functionDeclarations)으로 바꿉니다.
function toGeminiFunction_(tool) {
  const conv = function (sc) {
    const o = { type: String(sc.type).toUpperCase() };
    if (sc.description) o.description = sc.description;
    if (sc.enum) o.enum = sc.enum;
    if (sc.properties) {
      o.properties = {};
      Object.keys(sc.properties).forEach(function (k) { o.properties[k] = conv(sc.properties[k]); });
    }
    if (sc.required) o.required = sc.required;
    return o;
  };
  const fn = { name: tool.name, description: tool.description };
  const props = tool.input_schema && tool.input_schema.properties;
  if (props && Object.keys(props).length) fn.parameters = conv(tool.input_schema);   // 입력이 없는 도구는 parameters 생략
  return fn;
}

// 챗봇 한 턴 동안 쓸 데이터와 지침을 준비합니다.
function chatData_() {
  const r = computeAll_(true);
  const cfg = r.ctx.cfg;
  const models = buildStoreModels_(r);
  const levelName = Object.keys(LEVELS).filter(function (k) { return LEVELS[k] === cfg.level; })[0];
  const brands = unique_(models.map(function (m) { return m.brand; }));
  const system = [
    '당신은 사내 직영점 폐점 의사결정을 돕는 분석 도우미입니다. Google Sheets의 폐점분석 데이터를 도구로 조회해 답합니다.',
    '',
    '[원칙]',
    '- 숫자는 반드시 도구 결과에서 가져오세요. 직접 추정하거나 지어내지 말고, 데이터에 없으면 없다고 답하세요.',
    '- 금액 단위는 천원입니다. 큰 금액은 "123,456천원(약 1.2억원)"처럼 원화 환산을 함께 적으세요.',
    '- 최적 폐점 시점은 가정(손익 추세, 잔여임차료 반영률, 원복비 수준)에 따라 달라집니다. 결론과 함께 가정을 밝히고, 결론이 바뀔 만한 가정이 있으면 민감도를 확인해 알려주세요.',
    '- 추세 설명력(R²)이 0.3 미만이면 추세가 불확실하다고 알려주고, 평균 유지(forecast=flat) 결과도 함께 확인하세요.',
    '- 확정도 등급이 B나 C면 내부 참고용 추정치임을 알려주세요. 데이터점검에 문제가 있으면 함께 알려주세요.',
    '- 답변은 한국어로, 결론을 먼저 쓰고 근거를 짧게 붙이세요. 여러 점포를 비교할 때는 마크다운 표를 쓰세요.',
    '',
    '[계산 기준]',
    '- 예상폐점비용(원복+매몰) = 예상원복비 + 매몰비(= 미상각 잔존자산가치)',
    '- 예상총손실액 = 잔여임차료 × 반영률 + 예상폐점비용 + 기타 폐점비 + 위약금',
    '- 순손익(M월 폐점) = 이번 달부터 M월까지 예상 영업이익 합계 − M월 폐점 비용. 순손익이 가장 큰 달이 최적 폐점 시점입니다.',
    '- 예상 영업이익 = 최근 ' + cfg.trendMonths + '개월 월 영업이익의 직선 추세(trend) 또는 평균(flat).',
    '- 손익현황 비용에 감가상각비가 ' + (cfg.depInPL ? '포함되어 있다고' : '포함되어 있지 않다고') + ' 가정합니다.',
    '- 원복비: 같은 입지 유형·철거 범위 과거 사례 ' + cfg.minCases + '건 이상이면 그 평당 단가 × 평수, 부족하면 전체 사례 기준.',
    '',
    '[현재 설정]',
    '- 계산 기준일: ' + partsToString_(r.ctx.baseDate) + ' (이번 달 = ' + r.ctx.baseDate.y + '-' + pad2_(r.ctx.baseDate.m) + ')',
    '- 잔여임차료 반영률: ' + Math.round(cfg.rentRate * 100) + '%, 원복비 수준: ' + levelName,
    '- 점포 ' + models.length + '개, 브랜드: ' + brands.join(', ')
  ].join('\n');
  return { r: r, models: models, system: system,
    defaults: { rate: cfg.rentRate, level: cfg.level, method: 'trend' } };
}

// 챗봇 도구 실행
function runChatTool_(name, input, data) {
  const optsOf = function () {
    const o = { rate: data.defaults.rate, level: data.defaults.level, method: data.defaults.method };
    if (typeof input.rent_rate_percent === 'number') o.rate = Math.min(Math.max(input.rent_rate_percent, 0), 100) / 100;
    if (['low', 'mid', 'high'].indexOf(input.restore_level) !== -1) o.level = input.restore_level;
    if (['trend', 'flat'].indexOf(input.forecast) !== -1) o.method = input.forecast;
    return o;
  };
  const assumptions = function (o) {
    return { rent_rate_percent: Math.round(o.rate * 100), restore_level: o.level, forecast: o.method };
  };
  const brief = function (row) {
    return row ? { month: row.ym, net_value: row.value, closing_cost: row.cost, cumulative_op: row.cumOp } : null;
  };

  if (name === 'list_stores') {
    const o = optsOf();
    return data.models
      .filter(function (m) { return !input.brand || m.brand === str_(input.brand); })
      .filter(function (m) { return !input.review_only || m.review === 'Y'; })
      .map(function (m) {
        const ev = evaluateClosing_(m, o);
        return {
          brand: m.brand, store: m.store, review: m.review,
          cumulative_op: m.status.cumOP, recent3_avg_op: m.status.recent3,
          op_trend_per_month: m.trend.slope, total_loss_now: m.status.totalLossNow,
          lease_end: m.leaseEndYm, best_month: ev.rows.length ? ev.rows[ev.best].ym : null,
          recommendation: ev.recommendation, gain_vs_lease_end: ev.gainVsLast, gain_vs_now: ev.gainVsNow,
          grade: m.grade, data_check: m.check
        };
      });
  }

  if (name === 'get_store_detail') {
    const m = findStore_(data.models, input.store, input.brand);
    return {
      brand: m.brand, store: m.store, review: m.review, area_pyeong: m.area,
      status: m.status, trend: m.trend, recent_months: m.history.slice(-6),
      restore_estimate: m.fixed.restore, other_closing_cost: m.fixed.other, penalty: m.fixed.penalty,
      restore_basis: m.basis, judgement: m.judge, grade: m.grade, data_check: m.check,
      lease_end: m.leaseEndYm, planned_close_month: m.plannedYm
    };
  }

  if (name === 'simulate_closing') {
    const m = findStore_(data.models, input.store, input.brand);
    const o = optsOf();
    const ev = evaluateClosing_(m, o);
    const ym = toYm_(input.close_month, data.r.ctx.tz);
    const idx = ev.rows.map(function (x) { return x.ym; }).indexOf(ym);
    if (idx === -1) {
      throw new Error('폐점 월은 ' + (ev.rows.length ? ev.rows[0].ym + ' ~ ' + ev.rows[ev.last].ym : '없음') +
        ' 범위에서 고를 수 있습니다. (이번 달 ~ 계약 종료월)');
    }
    const row = ev.rows[idx];
    return {
      store: storeLabel_(m.brand, m.store), close_month: ym, assumptions: assumptions(o),
      closing_cost: { total: row.cost, remaining_rent: row.rent, sunk_residual_asset: row.residual,
        restore: row.restore, other: row.other, penalty: row.penalty },
      cumulative_op_until_close: row.cumOp, net_value: row.value,
      best: brief(ev.rows[ev.best]), lease_end_operation: brief(ev.rows[ev.last]), immediate: brief(ev.rows[0]),
      diff_vs_best: row.value - ev.rows[ev.best].value, recommendation: ev.recommendation
    };
  }

  if (name === 'find_optimal_closing') {
    const m = findStore_(data.models, input.store, input.brand);
    const o = optsOf();
    const ev = evaluateClosing_(m, o);
    return {
      store: storeLabel_(m.brand, m.store), assumptions: assumptions(o),
      trend: m.trend, recommendation: ev.recommendation,
      best: brief(ev.rows[ev.best]), immediate: brief(ev.rows[0]), lease_end_operation: brief(ev.rows[ev.last]),
      gain_vs_now: ev.gainVsNow, gain_vs_lease_end: ev.gainVsLast,
      // 3개월 간격 순손익 흐름 (곡선 모양 파악용)
      net_value_curve: ev.rows.filter(function (x, i) { return i % 3 === 0 || i === ev.last; }).map(brief)
    };
  }

  if (name === 'get_brand_closing_actuals') {
    const act = data.r.bench.actuals;
    const pick = function (label, g) {
      return { brand: label, closed_stores: g.n, restore_total: g.restore, sunk_total: g.sunk,
        penalty_total: g.penalty, restore_avg: g.restore / g.n, sunk_avg: g.sunk / g.n,
        settled: g.settled, open: g.open };
    };
    const list = Object.keys(act.byBrand)
      .filter(function (b) { return !input.brand || b === str_(input.brand); })
      .map(function (b) { return pick(b, act.byBrand[b]); });
    if (!input.brand && act.all.n) list.push(pick('전체', act.all));
    if (!list.length) throw new Error('해당 브랜드의 폐점 정산 실적이 없습니다. 있는 브랜드: ' + Object.keys(act.byBrand).join(', '));
    return list;
  }

  if (name === 'get_restore_benchmark') {
    const b = data.r.bench;
    const groups = Object.keys(b.groups).map(function (k) {
      const g = b.groups[k];
      return { type: k.replace('|', ' · '), cases: g.values.length,
        per_pyeong_low: g.stats.min, per_pyeong_mid: g.stats.median, per_pyeong_high: g.stats.max };
    });
    return { groups: groups, all_cases: b.all.values.length,
      all: b.all.stats, other_closing_cost_avg: b.otherAvg, min_cases: data.r.ctx.cfg.minCases };
  }

  throw new Error('알 수 없는 도구: ' + name);
}

// 점포 표시 이름: 점포명이 이미 브랜드로 시작하면(예: PB 양재본점) 브랜드를 다시 붙이지 않음
function storeLabel_(brand, store) {
  return str_(store).toUpperCase().indexOf(str_(brand).toUpperCase() + ' ') === 0 ? store : brand + ' ' + store;
}

// 점포명으로 점포 찾기 (공백 무시, 일부 일치 허용)
function findStore_(models, store, brand) {
  const norm = function (s) { return str_(s).replace(/\s/g, '').toLowerCase(); };
  const q = norm(store);
  const pool = models.filter(function (m) { return !brand || norm(m.brand) === norm(brand); });
  let hits = pool.filter(function (m) { return norm(m.store) === q; });
  if (!hits.length) hits = pool.filter(function (m) { return norm(m.store).indexOf(q) !== -1 || q.indexOf(norm(m.store)) !== -1; });
  if (hits.length === 1) return hits[0];
  const names = (hits.length ? hits : pool).slice(0, 30).map(function (m) { return storeLabel_(m.brand, m.store); });
  throw new Error((hits.length ? '여러 점포가 해당됩니다: ' : '점포를 찾을 수 없습니다. 등록된 점포: ') + names.join(', '));
}

// 도구 결과의 숫자를 보기 좋게 반올림 (금액은 정수, 소수는 둘째 자리)
function roundDeep_(v) {
  if (typeof v === 'number') return Math.abs(v) >= 100 ? Math.round(v) : Math.round(v * 100) / 100;
  if (Array.isArray(v)) return v.map(roundDeep_);
  if (v && typeof v === 'object') {
    const o = {};
    Object.keys(v).forEach(function (k) { o[k] = roundDeep_(v[k]); });
    return o;
  }
  return v;
}


// ============================================================================
// 대시보드·챗봇 화면 (HTML)
//   - 메뉴 [📊 대시보드 열기] / [💬 챗봇 열기]에서 쓰는 화면입니다.
//   - 이 아래는 수정하지 않아도 됩니다. (별도 HTML 파일을 만들 필요가 없습니다)
// ============================================================================

const DASHBOARD_HTML_ = `<!DOCTYPE html>
<html lang="ko">
<head>
<base target="_top">
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>직영점 폐점 분석 대시보드</title>
<!-- 차트 라이브러리 (Chart.js) -->
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.js"></script>
<style>
  /* ── 색상 (밝은 화면) ── */
  :root {
    color-scheme: light;
    --surface: #f6f5f2;
    --card: #ffffff;
    --card-2: #faf9f7;
    --border: #e3e2de;
    --text: #0b0b0b;
    --text-2: #52514e;
    --text-3: #7a7974;
    --accent: #2a78d6;
    --accent-soft: #e6f0fb;
    --s1: #2a78d6;  /* 잔여임차료 / 영업이익 / 순손익 */
    --s2: #eb6834;  /* 매몰비 */
    --s3: #1baf7a;  /* 원복비 */
    --s4: #eda100;  /* 기타·위약금 */
    --good: #0ca30c;
    --warning: #fab219;
    --serious: #ec835a;
    --critical: #d03b3b;
    --good-text: #006300;
    --bad-text: #b42323;
    --grid: #ecebe7;
    --axis: #b9b8b2;
  }
  /* ── 색상 (어두운 화면) ── */
  @media (prefers-color-scheme: dark) {
    :root {
      color-scheme: dark;
      --surface: #141413;
      --card: #1e1e1d;
      --card-2: #252524;
      --border: #383835;
      --text: #f0efec;
      --text-2: #c3c2b7;
      --text-3: #94938a;
      --accent: #3987e5;
      --accent-soft: #1d2a3b;
      --s1: #3987e5;
      --s2: #d95926;
      --s3: #199e70;
      --s4: #c98500;
      --good-text: #5fcf5f;
      --bad-text: #f08080;
      --grid: #2c2c2a;
      --axis: #55554f;
    }
  }

  * { box-sizing: border-box; }
  html, body { margin: 0; background: var(--surface); color: var(--text); }
  body { font: 14px/1.5 -apple-system, "Segoe UI", "Malgun Gothic", "Apple SD Gothic Neo", sans-serif; }
  button, input, select, textarea { font: inherit; color: inherit; }
  .num { font-variant-numeric: tabular-nums; }
  .muted { color: var(--text-3); }
  .neg { color: var(--bad-text); }
  .pos { color: var(--good-text); }

  /* ── 상단 ── */
  header.top { display: flex; align-items: center; gap: 12px; padding: 14px 20px 10px; flex-wrap: wrap; }
  header.top h1 { font-size: 18px; margin: 0; font-weight: 700; }
  header.top .meta { color: var(--text-3); font-size: 12.5px; }
  header.top .spacer { flex: 1; }
  .btn { border: 1px solid var(--border); background: var(--card); border-radius: 8px; padding: 6px 12px; cursor: pointer; }
  .btn:hover { border-color: var(--accent); }
  .btn.primary { background: var(--accent); border-color: var(--accent); color: #fff; }
  .btn:disabled { opacity: .5; cursor: default; }
  a.btn { text-decoration: none; color: var(--text); display: inline-block; }
  a.btn[hidden] { display: none; }

  /* ── 탭 ── */
  nav.tabs { display: flex; gap: 4px; padding: 0 20px; border-bottom: 1px solid var(--border); }
  nav.tabs button { border: 0; background: none; padding: 10px 14px; cursor: pointer; color: var(--text-2);
    border-bottom: 2px solid transparent; margin-bottom: -1px; font-weight: 600; }
  nav.tabs button[aria-selected="true"] { color: var(--text); border-bottom-color: var(--accent); }
  section.tab { display: none; padding: 16px 20px 28px; }
  section.tab.active { display: block; }

  /* ── 조건 선택 줄 ── */
  .bar { display: flex; flex-wrap: wrap; gap: 10px 16px; align-items: flex-end; background: var(--card);
    border: 1px solid var(--border); border-radius: 12px; padding: 12px 14px; margin-bottom: 14px; }
  .field { display: flex; flex-direction: column; gap: 4px; min-width: 120px; }
  .field label { font-size: 12px; color: var(--text-3); font-weight: 600; }
  .field select, .field input[type="text"], .field input[list] { border: 1px solid var(--border); background: var(--card-2);
    border-radius: 8px; padding: 6px 8px; min-width: 120px; }
  .field.wide input { min-width: 240px; }
  .field .range { display: flex; align-items: center; gap: 8px; }
  .field .range input { width: 140px; accent-color: var(--accent); }
  .bar .sep { width: 1px; align-self: stretch; background: var(--border); }
  .bar .hint { font-size: 12px; color: var(--text-3); align-self: center; }

  /* ── 카드·격자 ── */
  .card { background: var(--card); border: 1px solid var(--border); border-radius: 12px; padding: 14px 16px; }
  .card h2 { font-size: 14px; margin: 0 0 2px; }
  .card .sub { font-size: 12px; color: var(--text-3); margin: 0 0 10px; }
  .grid { display: grid; gap: 14px; margin-bottom: 14px; }
  .g-sim { grid-template-columns: minmax(300px, 380px) 1fr; }
  .g-2 { grid-template-columns: 1fr 1fr; }
  @media (max-width: 980px) { .g-sim, .g-2 { grid-template-columns: 1fr; } }
  .chart-box { position: relative; height: 270px; }

  /* ── 권고 배너 ── */
  .reco { display: flex; gap: 14px; align-items: center; background: var(--accent-soft); border: 1px solid var(--border);
    border-radius: 12px; padding: 12px 16px; margin-bottom: 14px; }
  .reco .icon { font-size: 22px; }
  .reco .title { font-weight: 700; font-size: 16px; }
  .reco .desc { color: var(--text-2); font-size: 13px; }

  /* ── 현황 타일 ── */
  .kpis { display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap: 10px; margin-bottom: 14px; }
  .kpi { background: var(--card); border: 1px solid var(--border); border-radius: 12px; padding: 10px 12px; }
  .kpi .k { font-size: 12px; color: var(--text-3); font-weight: 600; }
  .kpi .v { font-size: 19px; font-weight: 700; margin-top: 2px; }
  .kpi .s { font-size: 12px; color: var(--text-3); }
  .badge { display: inline-flex; align-items: center; gap: 6px; font-weight: 700; }
  .dot { width: 9px; height: 9px; border-radius: 50%; display: inline-block; }

  /* ── 선택 월 폐점 카드 ── */
  .hero { font-size: 28px; font-weight: 800; letter-spacing: -0.5px; }
  .hero small { font-size: 14px; font-weight: 600; color: var(--text-2); }
  .eok { color: var(--text-3); font-size: 12.5px; }
  .rows { margin: 10px 0 0; padding: 0; list-style: none; }
  .rows li { display: flex; align-items: center; gap: 8px; padding: 5px 0; border-bottom: 1px dashed var(--border); }
  .rows li:last-child { border-bottom: 0; }
  .rows .sw { width: 10px; height: 10px; border-radius: 3px; flex: none; }
  .rows .l { flex: 1; color: var(--text-2); }
  .rows .r { font-weight: 600; }
  .rows li.total { border-top: 1px solid var(--border); border-bottom: 0; margin-top: 4px; padding-top: 8px; }
  .rows li.total .l { color: var(--text); font-weight: 700; }

  /* ── 상세 정보 표 ── */
  .info { display: grid; grid-template-columns: repeat(auto-fit, minmax(260px, 1fr)); gap: 4px 24px; font-size: 13px; }
  .info div { display: flex; gap: 10px; padding: 4px 0; border-bottom: 1px solid var(--border); }
  .info dt { color: var(--text-3); min-width: 110px; }
  .info dd { margin: 0; flex: 1; }

  /* ── 전체 점포 표 ── */
  .table-wrap { overflow: auto; border: 1px solid var(--border); border-radius: 12px; background: var(--card); }
  table.list { border-collapse: collapse; width: 100%; font-size: 13px; }
  table.list th, table.list td { padding: 8px 10px; border-bottom: 1px solid var(--border); text-align: right; white-space: nowrap; }
  table.list th { position: sticky; top: 0; background: var(--card-2); color: var(--text-2); font-weight: 600; font-size: 12px; }
  table.list td.t, table.list th.t { text-align: left; }
  table.list tbody tr { cursor: pointer; }
  table.list tbody tr:hover { background: var(--accent-soft); }

  /* ── 챗봇 ── */
  .chat { display: flex; flex-direction: column; height: calc(100vh - 120px); min-height: 420px; max-width: 980px; margin: 0 auto; }
  .chat-log { flex: 1; overflow-y: auto; padding: 4px 2px 12px; display: flex; flex-direction: column; gap: 10px; }
  .msg { max-width: 88%; padding: 10px 13px; border-radius: 14px; line-height: 1.6; word-break: break-word; }
  .msg.user { align-self: flex-end; background: var(--accent); color: #fff; border-bottom-right-radius: 4px; white-space: pre-wrap; }
  .msg.bot { align-self: flex-start; background: var(--card); border: 1px solid var(--border); border-bottom-left-radius: 4px; }
  .msg.bot p { margin: 0 0 8px; } .msg.bot p:last-child { margin-bottom: 0; }
  .msg.bot ul, .msg.bot ol { margin: 4px 0 8px; padding-left: 20px; }
  .msg.bot h3, .msg.bot h4 { margin: 8px 0 4px; font-size: 14px; }
  .msg.bot table { border-collapse: collapse; margin: 6px 0; font-size: 12.5px; display: block; overflow-x: auto; }
  .msg.bot th, .msg.bot td { border: 1px solid var(--border); padding: 4px 8px; white-space: nowrap; }
  .msg.bot th { background: var(--card-2); }
  .msg.bot code { background: var(--card-2); padding: 1px 4px; border-radius: 4px; }
  .msg.note { align-self: center; background: none; color: var(--text-3); font-size: 12.5px; padding: 2px; }
  .msg.err { align-self: center; background: none; color: var(--bad-text); font-size: 13px; }
  .tools { display: flex; flex-wrap: wrap; gap: 6px; margin-bottom: 6px; }
  .tool-chip { font-size: 11.5px; color: var(--text-2); background: var(--card-2); border: 1px solid var(--border); border-radius: 999px; padding: 1px 8px; }
  .typing { align-self: flex-start; color: var(--text-3); font-size: 13px; padding: 6px 4px; }
  .chips { display: flex; flex-wrap: wrap; gap: 6px; margin: 6px 0 10px; }
  .chips button { border: 1px solid var(--border); background: var(--card); border-radius: 999px; padding: 5px 11px; cursor: pointer; font-size: 12.5px; color: var(--text-2); }
  .chips button:hover { border-color: var(--accent); color: var(--text); }
  .chat-input { display: flex; gap: 8px; align-items: flex-end; }
  .chat-input textarea { flex: 1; resize: none; border: 1px solid var(--border); background: var(--card); border-radius: 12px; padding: 10px 12px; min-height: 46px; max-height: 160px; }
  .chat-foot { display: flex; justify-content: space-between; gap: 8px; font-size: 11.5px; color: var(--text-3); margin-top: 6px; }
  .chat-foot button { border: 0; background: none; color: var(--accent); cursor: pointer; padding: 0; font-size: 11.5px; }

  /* ── 사이드바(챗봇만) 모드 ── */
  body.mode-chat header.top, body.mode-chat nav.tabs { display: none; }
  body.mode-chat section.tab { padding: 10px; }
  body.mode-chat .chat { height: calc(100vh - 20px); }
  body.mode-chat .msg { max-width: 96%; font-size: 13px; }
  body.mode-chat .chat-foot { flex-direction: column-reverse; align-items: flex-start; }
  .chips[hidden] { display: none; }

  .loading, .empty { padding: 60px 20px; text-align: center; color: var(--text-3); }
</style>
</head>
<body>
<header class="top">
  <h1>📊 직영점 폐점 분석 대시보드</h1>
  <span class="meta" id="meta">불러오는 중…</span>
  <span class="spacer"></span>
  <!-- 시트 주소는 코드에 넣지 않고, 실행할 때 서버에서 받아 채웁니다 -->
  <a class="btn" id="sheet-link" target="_blank" rel="noopener" hidden title="이 대시보드의 원본 데이터가 있는 구글 시트를 새 탭에서 엽니다">📄 원본 시트 열기</a>
  <button class="btn" id="reload" title="시트 데이터를 다시 불러옵니다">🔄 새로고침</button>
</header>

<nav class="tabs" role="tablist">
  <button role="tab" data-tab="store" aria-selected="true">점포 분석</button>
  <button role="tab" data-tab="all" aria-selected="false">전체 점포</button>
  <button role="tab" data-tab="chat" aria-selected="false">💬 챗봇</button>
</nav>

<!-- ───────────── 점포 분석 ───────────── -->
<section class="tab active" id="tab-store">
  <div class="loading" id="loading">시트 데이터를 계산하는 중입니다…</div>
  <div id="store-body" hidden>
    <div class="bar">
      <div class="field"><label for="brand">브랜드</label><select id="brand"></select></div>
      <div class="field wide"><label for="store">점포 (입력하면 검색)</label>
        <input id="store" list="store-list" placeholder="점포명을 입력하세요" autocomplete="off">
        <datalist id="store-list"></datalist></div>
      <div class="field"><label for="month">폐점 월</label><select id="month"></select></div>
      <div class="sep"></div>
      <div class="field"><label for="rate">잔여임차료 반영률</label>
        <div class="range"><input type="range" id="rate" min="0" max="100" step="5"><span id="rate-v" class="num"></span></div></div>
      <div class="field"><label for="level">원복비 수준</label>
        <select id="level"><option value="low">적게 (최소)</option><option value="mid">보통 (중앙값)</option><option value="high">많이 (최대)</option></select></div>
      <div class="field"><label for="method">손익 전망</label>
        <select id="method"><option value="trend">최근 추세 연장</option><option value="flat">최근 평균 유지</option></select></div>
      <button class="btn" id="ask-bot" title="이 점포에 대해 챗봇에게 묻기">💬 이 점포 질문하기</button>
    </div>

    <div class="reco" id="reco"></div>
    <div class="kpis" id="kpis"></div>

    <div class="grid g-sim">
      <div class="card" id="sim"></div>
      <div class="card">
        <h2>폐점 시점별 순손익</h2>
        <p class="sub">이번 달부터 그 달까지 예상 영업이익 − 그 달 폐점 비용 · 가장 높은 달이 최적 폐점 시점 (천원)</p>
        <div class="chart-box"><canvas id="c-value" aria-label="폐점 시점별 순손익 차트"></canvas></div>
      </div>
    </div>

    <div class="grid g-2">
      <div class="card">
        <h2>월별 영업이익 — 실적과 전망</h2>
        <p class="sub" id="trend-sub">실선 = 실적, 점선 = 전망 (천원)</p>
        <div class="chart-box"><canvas id="c-op" aria-label="월별 영업이익 차트"></canvas></div>
      </div>
      <div class="card">
        <h2>폐점 월별 예상 폐점 비용</h2>
        <p class="sub">늦게 닫을수록 남은 임차료·매몰비가 줄어듭니다 (천원)</p>
        <div class="chart-box"><canvas id="c-cost" aria-label="폐점 월별 예상 비용 차트"></canvas></div>
      </div>
    </div>

    <div class="card">
      <h2>점포 정보 · 산정 근거</h2>
      <p class="sub">데이터점검에 표시된 항목은 시트에서 보완하면 자동 반영됩니다</p>
      <dl class="info" id="info"></dl>
    </div>
  </div>
</section>

<!-- ───────────── 전체 점포 ───────────── -->
<section class="tab" id="tab-all">
  <div class="bar">
    <div class="field"><label for="all-brand">브랜드</label><select id="all-brand"></select></div>
    <div class="field"><label for="all-filter">보기</label>
      <select id="all-filter"><option value="all">전체 점포</option><option value="review">폐점검토 점포만</option><option value="early">조기 폐점이 유리한 점포</option></select></div>
    <span class="hint" id="all-hint"></span>
  </div>
  <div class="table-wrap">
    <table class="list">
      <thead><tr>
        <th class="t">브랜드</th><th class="t">점포</th><th class="t">검토</th>
        <th>누적 영업이익</th><th>최근3개월 평균</th><th>추세(월)</th><th>현재 예상총손실</th>
        <th class="t">계약종료</th><th class="t">최적 폐점</th><th class="t">권고</th><th>계약만료 대비 이득</th><th class="t">확정도</th>
      </tr></thead>
      <tbody id="all-body"></tbody>
    </table>
  </div>
  <p class="muted" style="font-size:12px">금액 단위: 천원 · 행을 누르면 점포 분석으로 이동합니다 · 가정은 점포 분석 탭의 반영률·원복비 수준·손익 전망을 따릅니다</p>
</section>

<!-- ───────────── 챗봇 ───────────── -->
<section class="tab" id="tab-chat">
  <div class="chat">
    <div class="chat-log" id="chat-log" aria-live="polite"></div>
    <div class="chips" id="chips"></div>
    <div class="chat-input">
      <textarea id="chat-text" rows="1" placeholder="예) 양재본점은 언제 폐점하는 게 좋아?  (Enter 보내기, Shift+Enter 줄바꿈)"></textarea>
      <button class="btn primary" id="chat-send">보내기</button>
    </div>
    <div class="chat-foot">
      <span>질문과 조회 결과는 Google Gemini API로 전송됩니다 · 숫자는 시트 데이터를 조회해 계산합니다</span>
      <button id="chat-new">새 대화</button>
    </div>
  </div>
</section>

<script>
var MODE = <?!= JSON.stringify(mode) ?>;

// ── 서버(Code.gs)와 같은 계산식: 폐점 시점별 비용·순손익 ──
<?!= sharedJs ?>

(function () {
  'use strict';
  document.body.classList.add('mode-' + MODE);

  var state = { data: null, models: [], key: null, month: null,
    opts: { rate: 1, level: 'mid', method: 'trend' }, charts: {}, chat: [], busy: false };

  // ───────── 공통 도우미 ─────────
  var $ = function (id) { return document.getElementById(id); };
  function esc(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }
  function fmt(n) { return (n == null || isNaN(n)) ? '-' : Math.round(n).toLocaleString('ko-KR'); }
  function fmtS(n) { return (n > 0 ? '+' : '') + fmt(n); }
  function cls(n) { return n < 0 ? 'neg' : ''; }
  // 천원 → "약 1.2억원" / "약 3,400만원"
  function eok(n) {
    if (n == null || isNaN(n)) return '';
    var a = Math.abs(n), s = n < 0 ? '−' : '';
    if (a >= 100000) return '약 ' + s + (a / 100000).toFixed(1) + '억원';
    if (a >= 10) return '약 ' + s + Math.round(a / 10).toLocaleString('ko-KR') + '만원';
    return '';
  }
  // 차트 축: 천원 값을 억/만 단위로 짧게
  function axisFmt(v) {
    var a = Math.abs(v), s = v < 0 ? '-' : '';
    if (a >= 100000) return s + +(a / 100000).toFixed(1) + '억';
    if (a >= 10) return s + Math.round(a / 10).toLocaleString('ko-KR') + '만';
    return String(v);
  }
  function css(name) { return getComputedStyle(document.documentElement).getPropertyValue(name).trim(); }
  function levelName(l) { return { low: '적게(최소)', mid: '보통(중앙값)', high: '많이(최대)' }[l]; }
  function gradeBadge(g) {
    if (!g) return '<span class="muted">미검토</span>';
    var c = g.charAt(0) === 'A' ? 'var(--good)' : g.charAt(0) === 'B' ? 'var(--warning)' : 'var(--serious)';
    return '<span class="badge"><span class="dot" style="background:' + c + '"></span>' + esc(g.split('·')[0].trim()) + '</span>';
  }
  function model() { return state.models.filter(function (m) { return m.key === state.key; })[0]; }

  // ───────── 탭 ─────────
  function showTab(name) {
    Array.prototype.forEach.call(document.querySelectorAll('nav.tabs button'), function (b) {
      b.setAttribute('aria-selected', b.getAttribute('data-tab') === name ? 'true' : 'false');
    });
    Array.prototype.forEach.call(document.querySelectorAll('section.tab'), function (s) {
      s.classList.toggle('active', s.id === 'tab-' + name);
    });
    if (name === 'all') renderAll();
    if (name === 'chat') $('chat-text').focus();
  }
  Array.prototype.forEach.call(document.querySelectorAll('nav.tabs button'), function (b) {
    b.addEventListener('click', function () { showTab(b.getAttribute('data-tab')); });
  });

  // ───────── 데이터 불러오기 ─────────
  function load() {
    $('loading').hidden = false; $('store-body').hidden = true;
    $('loading').textContent = '시트 데이터를 계산하는 중입니다…';
    google.script.run
      .withSuccessHandler(function (json) { onData(JSON.parse(json)); })
      .withFailureHandler(function (err) { $('loading').textContent = '불러오지 못했습니다: ' + err.message; })
      .getDashboardData();
  }

  function onData(d) {
    state.data = d;
    state.models = d.models;
    state.opts = { rate: d.defaults.rate, level: d.defaults.level, method: d.defaults.method };
    $('meta').textContent = '계산 기준일 ' + d.asOf + ' · 금액 단위 천원 · 갱신 ' + d.updated;
    if (d.sheetUrl) { $('sheet-link').href = d.sheetUrl; $('sheet-link').hidden = false; }
    if (!state.models.length) {
      $('loading').textContent = '점포 데이터가 없습니다. 시트 메뉴 [🛠 시트 만들기] 후 데이터를 입력하고 다시 열어 주세요.';
      return;
    }
    $('rate').value = Math.round(state.opts.rate * 100);
    $('level').value = state.opts.level;
    $('method').value = state.opts.method;

    var brands = state.models.map(function (m) { return m.brand; }).filter(function (b, i, a) { return a.indexOf(b) === i; });
    $('brand').innerHTML = '<option value="">전체 브랜드</option>' + brands.map(function (b) { return '<option>' + esc(b) + '</option>'; }).join('');
    $('all-brand').innerHTML = $('brand').innerHTML;
    fillStoreList();

    // 처음 보여줄 점포: 이전 선택 → 폐점검토 점포 → 첫 점포
    var keep = state.models.filter(function (m) { return m.key === state.key; })[0];
    var first = keep || state.models.filter(function (m) { return m.review === 'Y'; })[0] || state.models[0];
    selectStore(first.key, true);
    $('loading').hidden = true; $('store-body').hidden = false;
    renderAll();
  }

  function fillStoreList() {
    var b = $('brand').value;
    $('store-list').innerHTML = state.models
      .filter(function (m) { return !b || m.brand === b; })
      .map(function (m) { return '<option value="' + esc(m.brand + ' · ' + m.store) + '">'; }).join('');
  }

  function selectStore(key, resetMonth) {
    state.key = key;
    var m = model();
    $('store').value = m.brand + ' · ' + m.store;
    if (resetMonth) state.month = null;
    renderStore();
  }

  // ───────── 조건 변경 ─────────
  $('brand').addEventListener('change', function () {
    fillStoreList();
    var b = $('brand').value;
    var m = state.models.filter(function (x) { return !b || x.brand === b; })[0];
    if (m) selectStore(m.key, true);
  });
  $('store').addEventListener('change', pickTyped);
  $('store').addEventListener('input', pickTyped);
  function pickTyped() {
    var v = $('store').value.trim();
    var m = state.models.filter(function (x) { return x.brand + ' · ' + x.store === v; })[0];
    if (m && m.key !== state.key) selectStore(m.key, true);
  }
  $('store').addEventListener('focus', function () { this.select(); });
  $('month').addEventListener('change', function () { state.month = this.value; renderStore(); });
  $('rate').addEventListener('input', function () { state.opts.rate = this.value / 100; renderStore(); });
  $('level').addEventListener('change', function () { state.opts.level = this.value; renderStore(); });
  $('method').addEventListener('change', function () { state.opts.method = this.value; renderStore(); });
  $('reload').addEventListener('click', load);
  $('ask-bot').addEventListener('click', function () {
    var m = model();
    showTab('chat');
    $('chat-text').value = m.brand + ' ' + m.store + '은(는) 언제 폐점하는 게 가장 좋을까? 근거와 가정, 결론이 바뀔 수 있는 조건도 알려줘.';
    autoGrow();
  });
  $('all-brand').addEventListener('change', renderAll);
  $('all-filter').addEventListener('change', renderAll);

  // ───────── 점포 분석 화면 ─────────
  function renderStore() {
    var m = model();
    if (!m) return;
    var ev = evaluateClosing_(m, state.opts);
    $('rate-v').textContent = Math.round(state.opts.rate * 100) + '%';

    if (!ev.rows.length) {
      $('reco').innerHTML = '<span class="icon">⚠️</span><div><div class="title">계산할 기간이 없습니다</div><div class="desc">임차종료일이 이미 지났는지 [자산투자비]를 확인해 주세요.</div></div>';
      return;
    }
    // 폐점 월 선택: 이전 선택 → 폐점예정월 → 최적 시점
    var yms = ev.rows.map(function (r) { return r.ym; });
    if (yms.indexOf(state.month) === -1) {
      state.month = yms.indexOf(m.plannedYm) !== -1 ? m.plannedYm : yms[ev.best];
    }
    $('month').innerHTML = ev.rows.map(function (r, i) {
      var tags = [];
      if (i === 0) tags.push('이번 달');
      if (i === ev.best) tags.push('최적');
      if (r.ym === m.plannedYm) tags.push('폐점예정');
      if (i === ev.last && m.leaseEndYm === r.ym) tags.push('계약 만료');
      return '<option value="' + r.ym + '"' + (r.ym === state.month ? ' selected' : '') + '>' +
        r.ym + (tags.length ? ' (' + tags.join(', ') + ')' : '') + '</option>';
    }).join('');
    var si = yms.indexOf(state.month);
    var sel = ev.rows[si], best = ev.rows[ev.best], last = ev.rows[ev.last], now = ev.rows[0];

    // 권고 배너
    var icon = ev.best === 0 && ev.last > 0 ? '🚩' : ev.best === ev.last ? '✅' : '📅';
    var desc = [];
    if (ev.last > 0) {
      desc.push('계약 만료(' + last.ym + ')까지 운영 대비 <b class="num">' + fmtS(ev.gainVsLast) + '</b>천원');
      desc.push('즉시 폐점(' + now.ym + ') 대비 <b class="num">' + fmtS(ev.gainVsNow) + '</b>천원');
    }
    desc.push('가정: 잔여임차료 ' + Math.round(state.opts.rate * 100) + '% 반영 · 원복비 ' + levelName(state.opts.level) +
      ' · ' + (state.opts.method === 'flat' ? '최근 평균 유지' : '최근 추세 연장'));
    $('reco').innerHTML = '<span class="icon">' + icon + '</span><div><div class="title">최적 폐점 시점: ' + best.ym +
      ' — ' + esc(ev.recommendation) + '</div><div class="desc">' + desc.join(' · ') + '</div></div>';

    // 현황 타일
    var st = m.status, tr = m.trend;
    var trendTxt = tr.n >= 3 ? fmtS(tr.slope) : '-';
    var r2 = tr.r2 == null ? '추세 계산 불가 (3개월 미만)' : 'R² ' + tr.r2.toFixed(2) + (tr.r2 < 0.3 ? ' · 추세 불확실' : '');
    $('kpis').innerHTML = [
      kpi('누적 매출', fmt(st.cumSales), (st.firstYm ? st.firstYm + ' ~ ' + st.lastYm + ' · ' : '') + st.months + '개월'),
      kpi('누적 영업이익', '<span class="' + cls(st.cumOP) + '">' + fmt(st.cumOP) + '</span>', '이익률 ' + (st.margin == null ? '-' : (st.margin * 100).toFixed(1) + '%')),
      kpi('최근 3개월 평균 영업이익', '<span class="' + cls(st.recent3) + '">' + fmt(st.recent3) + '</span>', '최근월 ' + fmt(st.lastOp)),
      kpi('영업이익 추세 (월)', '<span class="' + cls(tr.slope) + '">' + trendTxt + '</span>', r2),
      kpi('잔존자산가치 (매몰비)', fmt(st.residualNow), st.baseDate + ' 기준'),
      kpi('잔여 임차', fmt(st.remainRent), (st.remainMonths || 0).toFixed(1) + '개월 · 월 ' + fmt(st.rent)),
      kpi('확정도 등급', gradeBadge(m.grade), m.review === 'Y' ? '폐점검토 점포' : '폐점검토 미등록')
    ].join('');

    // 선택 월 폐점 카드
    var diff = sel.value - best.value;
    $('sim').innerHTML =
      '<h2>' + sel.ym + ' 말 폐점 시</h2><p class="sub">' + (si === ev.best ? '✓ 최적 폐점 시점입니다' : '최적(' + best.ym + ') 대비 순손익 <b class="num neg">' + fmtS(diff) + '</b>천원') + '</p>' +
      '<div class="hero num">' + fmt(sel.cost) + ' <small>천원</small></div><div class="eok">예상 폐점 비용 합계 · ' + eok(sel.cost) + '</div>' +
      '<ul class="rows num">' +
        row('var(--s1)', '잔여임차료 × ' + Math.round(state.opts.rate * 100) + '%', sel.rent) +
        row('var(--s2)', '매몰비 (잔존자산가치)', sel.residual) +
        row('var(--s3)', '원복비 (' + levelName(state.opts.level) + ')', sel.restore) +
        row('var(--s4)', '기타 폐점비 + 위약금', sel.other + sel.penalty) +
      '</ul>' +
      '<ul class="rows num">' +
        row(null, '이번 달 ~ ' + sel.ym + ' 예상 영업이익', sel.cumOp) +
        row(null, '순손익 (영업이익 − 폐점 비용)', sel.value, true) +
      '</ul>';

    // 점포 정보
    $('info').innerHTML = [
      info('브랜드 · 점포', m.brand + ' · ' + m.store),
      info('평수', m.area ? m.area + '평' : '미입력'),
      info('오픈(투자)일', st.openDate || '미입력'),
      info('임차 종료일', st.leaseEndDate || '미입력'),
      info('월 임차료', fmt(st.rent) + '천원'),
      info('보증금 (회수 예정)', fmt(st.deposit) + '천원'),
      info('원복비 범위', fmt(m.fixed.restore.low) + ' ~ ' + fmt(m.fixed.restore.high) + '천원 (보통 ' + fmt(m.fixed.restore.mid) + ')'),
      info('원복비 적용기준', m.basis || '-'),
      info('판단 (주의 항목)', m.judge || '-'),
      info('브랜드 과거 건당 원상복구', st.pastRestore == null ? '실적 없음' : fmt(st.pastRestore) + '천원'),
      info('브랜드 과거 건당 매몰비', st.pastSunk == null ? '실적 없음' : fmt(st.pastSunk) + '천원'),
      info('현재 예상총손실액', fmt(st.totalLossNow) + '천원 (' + st.baseDate + ' 폐점 기준)'),
      info('폐점예정월', m.plannedYm || '-'),
      info('데이터점검', m.check === 'OK' ? '✓ 이상 없음' : '<span class="neg">⚠ ' + esc(m.check) + '</span>', true)
    ].join('');
    $('trend-sub').textContent = '실선 = 실적, 점선 = 전망 (' + (state.opts.method === 'flat' ? '최근 ' + tr.n + '개월 평균 유지' : '최근 ' + tr.n + '개월 추세 연장') + ') · 천원';

    drawCharts(m, ev, si);
  }
  function kpi(k, v, s) { return '<div class="kpi"><div class="k">' + k + '</div><div class="v num">' + v + '</div><div class="s">' + esc(s) + '</div></div>'; }
  function row(color, label, v, total) {
    return '<li' + (total ? ' class="total"' : '') + '>' + (color ? '<span class="sw" style="background:' + color + '"></span>' : '') +
      '<span class="l">' + label + '</span><span class="r ' + cls(v) + '">' + fmt(v) + '</span></li>';
  }
  function info(k, v, raw) { return '<div><dt>' + k + '</dt><dd>' + (raw ? v : esc(v)) + '</dd></div>'; }

  // ───────── 차트 ─────────
  // 선택 월·최적 월에 세로선을 그리는 작은 플러그인
  var markerPlugin = {
    id: 'markers',
    afterDatasetsDraw: function (chart, args, opts) {
      var items = opts.items || [], ctx = chart.ctx, area = chart.chartArea, x = chart.scales.x;
      items.forEach(function (it, n) {
        if (it.index == null || it.index < 0) return;
        var px = x.getPixelForValue(it.index);
        ctx.save();
        ctx.strokeStyle = it.color; ctx.lineWidth = 1.5; ctx.setLineDash(it.dash || []);
        ctx.beginPath(); ctx.moveTo(px, area.top); ctx.lineTo(px, area.bottom); ctx.stroke();
        ctx.setLineDash([]);
        ctx.font = '600 11px sans-serif'; ctx.fillStyle = it.color;
        ctx.textAlign = px > area.right - 50 ? 'right' : 'left';
        ctx.fillText(it.label, px + (ctx.textAlign === 'right' ? -4 : 4), area.top + 11 + n * 13);
        ctx.restore();
      });
    }
  };

  function baseOptions(yTitle) {
    var grid = css('--grid'), axis = css('--axis'), text = css('--text-2');
    return {
      responsive: true, maintainAspectRatio: false, animation: { duration: 250 },
      interaction: { mode: 'index', intersect: false },
      plugins: {
        legend: { position: 'top', align: 'start', labels: { color: text, boxWidth: 10, boxHeight: 10, usePointStyle: true, pointStyle: 'rectRounded' } },
        tooltip: { callbacks: { label: function (c) { return ' ' + c.dataset.label + ': ' + fmt(c.parsed.y) + '천원'; } } }
      },
      scales: {
        x: { grid: { display: false }, border: { color: axis }, ticks: { color: text, maxRotation: 0, autoSkipPadding: 14 } },
        y: { grid: { color: function (c) { return c.tick && c.tick.value === 0 ? axis : grid; } }, border: { display: false },
          ticks: { color: text, callback: axisFmt }, title: { display: !!yTitle, text: yTitle, color: text } }
      }
    };
  }

  function drawCharts(m, ev, si) {
    if (typeof Chart === 'undefined') return;
    var s1 = css('--s1'), s2 = css('--s2'), s3 = css('--s3'), s4 = css('--s4'), card = css('--card'), t3 = css('--text-3');
    var rows = ev.rows;
    var labels = rows.map(function (r) { return r.ym; });
    var marks = [{ index: ev.best, label: '최적 ' + rows[ev.best].ym, color: css('--good-text'), dash: [4, 3] }];
    if (si !== ev.best) marks.push({ index: si, label: '선택 ' + rows[si].ym, color: css('--text') });

    // (1) 순손익 곡선
    var o1 = baseOptions();
    o1.plugins.legend.display = false;
    o1.plugins.markers = { items: marks };
    o1.plugins.tooltip.callbacks.afterBody = function (items) {
      var r = rows[items[0].dataIndex];
      return ['예상 영업이익 누계 ' + fmt(r.cumOp) + '천원', '폐점 비용 ' + fmt(r.cost) + '천원'];
    };
    upsert('value', 'line', {
      labels: labels,
      datasets: [{ label: '순손익', data: rows.map(function (r) { return r.value; }), borderColor: s1, backgroundColor: s1, borderWidth: 2,
        pointRadius: rows.map(function (r, i) { return i === ev.best ? 6 : i === si ? 5 : 0; }),
        pointHoverRadius: 6, pointBorderColor: card, pointBorderWidth: 2, tension: 0.2 }]
    }, o1);

    // (2) 월별 영업이익 실적 + 전망
    var hist = m.history || [];
    var all = hist.map(function (h) { return h.ym; }).concat(labels).filter(function (v, i, a) { return a.indexOf(v) === i; }).sort();
    var hMap = {}; hist.forEach(function (h) { hMap[h.ym] = h.op; });
    var fMap = {}; rows.forEach(function (r) { fMap[r.ym] = r.op; });
    var o2 = baseOptions();
    o2.plugins.markers = { items: marks.map(function (k) { return { index: all.indexOf(rows[k.index].ym), label: k.label, color: k.color, dash: k.dash }; }) };
    upsert('op', 'line', {
      labels: all,
      datasets: [
        { label: '실적 (실선)', data: all.map(function (y) { return y in hMap ? hMap[y] : null; }), borderColor: s1, backgroundColor: s1,
          borderWidth: 2, pointRadius: 0, pointHoverRadius: 4, tension: 0.2, spanGaps: false },
        { label: '전망 (점선)', data: all.map(function (y) { return y in fMap ? fMap[y] : null; }), borderColor: s1, backgroundColor: s1,
          borderWidth: 2, borderDash: [5, 4], pointRadius: 0, pointHoverRadius: 4, tension: 0.2 }
      ]
    }, o2);

    // (3) 폐점 월별 비용 (쌓은 막대)
    var o3 = baseOptions();
    o3.scales.x.stacked = true; o3.scales.y.stacked = true;
    o3.plugins.markers = { items: marks };
    o3.plugins.tooltip.callbacks.footer = function (items) { return '합계 ' + fmt(rows[items[0].dataIndex].cost) + '천원'; };
    var bar = function (label, color, fn) {
      return { label: label, data: rows.map(fn), backgroundColor: color, borderColor: card,
        borderWidth: { top: 2, right: 0, bottom: 0, left: 0 }, borderSkipped: false, categoryPercentage: 0.82, barPercentage: 0.92 };
    };
    upsert('cost', 'bar', {
      labels: labels,
      datasets: [
        bar('잔여임차료', s1, function (r) { return r.rent; }),
        bar('매몰비', s2, function (r) { return r.residual; }),
        bar('원복비', s3, function (r) { return r.restore; }),
        bar('기타·위약금', s4, function (r) { return r.other + r.penalty; })
      ]
    }, o3);
  }

  function upsert(id, type, data, options) {
    var c = state.charts[id];
    if (c && c.config.type === type) { c.data = data; c.options = options; c.update(); return; }
    if (c) c.destroy();
    state.charts[id] = new Chart($('c-' + id), { type: type, data: data, options: options, plugins: [markerPlugin] });
  }

  // 화면 모드(밝게/어둡게)가 바뀌면 차트 색을 다시 칠합니다.
  if (window.matchMedia) {
    var mq = window.matchMedia('(prefers-color-scheme: dark)');
    var redraw = function () { Object.keys(state.charts).forEach(function (k) { state.charts[k].destroy(); }); state.charts = {}; if (state.key) renderStore(); };
    if (mq.addEventListener) mq.addEventListener('change', redraw);
  }

  // ───────── 전체 점포 표 ─────────
  function renderAll() {
    if (!state.models.length) return;
    var b = $('all-brand').value, f = $('all-filter').value;
    var list = state.models.map(function (m) { return { m: m, ev: evaluateClosing_(m, state.opts) }; })
      .filter(function (x) { return !b || x.m.brand === b; })
      .filter(function (x) { return f !== 'review' || x.m.review === 'Y'; })
      .filter(function (x) { return f !== 'early' || (x.ev.rows.length > 1 && x.ev.best < x.ev.last); });
    list.sort(function (a, c) { return a.m.brand.localeCompare(c.m.brand, 'ko') || (c.ev.gainVsLast - a.ev.gainVsLast); });
    $('all-hint').textContent = list.length + '개 점포 · 가정: 잔여임차료 ' + Math.round(state.opts.rate * 100) + '% · 원복비 ' +
      levelName(state.opts.level) + ' · ' + (state.opts.method === 'flat' ? '평균 유지' : '추세 연장');
    $('all-body').innerHTML = list.length ? list.map(function (x) {
      var m = x.m, ev = x.ev, st = m.status, has = ev.rows.length > 0;
      return '<tr data-key="' + esc(m.key) + '">' +
        '<td class="t">' + esc(m.brand) + '</td><td class="t">' + esc(m.store) + '</td><td class="t">' + (m.review === 'Y' ? '●' : '') + '</td>' +
        '<td class="num ' + cls(st.cumOP) + '">' + fmt(st.cumOP) + '</td>' +
        '<td class="num ' + cls(st.recent3) + '">' + fmt(st.recent3) + '</td>' +
        '<td class="num ' + cls(m.trend.slope) + '">' + (m.trend.n >= 3 ? fmtS(m.trend.slope) : '-') + '</td>' +
        '<td class="num">' + fmt(st.totalLossNow) + '</td>' +
        '<td class="t">' + esc(m.leaseEndYm || '-') + '</td>' +
        '<td class="t"><b>' + (has ? ev.rows[ev.best].ym : '-') + '</b></td>' +
        '<td class="t">' + esc(ev.recommendation) + '</td>' +
        '<td class="num">' + (has ? fmtS(ev.gainVsLast) : '-') + '</td>' +
        '<td class="t">' + gradeBadge(m.grade) + '</td></tr>';
    }).join('') : '<tr><td class="t muted" colspan="12">조건에 맞는 점포가 없습니다.</td></tr>';
  }
  $('all-body').addEventListener('click', function (e) {
    var tr = e.target.closest('tr[data-key]');
    if (!tr) return;
    selectStore(tr.getAttribute('data-key'), true);
    showTab('store');
  });

  // ───────── 챗봇 ─────────
  var TOOL_NAMES = { list_stores: '점포 목록 조회', get_store_detail: '점포 상세 조회', simulate_closing: '폐점 시뮬레이션',
    find_optimal_closing: '최적 폐점 시점 계산', get_brand_closing_actuals: '브랜드 폐점 실적 조회', get_restore_benchmark: '원복비 기준값 조회' };
  var SUGGEST = ['조기 폐점이 유리한 점포가 있어?', '적자 점포들의 최적 폐점 시점을 표로 정리해줘',
    'PB 브랜드 과거 폐점 비용 평균은?', '잔여임차료를 50%만 내도 된다면 결론이 바뀌는 점포는?'];
  $('chips').innerHTML = SUGGEST.map(function (s) { return '<button>' + esc(s) + '</button>'; }).join('');
  $('chips').addEventListener('click', function (e) { if (e.target.tagName === 'BUTTON') { $('chat-text').value = e.target.textContent; send(); } });
  $('chat-send').addEventListener('click', send);
  $('chat-new').addEventListener('click', function () { if (!state.busy) { state.chat = []; renderChat(); } });
  $('chat-text').addEventListener('keydown', function (e) {
    if (e.key === 'Enter' && !e.shiftKey && !e.isComposing) { e.preventDefault(); send(); }
  });
  $('chat-text').addEventListener('input', autoGrow);
  function autoGrow() { var t = $('chat-text'); t.style.height = 'auto'; t.style.height = Math.min(t.scrollHeight, 160) + 'px'; }

  function send() {
    var text = $('chat-text').value.trim();
    if (!text || state.busy) return;
    state.busy = true;
    $('chat-text').value = ''; autoGrow();
    state.chat.push({ role: 'user', parts: [{ text: text }] });   // Gemini 대화 형식
    renderChat('분석 중입니다… (시트 조회·계산에 10~40초 걸릴 수 있어요)');
    google.script.run
      .withSuccessHandler(function (json) {
        var r = JSON.parse(json);
        state.busy = false;
        if (r.error) {
          state.chat.pop();                // 실패한 질문은 대화에서 빼고 입력창으로 돌려놓기
          $('chat-text').value = text; autoGrow();
          renderChat(null, r.error);
          return;
        }
        state.chat = r.messages;
        renderChat(null, null, r.note);
      })
      .withFailureHandler(function (err) {
        state.busy = false;
        state.chat.pop();
        $('chat-text').value = text; autoGrow();
        renderChat(null, '요청 실패: ' + err.message);
      })
      .chatTurn(JSON.stringify(state.chat));
  }

  function renderChat(typing, error, note) {
    var html = [];
    if (!state.chat.length && !typing) {
      html.push('<div class="msg bot"><p><b>안녕하세요! 폐점 분석 챗봇입니다.</b></p><p>점포 현황, 특정 월 폐점 시 예상 비용, 최적 폐점 시점, 브랜드별 과거 폐점 비용 등을 물어보세요. ' +
        '숫자는 시트 데이터를 직접 조회·계산해서 답합니다.</p></div>');
    }
    // 한 질문에 대한 답(도구 조회 여러 번 + 최종 답변)을 말풍선 하나로 묶어서 보여줍니다.
    var tools = [], texts = [];
    var flush = function () {
      var text = texts.join('\\n\\n').trim();
      if (tools.length || text) {
        html.push('<div class="msg bot">' + (tools.length ? '<div class="tools">' + tools.join('') + '</div>' : '') + (text ? md(text) : '') + '</div>');
      }
      tools = []; texts = [];
    };
    state.chat.forEach(function (msg) {
      var parts = msg.parts || [];
      var userText = parts.filter(function (p) { return typeof p.text === 'string'; });
      if (msg.role === 'user' && userText.length) {
        flush();
        html.push('<div class="msg user">' + esc(userText.map(function (p) { return p.text; }).join('\\n')) + '</div>');
        return;
      }
      if (msg.role !== 'model') return;   // 도구 결과(functionResponse)는 화면에 표시하지 않음
      parts.forEach(function (p) {
        if (p.text && !p.thought) texts.push(p.text);
        if (p.functionCall) {
          var inp = p.functionCall.args || {};
          var target = inp.store || inp.brand ? ': ' + (inp.brand ? inp.brand + ' ' : '') + (inp.store || '') : '';
          tools.push('<span class="tool-chip">🔎 ' + esc((TOOL_NAMES[p.functionCall.name] || p.functionCall.name) + target + (inp.close_month ? ' · ' + inp.close_month : '')) + '</span>');
        }
      });
    });
    flush();
    if (typing) html.push('<div class="typing">⏳ ' + esc(typing) + '</div>');
    if (note) html.push('<div class="msg note">' + esc(note) + '</div>');
    if (error) html.push('<div class="msg err">⚠ ' + esc(error) + '</div>');
    $('chat-log').innerHTML = html.join('');
    $('chat-log').scrollTop = $('chat-log').scrollHeight;
    $('chips').hidden = state.chat.length > 0;   // 예시 질문은 대화를 시작하기 전에만 표시
    $('chat-send').disabled = state.busy;
  }

  // 간단한 마크다운 → HTML (굵게, 목록, 표, 제목, 코드)
  function md(src) {
    var inline = function (s) {
      return esc(s).replace(/\\*\\*(.+?)\\*\\*/g, '<b>$1</b>').replace(/\`([^\`]+)\`/g, '<code>$1</code>');
    };
    var lines = src.replace(/\\r/g, '').split('\\n'), out = [], i = 0;
    while (i < lines.length) {
      var l = lines[i];
      if (/^\\s*\\|/.test(l)) {                      // 표
        var rows = [];
        while (i < lines.length && /^\\s*\\|/.test(lines[i])) { rows.push(lines[i]); i++; }
        var cells = function (r) { return r.trim().replace(/^\\||\\|$/g, '').split('|').map(function (c) { return c.trim(); }); };
        var body = rows.filter(function (r) { return !/^\\s*\\|?\\s*:?-{2,}/.test(r); });
        out.push('<table>' + body.map(function (r, k) {
          var tag = k === 0 ? 'th' : 'td';
          return '<tr>' + cells(r).map(function (c) { return '<' + tag + '>' + inline(c) + '</' + tag + '>'; }).join('') + '</tr>';
        }).join('') + '</table>');
        continue;
      }
      if (/^\\s*[-*•]\\s+/.test(l) || /^\\s*\\d+[.)]\\s+/.test(l)) {   // 목록
        var ordered = /^\\s*\\d+[.)]\\s+/.test(l), items = [];
        while (i < lines.length && (/^\\s*[-*•]\\s+/.test(lines[i]) || /^\\s*\\d+[.)]\\s+/.test(lines[i]))) {
          items.push('<li>' + inline(lines[i].replace(/^\\s*([-*•]|\\d+[.)])\\s+/, '')) + '</li>'); i++;
        }
        out.push((ordered ? '<ol>' : '<ul>') + items.join('') + (ordered ? '</ol>' : '</ul>'));
        continue;
      }
      var h = l.match(/^(#{1,4})\\s+(.*)$/);
      if (h) { out.push('<h4>' + inline(h[2]) + '</h4>'); i++; continue; }
      if (/^\\s*(-{3,}|\\*{3,})\\s*$/.test(l)) { i++; continue; }
      if (!l.trim()) { i++; continue; }
      var para = [];
      while (i < lines.length && lines[i].trim() && !/^\\s*\\|/.test(lines[i]) && !/^\\s*[-*•]\\s+/.test(lines[i]) &&
             !/^\\s*\\d+[.)]\\s+/.test(lines[i]) && !/^#{1,4}\\s/.test(lines[i])) { para.push(inline(lines[i])); i++; }
      out.push('<p>' + para.join('<br>') + '</p>');
    }
    return out.join('');
  }

  // ───────── 시작 ─────────
  renderChat();
  if (MODE === 'chat') { $('chat-text').placeholder = '예) 양재본점 언제 폐점하는 게 좋아?'; showTab('chat'); }
  else load();
})();
</script>
</body>
</html>
`;
