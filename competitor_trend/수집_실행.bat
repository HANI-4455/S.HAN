@echo off
chcp 65001 >nul
REM 경쟁사 동향 데이터를 이 PC(국내 IP)에서 수집해 GitHub에 올린다.
REM GitHub Actions 러너(해외 IP)는 투썸 사이트 접속이 막혀 투썸만 갱신되지 않는다.
REM 이 파일은 S.HAN 저장소를 clone 한 폴더 안의 competitor_trend 에서 실행해야 한다.

cd /d "%~dp0"

git rev-parse --is-inside-work-tree >nul 2>&1
if errorlevel 1 (
  echo [오류] 이 폴더는 git 저장소가 아닙니다.
  echo        git clone https://github.com/HANI-4455/S.HAN.git 로 내려받은 폴더에서 실행하세요.
  pause
  exit /b 1
)

echo [1/4] 최신 상태로 맞추는 중...
git pull --rebase || goto :fail

echo [2/4] 필요한 라이브러리 확인...
python -m pip install -q -r requirements.txt || goto :fail

echo [3/4] 경쟁사 제품 + 뉴스 수집...
python collect.py || goto :fail

echo [4/4] GitHub에 반영...
git add data/competitor_data.json
git diff --cached --quiet && (echo 변경 없음 - 올릴 내용이 없습니다.) || (
  git commit -m "경쟁사 동향 데이터 수동 갱신" && git push
)

echo.
echo 완료: https://hani-4455.github.io/S.HAN/competitor_trend.html
pause
exit /b 0

:fail
echo.
echo [실패] 위 메시지를 확인하세요.
pause
exit /b 1
