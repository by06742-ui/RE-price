RE_price (경량판) — 서울 빌라 호별 예상 거래가 산출기
· 주소(지번 또는 도로명)만으로 조회. PNU 입력 불필요.
· 사용연수·구조 건축물대장 자동 반영, 호별 평당가·예상가, 대량조회 2시트 엑셀.
· 데이터 경량화: 약 40MB → 10MB (road/ho/coords를 npz 이진 압축)

[A] 웹앱(Streamlit): streamlit_app.py + 데이터 → GitHub 연동 후 Streamlit Cloud
[B] 데스크톱 EXE 만들기 — 두 가지 방법
   (1) 클라우드 자동빌드(권장, 내 PC에 설치 불필요):
       이 폴더를 통째로 GitHub 저장소에 올리면 .github/workflows/build.yml 이
       자동으로 Windows EXE를 빌드합니다. 저장소 > Actions 탭 > 최신 실행 >
       'RE_price-exe' 아티팩트 다운로드, 또는 Releases 페이지에서 RE_price.exe 다운로드.
   (2) 내 PC에서 직접: build.bat 더블클릭 → dist\RE_price.exe
       ※ 'pyinstaller 명령을 찾을 수 없음' 오류 방지를 위해 python -m PyInstaller 사용.

[데이터 파일] prices.json, PNU_coords_lite.npz, road_lite.npz, ho_lite.npz, dongnames.csv
