@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo [1/2] 라이브러리 설치 ...
python -m pip install --upgrade pip
python -m pip install numpy openpyxl pyinstaller
echo [2/2] EXE 빌드 ... (pyinstaller 명령 대신 python -m PyInstaller 사용)
python -m PyInstaller --onefile --windowed --name "RE_price" --collect-all openpyxl --exclude-module scipy --exclude-module pandas --exclude-module matplotlib --add-data "prices.json;." --add-data "PNU_coords_lite.npz;." --add-data "dongnames.csv;." --add-data "ho_lite.npz;." --add-data "road_lite.npz;." RE_price.py
echo.
echo 완료! dist\RE_price.exe
pause
