@echo off
chcp 65001 >nul
setlocal

REM 1) Python 3.11 이상 설치 필요
REM 2) 이 bat 파일을 winner_generator.py와 같은 폴더에서 실행하세요.

python -m venv .venv
call .venv\Scripts\activate
python -m pip install --upgrade pip
pip install -r requirements.txt
pip install pyinstaller

pyinstaller --onefile --windowed --name WinnerGenerator winner_generator.py

echo.
echo 빌드 완료: dist\WinnerGenerator.exe
echo 이 exe를 전체.csv / 블락.csv / 상품.csv / 당첨자.pptx가 있는 폴더에 넣고 실행하거나,
echo 실행 후 작업 폴더를 선택하세요.
pause
