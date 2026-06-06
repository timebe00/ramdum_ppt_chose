@echo off
chcp 65001 > nul

echo WinnerGenerator EXE 빌드 시작

python -m pip install pyinstaller python-pptx

python -m PyInstaller ^
  --onefile ^
  --noconsole ^
  --distpath . ^
  --workpath build ^
  --specpath build ^
  --name WinnerGenerator ^
  winner_generator.py

if exist WinnerGenerator.exe (
    echo 생성 완료: WinnerGenerator.exe
) else (
    echo 오류: WinnerGenerator.exe 파일이 생성되지 않았습니다
)

pause