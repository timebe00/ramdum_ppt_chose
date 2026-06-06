@echo off
chcp 65001 >nul
setlocal

REM 이 bat 파일을 winner_generator.py와 같은 폴더에 두고 실행하세요.
REM 빌드 결과 exe는 dist 폴더가 아니라 현재 폴더에 WinnerGenerator.exe 로 생성됩니다.

cd /d "%~dp0"

echo =================================
echo WinnerGenerator EXE Build Start
echo =================================

IF NOT EXIST "winner_generator.py" (
    echo 오류: 현재 폴더에 winner_generator.py 파일이 없습니다.
    pause
    exit /b 1
)

REM 기존 빌드 결과 정리
IF EXIST "build" rmdir /s /q "build"
IF EXIST "dist" rmdir /s /q "dist"
IF EXIST "WinnerGenerator.spec" del /q "WinnerGenerator.spec"
IF EXIST "WinnerGenerator.exe" del /q "WinnerGenerator.exe"

REM 가상환경 생성/활성화
IF NOT EXIST ".venv" (
    python -m venv .venv
)
call .venv\Scripts\activate

python -m pip install --upgrade pip

IF EXIST "requirements.txt" (
    pip install -r requirements.txt
) ELSE (
    pip install python-pptx pyinstaller
)

pip install pyinstaller

REM exe 빌드: 우선 dist에 생성됨
pyinstaller --onefile --windowed --name WinnerGenerator winner_generator.py

IF NOT EXIST "dist\WinnerGenerator.exe" (
    echo 오류: dist\WinnerGenerator.exe 파일이 생성되지 않았습니다.
    pause
    exit /b 1
)

REM 현재 폴더로 복사
copy /Y "dist\WinnerGenerator.exe" ".\WinnerGenerator.exe" >nul

IF EXIST "WinnerGenerator.exe" (
    echo.
    echo =================================
    echo 빌드 완료
    echo =================================
    echo 현재 폴더에 생성됨: %cd%\WinnerGenerator.exe
    echo.
) ELSE (
    echo 오류: WinnerGenerator.exe를 현재 폴더로 복사하지 못했습니다.
    pause
    exit /b 1
)

pause
