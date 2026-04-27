@echo off
chcp 65001 >nul
echo ============================================
echo   DataMan 의존 패키지 설치
echo ============================================
echo.

echo [1/3] 필수 패키지 설치 중...
pip install olefile PyMuPDF

echo.
echo [2/3] 선택 패키지 설치 중...
pip install python-docx pywin32 pytesseract Pillow

echo.
echo [3/3] Tesseract OCR 설치 확인 중...
where tesseract >nul 2>&1
if %errorlevel%==0 (
    echo   Tesseract OCR이 이미 설치되어 있습니다.
    tesseract --version 2>&1 | findstr /R "^tesseract"
) else (
    if exist "C:\Program Files\Tesseract-OCR\tesseract.exe" (
        echo   Tesseract OCR이 이미 설치되어 있습니다.
    ) else (
        echo.
        echo   ※ OCR 기능을 사용하려면 Tesseract OCR을 별도 설치해야 합니다.
        echo.
        echo   [설치 방법]
        echo   1. 아래 링크에서 Windows 인스톨러를 다운로드하세요:
        echo      https://github.com/UB-Mannheim/tesseract/wiki
        echo   2. 설치 시 "Additional language data" 에서 "Korean" 체크
        echo   3. 설치 경로는 기본값 유지 (C:\Program Files\Tesseract-OCR)
        echo.
        echo   Tesseract 없이도 프로그램은 실행되지만, 스캔 PDF의 OCR은 불가합니다.
        echo.
        set /p "OPEN_LINK=다운로드 페이지를 열까요? (Y/N): "
        if /i "!OPEN_LINK!"=="Y" (
            start https://github.com/UB-Mannheim/tesseract/wiki
        )
    )
)

echo.
echo ============================================
echo   설치 완료!
echo   실행: python dataman.py
echo ============================================
pause
