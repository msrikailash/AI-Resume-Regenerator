@echo off
echo ============================================
echo    VIEW YOUR LATEST RESUME
echo ============================================
echo.

cd /d "%~dp0client_resumes"

echo Looking for latest resume...
echo.

for /f "delims=" %%i in ('dir *.docx /b /od') do set latest=%%i

if defined latest (
    echo Found: %latest%
    echo.
    echo Opening in Microsoft Word...
    start "" "%latest%"
    echo.
    echo ============================================
    echo Resume opened!
    echo.
    echo This is your Krila-formatted DOCX file.
    echo.
    echo To convert to PDF:
    echo   In Word: File -^> Save As -^> PDF
    echo ============================================
) else (
    echo No resume files found yet.
    echo.
    echo Go to http://localhost:5003 to create one!
)

echo.
pause
