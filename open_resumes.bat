@echo off
echo ============================================
echo    KRILA RESUME FOLDER
echo ============================================
echo.
echo Opening your generated resumes...
echo.
echo Location: client_resumes\
echo.

start "" "%~dp0client_resumes"

echo.
echo ============================================
echo Folder opened!
echo.
echo Your Krila-formatted DOCX files are there.
echo.
echo To convert to PDF:
echo   1. Double-click DOCX file
echo   2. File -^> Save As -^> PDF
echo   3. Done!
echo ============================================
echo.
pause
