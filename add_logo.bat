@echo off
echo ============================================
echo    ADD KRIFY LOGO TO RESUMES
echo ============================================
echo.
echo INSTRUCTIONS:
echo.
echo 1. Find your Krify logo image file
echo    (the gear logo with green center)
echo.
echo 2. It should be in JPG format
echo.
echo 3. The folder will open now
echo.
echo 4. Copy your logo to that folder
echo.
echo 5. Rename it EXACTLY to: krify_logo.jpg
echo.
echo 6. Restart the app to see logo in PDFs!
echo.
echo ============================================
pause
echo.
echo Opening static folder...
start "" "%~dp0static\"
echo.
echo After adding logo:
echo   - Filename MUST be: krify_logo.jpg
echo   - Restart app: launch_groq.bat
echo.
pause
