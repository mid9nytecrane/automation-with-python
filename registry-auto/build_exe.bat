@echo off
REM ============================================================
REM Build NDC_Entry.exe  (run this on YOUR Windows PC, once)
REM Requires Python 3.10-3.12 (64-bit) with the "py" launcher.
REM Target PCs need nothing installed.
REM ============================================================
setlocal
cd /d "%~dp0"

echo [1/5] Creating clean virtual environment...
if exist .venv rmdir /s /q .venv
py -3 -m venv .venv || goto :fail
call .venv\Scripts\activate.bat

echo [2/5] Installing libraries...
python -m pip install --upgrade pip || goto :fail
pip install pyinstaller playwright pandas openpyxl xlrd python-dotenv || goto :fail

echo [3/5] Downloading Chromium INTO the playwright package folder...
REM PLAYWRIGHT_BROWSERS_PATH=0 = "store browsers inside the playwright package".
REM PyInstaller's --collect-all then copies them into the exe folder.
set PLAYWRIGHT_BROWSERS_PATH=0
python -m playwright install chromium || goto :fail

echo [4/5] Building with PyInstaller (onedir)...
pyinstaller --noconfirm --clean --onedir --console --name NDC_Entry ^
  --collect-all playwright ^
  --hidden-import xlrd --hidden-import openpyxl ^
  main.py || goto :fail

echo [5/5] Adding .env template...
copy /y ".env" "dist\registry-auto\.env" >nul

echo.
echo DONE. Give users the whole folder:  dist\NDC_Entry\
echo They must create dist\NDC_Entry\.env  (copy of .env.example) next to NDC_Entry.exe
pause
exit /b 0

:fail
echo.
echo BUILD FAILED - read the error above.
pause
exit /b 1
