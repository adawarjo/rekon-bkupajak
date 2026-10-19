@echo off
setlocal
cd /d "%~dp0"

python --version >nul 2>&1
if errorlevel 1 goto :nopython

python -c "import openpyxl" >nul 2>&1
if errorlevel 1 (
    echo Memasang library openpyxl...
    python -m pip install --user -r requirements.txt
    if errorlevel 1 goto :gagalpip
)

start "" pythonw "%~dp0app.py"
exit /b 0

:nopython
echo Python belum terpasang di komputer ini.
echo Unduh dari https://www.python.org/downloads/ dan centang "Add python.exe to PATH" saat instalasi.
pause
exit /b 1

:gagalpip
echo Gagal memasang openpyxl. Periksa koneksi internet lalu coba lagi.
pause
exit /b 1
