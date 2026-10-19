@echo off
rem Membuat ulang RekonsiliasiPajakSIPD.exe (single file) setelah kode app.py / konsolidasi.py diubah.
rem Modul yang tidak dipakai aplikasi dikecualikan supaya ukuran exe kecil (numpy/pandas ikut terbawa lewat openpyxl).
rem hashlib jangan dikecualikan: dibutuhkan openpyxl.
setlocal
cd /d "%~dp0"
set EXCLUDE=--exclude-module numpy --exclude-module pandas --exclude-module PIL --exclude-module lxml --exclude-module defusedxml
set EXCLUDE=%EXCLUDE% --exclude-module unittest --exclude-module pydoc --exclude-module doctest --exclude-module pdb
set EXCLUDE=%EXCLUDE% --exclude-module sqlite3 --exclude-module multiprocessing --exclude-module asyncio --exclude-module xmlrpc
set EXCLUDE=%EXCLUDE% --exclude-module http.server --exclude-module ssl --exclude-module _ssl --exclude-module _hashlib
set EXCLUDE=%EXCLUDE% --exclude-module lzma --exclude-module _lzma --exclude-module bz2 --exclude-module _bz2
python -m pip install --user -r requirements.txt pyinstaller
python -m PyInstaller --noconfirm --clean --onefile --windowed --name RekonsiliasiPajakSIPD --paths . --distpath . --workpath "%TEMP%\pyi_konsolidasi" --specpath "%TEMP%\pyi_konsolidasi" %EXCLUDE% app.py
pause
