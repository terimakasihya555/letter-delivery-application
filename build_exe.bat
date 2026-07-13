@echo off
cd /d "%~dp0"

echo Installing requirements...
python -m pip install -r requirements.txt

echo.
echo Building desktop app...
python -m PyInstaller --noconsole --onedir --name "AplikasiPengirimanSurat" --add-data "templates;templates" --add-data "assets;assets" run_desktop.py

echo.
echo Build selesai.
echo File aplikasi ada di folder:
echo dist\AplikasiPengirimanSurat
pause