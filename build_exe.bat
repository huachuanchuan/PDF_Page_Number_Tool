@echo off
setlocal EnableDelayedExpansion
cd /d "%~dp0"

echo [1/3] Installing dependencies...
python -m pip install --upgrade pip
if errorlevel 1 goto :error

python -m pip install -r requirements.txt
if errorlevel 1 goto :error

set APP_NAME=PDF_Page_Number_Tool
set EXCLUDES=--exclude-module matplotlib --exclude-module mpl_toolkits --exclude-module pandas --exclude-module scipy --exclude-module openpyxl --exclude-module jupyter --exclude-module IPython --exclude-module notebook --exclude-module seaborn --exclude-module statsmodels --exclude-module numpy --exclude-module PIL --exclude-module psutil --exclude-module requests --exclude-module urllib3 --exclude-module certifi --exclude-module charset_normalizer --exclude-module cryptography

echo [2/3] Building Windows executable with app icon...
python -m PyInstaller --noconfirm --clean --onefile --windowed --icon app_icon.ico !EXCLUDES! --name !APP_NAME! pdf.py
if errorlevel 1 (
    echo Default output name may be occupied. Retrying with fallback name...
    set APP_NAME=PDF_Page_Number_Tool_logo
    python -m PyInstaller --noconfirm --clean --onefile --windowed --icon app_icon.ico !EXCLUDES! --name !APP_NAME! pdf.py
    if errorlevel 1 goto :error
)

echo [3/3] Build complete.
echo EXE file: dist\!APP_NAME!.exe
pause
exit /b 0

:error
echo Build failed. Please review the errors above.
pause
exit /b 1
