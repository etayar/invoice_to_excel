@echo off
chcp 65001 > nul
echo =====================================================
echo  invoice_to_excel -- Build script
echo =====================================================
echo.

echo [1/2] Installing Python requirements...
pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo ERROR: Failed to install requirements.
    echo Make sure Python 3.9+ and pip are installed and available in PATH.
    pause
    exit /b 1
)

echo.
echo [2/2] Building Windows executable with PyInstaller...
pyinstaller --onefile --windowed --name InvoiceToExcel ^
    --hidden-import=pdfplumber ^
    --hidden-import=pdfminer ^
    --hidden-import=pdfminer.high_level ^
    --hidden-import=pdfminer.layout ^
    --hidden-import=openpyxl ^
    --hidden-import=pandas ^
    --hidden-import=xlrd ^
    app.py

if errorlevel 1 (
    echo.
    echo ERROR: PyInstaller build failed. See messages above.
    pause
    exit /b 1
)

echo.
echo =====================================================
echo  Build complete!
echo  Find InvoiceToExcel.exe inside the  dist\  folder.
echo =====================================================
pause
