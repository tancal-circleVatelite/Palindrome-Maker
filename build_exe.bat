@echo off
setlocal
cd /d "%~dp0"

set PYTHON_EXE=C:\Users\Ritsu\AppData\Local\Programs\Python\Python310\python.exe

if not exist "%PYTHON_EXE%" (
    echo Python 3.10 interpreter not found: %PYTHON_EXE%
    exit /b 1
)

"%PYTHON_EXE%" -m pip install --upgrade PyInstaller openpyxl
"%PYTHON_EXE%" -m PyInstaller --clean --noconfirm "PalindromeMaker.spec"
if errorlevel 1 (
    echo PyInstaller build failed.
    exit /b 1
)

echo.
echo Build finished.
echo Output: dist\PalindromeMaker.exe
endlocal
