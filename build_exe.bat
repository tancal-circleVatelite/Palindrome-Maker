@echo off
setlocal

set PYTHON_DIR=C:\Users\Ritsu\AppData\Local\Programs\Python\Python310
set PYTHON=%PYTHON_DIR%\python.exe
set REPO=%~dp0
set VENV=%REPO%.build-venv
set BUILD=%REPO%build
set DIST=%REPO%dist

if exist "%BUILD%" rmdir /s /q "%BUILD%"
if exist "%DIST%" rmdir /s /q "%DIST%"
if exist "%VENV%" rmdir /s /q "%VENV%"

"%PYTHON%" -m venv "%VENV%"
call "%VENV%\Scripts\activate.bat"

python -m pip install --upgrade pip
python -m pip install pyinstaller openpyxl

python -m PyInstaller ^
  --clean ^
  --noconfirm ^
  --onefile ^
  --windowed ^
  --name PalindromeMaker ^
  --hidden-import tkinter ^
  --hidden-import _tkinter ^
  --collect-all tkinter ^
  --add-data "%REPO%juman_bunsetsu.csv;." ^
  --add-data "%PYTHON_DIR%\tcl;tcl" ^
  --add-binary "%PYTHON_DIR%\DLLs\_tkinter.pyd;." ^
  --add-binary "%PYTHON_DIR%\DLLs\tcl86t.dll;." ^
  --add-binary "%PYTHON_DIR%\DLLs\tk86t.dll;." ^
  "%REPO%palindrome_gui.py"

if exist "%REPO%dist\PalindromeMaker.exe" (
  echo.
  echo Build succeeded.
  echo Output: %REPO%dist\PalindromeMaker.exe
) else (
  echo.
  echo Build failed. Check the output above.
  exit /b 1
)

exit /b 0
