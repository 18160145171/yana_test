@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"

echo [1/6] Check Python...
python --version >nul 2>&1
if errorlevel 1 (
  echo Python not found. Please install Python 3.10+ and add to PATH.
  pause
  exit /b 1
)

echo [2/6] Install runtime dependencies...
python -m pip install -r requirements.txt
if errorlevel 1 (
  echo Default index failed, retry with Tsinghua mirror...
  python -m pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
  if errorlevel 1 (
    echo Runtime dependency install failed.
    pause
    exit /b 1
  )
)

echo [3/6] Install pyinstaller...
python -m pip install pyinstaller
if errorlevel 1 (
  echo Default index failed, retry with Tsinghua mirror...
  python -m pip install pyinstaller -i https://pypi.tuna.tsinghua.edu.cn/simple
  if errorlevel 1 (
    echo PyInstaller install failed.
    pause
    exit /b 1
  )
)

echo [4/6] Clean old build artifacts...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist "TestCaseGenerator.spec" del /q "TestCaseGenerator.spec"

echo [5/6] Build EXE...
python -m PyInstaller ^
  --noconfirm ^
  --name "TestCaseGenerator" ^
  --onedir ^
  --windowed ^
  --add-data "app.py;." ^
  --hidden-import streamlit.web.bootstrap ^
  launcher.py

if errorlevel 1 (
  echo Build failed. See logs above.
  pause
  exit /b 1
)

echo [6/6] Done.
echo EXE path: %cd%\dist\TestCaseGenerator\TestCaseGenerator.exe
pause
endlocal
