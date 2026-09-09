@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"

echo [1/4] Check Python...
python --version >nul 2>&1
if errorlevel 1 (
  echo Python not found. Please install Python 3.10+ and add to PATH.
  pause
  exit /b 1
)

echo [2/4] Check dependencies...
python -c "import streamlit,openai,pandas,openpyxl,docx,pypdf" >nul 2>&1
if errorlevel 1 (
  echo Installing dependencies from default index...
  python -m pip install -r requirements.txt
  if errorlevel 1 (
    echo Default index failed, retry with Tsinghua mirror...
    python -m pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
    if errorlevel 1 (
      echo Dependency install failed. Please check network/proxy and rerun.
      pause
      exit /b 1
    )
  )
)

echo [3/4] Start app...
start "" http://localhost:8501
python -m streamlit run app.py --server.headless true

echo [4/4] App exited.
endlocal
