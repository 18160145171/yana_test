@echo off
setlocal
chcp 65001 >nul

cd /d "%~dp0"

echo [1/3] 检查 Python...
python --version >nul 2>&1
if errorlevel 1 (
  echo 未检测到 Python，请先安装 Python 3.10+，并勾选 "Add Python to PATH"。
  pause
  exit /b 1
)

echo [2/3] 检查依赖...
python -c "import streamlit,openai,pandas,openpyxl,docx,pypdf" >nul 2>&1
if errorlevel 1 (
  echo 首次运行，开始安装依赖...
  pip install -r requirements.txt
  if errorlevel 1 (
    echo 依赖安装失败，请检查网络/代理后重试。
    pause
    exit /b 1
  )
)

echo [3/3] 启动程序...
start "" http://localhost:8501
streamlit run app.py --server.headless true

endlocal
