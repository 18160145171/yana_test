@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"

echo [1/5] 检查 Python...
python --version >nul 2>&1
if errorlevel 1 (
  echo 未检测到 Python，请先安装 Python 3.10+ 并添加到 PATH。
  pause
  exit /b 1
)

echo [2/5] 安装运行依赖...
pip install -r requirements.txt
if errorlevel 1 (
  echo 运行依赖安装失败，请检查网络/代理。
  pause
  exit /b 1
)

echo [3/5] 安装打包依赖...
pip install pyinstaller
if errorlevel 1 (
  echo PyInstaller 安装失败，请检查网络/代理。
  pause
  exit /b 1
)

echo [4/5] 清理旧产物...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist "测试用例生成器.spec" del /q "测试用例生成器.spec"

echo [5/5] 开始打包 EXE...
pyinstaller ^
  --noconfirm ^
  --name "测试用例生成器" ^
  --onedir ^
  --windowed ^
  --add-data "app.py;." ^
  launcher.py

if errorlevel 1 (
  echo 打包失败，请查看上方日志。
  pause
  exit /b 1
)

echo.
echo 打包完成！
echo EXE 路径：%cd%\dist\测试用例生成器\测试用例生成器.exe
pause
endlocal
