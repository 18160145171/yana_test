# 测试用例生成器

支持两种输入方式：
- 上传需求文档：`txt` / `docx` / `pdf`
- 直接粘贴需求文本

支持输出：
- 在线预览表格
- 导出 `Excel (.xlsx)`
- 导出 `TSV`（可直接粘贴到 Excel）

## 1. 安装

```bash
cd F:\testcase-generator
pip install -r requirements.txt
```

## 2. API Key 方式

推荐方式：启动后在页面左侧直接填写 `API Key`（仅当前会话，不落盘）。

也支持环境变量方式（可选）：

```powershell
$env:OPENAI_API_KEY="你的API_KEY"
```

### 智谱 API 配置示例

- API 提供方：`智谱`
- Base URL：`https://open.bigmodel.cn/api/paas/v4`
- 模型标识：`glm-4-plus`（或你有权限的其他模型）
- API Key：你的智谱 Key

### Kimi API 配置示例

- API 提供方：`Kimi(月之暗面)`
- Base URL：`https://api.moonshot.cn/v1`
- 模型标识：`moonshot-v1-8k`（或你有权限的 Kimi 模型）
- API Key：你的 Kimi Key

### 5spiritual API 配置示例

- API 提供方：`5spiritual（Responses）`
- Base URL：`https://5spiritual.com`
- 模型标识：`openai/gpt-5.5`
- 接口端点：`Responses（/v1/responses）`
- 推理强度：`中`
- API Key：你的 5spiritual Key

可选：设置模型名（默认 `gpt-5.3-codex-fast`）：

```powershell
$env:MODEL_NAME="gpt-5.3-codex-fast"
```

## 3. 运行

### 方式A（推荐，双击启动）

双击项目目录里的：

- `RUN_APP.bat`

### 方式B（命令行）

```bash
streamlit run app.py
```

## 4. Cloudflare 固定公网访问

本项目是 Streamlit/Python 应用，不能作为纯静态 HTML 直接托管到 Cloudflare Pages。推荐方式是使用 Cloudflare Tunnel：

- 临时访问：`cloudflared tunnel --url http://127.0.0.1:8501`
- 固定访问：Cloudflare 账号 + 已接入 Cloudflare 的域名 + named tunnel

固定访问不要求提交 Git。它会把你指定的域名（例如 `cases.example.com`）转发到本机的 `http://127.0.0.1:8501`。

前置条件：

1. `cloudflared.exe` 已放到 `F:\tools\cloudflared.exe`
2. Cloudflare 账号里已有一个接入 Cloudflare DNS 的域名
3. 本机可以长期运行 Streamlit 应用；如果本机关机或应用没启动，公网地址会打不开

首次配置：

```powershell
cd F:\testcase-generator
.\SETUP_CLOUDFLARE_TUNNEL.bat cases.example.com
```

脚本会要求登录 Cloudflare，并创建固定 Tunnel、DNS 路由和 `%USERPROFILE%\.cloudflared\config.yml`。

启动公网访问：

```powershell
cd F:\testcase-generator
.\START_PUBLIC_TUNNEL.bat
```

固定公网地址：

```text
https://cases.example.com
```

## 5. 打包 EXE（双击可用）

在项目目录双击：

- `BUILD_EXE.bat`

打包完成后可执行文件在：

- `dist\TestCaseGenerator\TestCaseGenerator.exe`

说明：
- 首次打包会自动安装 `pyinstaller`，耗时较长属正常。
- EXE 运行后会自动启动本地网页并打开浏览器。

## 6. 使用说明

1. 上传文档或粘贴需求文本
2. 配置最少用例数量、是否包含 P2
3. 点击“生成测试用例”
4. 下载导出的 Excel 文件

## 7. 输出列格式

- 所属模块
- 用例编号
- 测试用例名称
- 细分项
- 测试输入
- 期望结果
- 优先级（P0/P1/P2）
