import ast
import io
import json
import os
import re
from typing import Any, Dict, List, Optional, Tuple
from pathlib import Path

import httpx
import pandas as pd
import streamlit as st
from docx import Document
from openai import (
    OpenAI,
    APIConnectionError,
    AuthenticationError,
    PermissionDeniedError,
    NotFoundError,
    RateLimitError,
    BadRequestError,
)
from pypdf import PdfReader


DEFAULT_MODEL = os.getenv("MODEL_NAME", "gpt-5.3-codex-fast")
DEFAULT_COLUMNS = [
    "所属模块",
    "用例编号",
    "测试用例名称",
    "细分项",
    "测试输入",
    "期望结果",
    "优先级",
]

# 各模型常用字段别名 -> 统一映射到 DEFAULT_COLUMNS
COLUMN_ALIASES: Dict[str, Tuple[str, ...]] = {
    "所属模块": ("所属模块", "module", "模块", "功能模块", "业务模块", "mod", "功能"),
    "用例编号": ("用例编号", "编号", "case_id", "caseid", "id", "用例ID", "用例id"),
    "测试用例名称": (
        "测试用例名称",
        "用例名称",
        "case_name",
        "casename",
        "标题",
        "名称",
        "title",
        "name",
        "用例名",
        "case",
        "测试点",
    ),
    "细分项": ("细分项", "子项", "子场景", "步骤", "detail", "subitem", "场景"),
    "测试输入": ("测试输入", "输入", "前置", "预置条件", "steps", "input", "操作步骤", "条件"),
    "期望结果": ("期望结果", "预期", "预期结果", "expected", "output", "验收标准", "结果"),
    "优先级": ("优先级", "priority", "等级", "级别", "重要级"),
}

PROVIDER_PRESETS = {
    "GPT(OpenAI)": {"base_url": "https://api.openai.com/v1", "model": "gpt-4.1"},
    "OpenAI兼容": {"base_url": "", "model": DEFAULT_MODEL},
    # 智谱 OpenAI 兼容根路径，末尾 / 与官方文档一致，避免与部分旧版客户端拼接路径异常
    "智谱": {"base_url": "https://open.bigmodel.cn/api/paas/v4/", "model": "glm-4-plus"},
    "Kimi(月之暗面)": {"base_url": "https://api.moonshot.cn/v1", "model": "moonshot-v1-8k"},
}
ENDPOINT_PRESETS = {
    "自动识别": "auto",
    "Responses（/v1/responses）": "responses",
    "Chat Completions（/v1/chat/completions）": "chat_completions",
}
CONFIG_PATH = Path.home() / ".testcase_generator_config.json"


def load_local_config() -> Dict:
    if not CONFIG_PATH.exists():
        return {}
    try:
        return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_local_config(data: Dict) -> None:
    try:
        CONFIG_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        # 本地保存失败不影响主流程
        pass


def read_txt(file_bytes: bytes) -> str:
    for enc in ("utf-8", "gbk", "gb2312", "utf-16"):
        try:
            return file_bytes.decode(enc)
        except UnicodeDecodeError:
            continue
    return file_bytes.decode("utf-8", errors="ignore")


def read_docx(file_bytes: bytes) -> str:
    bio = io.BytesIO(file_bytes)
    doc = Document(bio)
    lines = [p.text for p in doc.paragraphs if p.text and p.text.strip()]
    return "\n".join(lines)


def read_pdf(file_bytes: bytes) -> str:
    bio = io.BytesIO(file_bytes)
    reader = PdfReader(bio)
    texts: List[str] = []
    for page in reader.pages:
        txt = page.extract_text() or ""
        if txt.strip():
            texts.append(txt.strip())
    return "\n".join(texts)


def parse_uploaded_file(uploaded_file) -> str:
    if uploaded_file is None:
        return ""
    file_name = uploaded_file.name.lower()
    raw = uploaded_file.read()
    if file_name.endswith(".txt"):
        return read_txt(raw)
    if file_name.endswith(".docx"):
        return read_docx(raw)
    if file_name.endswith(".pdf"):
        return read_pdf(raw)
    raise ValueError("仅支持 txt/docx/pdf 文件")


def build_prompt(source_text: str, case_count_hint: int, include_p2: bool) -> str:
    p2_rule = "必须包含 P0/P1/P2 全量覆盖。" if include_p2 else "优先输出 P0/P1，P2可选。"
    return f"""
你是资深测试架构师。根据输入需求文本生成结构化测试用例。

要求：
1) 输出必须是 JSON 数组，每一项是一个对象。
2) 字段固定为：{DEFAULT_COLUMNS}
3) 优先级只能是 P0/P1/P2。
4) 覆盖正常流程、异常流程、边界条件、状态切换、失败重试、兼容性。
5) 用例描述简洁易懂，可直接用于测试执行。
6) 用例编号从 TC-001 递增，不重复。
7) {p2_rule}
8) 目标用例数量不少于 {case_count_hint} 条。
9) 严禁输出除 JSON 之外的其他文本。

需求文本如下：
---
{source_text}
---
""".strip()


def build_supplement_prompt(source_text: str, existing_rows: List[Dict], need_count: int, include_p2: bool) -> str:
    p2_rule = "必须覆盖 P2 场景。" if include_p2 else "P2 可选。"
    existing_preview = json.dumps(existing_rows[:200], ensure_ascii=False)
    return f"""
你是资深测试架构师。当前已有一批测试用例，但条数不足且覆盖不全。
请“仅补充新增用例”，不要重复已有用例。

要求：
1) 仅输出 JSON 数组，不要任何解释文字。
2) 每条必须包含字段：{DEFAULT_COLUMNS}
3) 优先补充：异常流程、边界条件、失败重试、状态切换、兼容性、权限、幂等与稳定性。
4) 用例编号继续递增（可先占位，系统会二次规范化）。
5) 至少补充 {need_count} 条新用例。
6) {p2_rule}

原始需求文本：
---
{source_text}
---

已有用例（禁止重复）：
---
{existing_preview}
---
""".strip()


def dedupe_cases(rows: List[Dict]) -> List[Dict]:
    seen = set()
    result = []
    for row in rows:
        key = (
            row.get("所属模块", "").strip().lower(),
            row.get("测试用例名称", "").strip().lower(),
            row.get("细分项", "").strip().lower(),
            row.get("测试输入", "").strip().lower(),
        )
        if key in seen:
            continue
        seen.add(key)
        result.append(row)
    return result


def normalize_rows(rows: List[Dict]) -> List[Dict]:
    normalized = []
    for i, item in enumerate(rows, start=1):
        row = {col: str(item.get(col, "")).strip() for col in DEFAULT_COLUMNS}
        if row["优先级"] not in {"P0", "P1", "P2"}:
            row["优先级"] = "P1"
        row["用例编号"] = f"TC-{i:03d}"
        normalized.append(row)
    return normalized


def _map_row_aliases(item: Any) -> Dict[str, str]:
    """将单条用例 dict 的各模型字段名统一为 DEFAULT_COLUMNS。"""
    if not isinstance(item, dict):
        return {c: "" for c in DEFAULT_COLUMNS}
    lower_to_orig = {str(k).strip().lower(): k for k in item.keys()}
    out: Dict[str, str] = {}
    for col in DEFAULT_COLUMNS:
        val: Any = None
        if col in item:
            val = item[col]
        else:
            for alt in COLUMN_ALIASES.get(col, ()):
                lk = str(alt).strip().lower()
                orig = lower_to_orig.get(lk)
                if orig is not None:
                    val = item[orig]
                    break
        out[col] = "" if val is None else str(val).strip()
    return out


def _row_match_score(item: Any) -> int:
    """非空标准字段数量，用于判断一段 JSON 对象是否像一条用例。"""
    if not isinstance(item, dict):
        return 0
    m = _map_row_aliases(item)
    return sum(1 for c in DEFAULT_COLUMNS if m.get(c, "").strip())


def _finalize_case_rows(raw_list: Any) -> List[Dict]:
    """过滤、展开嵌套字符串、别名映射后的用例行列表。"""
    if not isinstance(raw_list, list):
        return []
    out: List[Dict] = []
    for item in raw_list:
        if isinstance(item, str):
            s = item.strip()
            if not s:
                continue
            try:
                item = json.loads(s)
            except Exception:
                try:
                    item = ast.literal_eval(s)
                except Exception:
                    continue
        if not isinstance(item, dict):
            continue
        mapped = _map_row_aliases(item)
        if not any(mapped.get(c, "").strip() for c in DEFAULT_COLUMNS):
            continue
        out.append(mapped)
    return out


def _json_loads_lenient(fragment: str) -> Any:
    """兼容尾随逗号、弯引号等常见非标准 JSON。"""
    s = (fragment or "").strip().lstrip("\ufeff")
    if not s:
        raise json.JSONDecodeError("empty", "", 0)
    variants: List[str] = []
    cur = s
    for _ in range(3):
        variants.append(cur)
        nxt = re.sub(r",\s*([}\]])", r"\1", cur)
        if nxt == cur:
            break
        cur = nxt
    smart = (
        s.replace("\u201c", '"')
        .replace("\u201d", '"')
        .replace("\u2018", "'")
        .replace("\u2019", "'")
    )
    variants.append(smart)
    variants.append(re.sub(r",\s*([}\]])", r"\1", smart))
    seen: set = set()
    for v in variants:
        if not v or v in seen:
            continue
        seen.add(v)
        try:
            return json.loads(v)
        except json.JSONDecodeError:
            continue
    raise json.JSONDecodeError("lenient parse failed", s, 0)


def _try_ast_literal(fragment: str) -> Optional[Any]:
    t = (fragment or "").strip().lstrip("\ufeff")
    if not t or t[0] not in "[{":
        return None
    try:
        return ast.literal_eval(t)
    except (ValueError, SyntaxError, MemoryError):
        return None


def _extract_balanced_region(text: str, start: int, open_ch: str, close_ch: str) -> str:
    if start < 0 or start >= len(text) or text[start] != open_ch:
        return ""
    depth = 0
    in_string = False
    escape = False
    for i in range(start, len(text)):
        ch = text[i]
        if in_string:
            if escape:
                escape = False
            elif ch == "\\":
                escape = True
            elif ch == '"':
                in_string = False
            continue
        else:
            if ch == '"':
                in_string = True
                continue
            if ch == open_ch:
                depth += 1
            elif ch == close_ch:
                depth -= 1
                if depth == 0:
                    return text[start : i + 1]
    return ""


def _each_balanced_substring(text: str, open_ch: str, close_ch: str, min_len: int = 12) -> List[str]:
    seen: set = set()
    out: List[str] = []
    for i, ch in enumerate(text):
        if ch != open_ch:
            continue
        frag = _extract_balanced_region(text, i, open_ch, close_ch)
        if frag and len(frag) >= min_len and frag not in seen:
            seen.add(frag)
            out.append(frag)
    return out


def _gather_markdown_fences(text: str) -> List[str]:
    blocks: List[str] = []
    for m in re.finditer(r"```(?:json|JSON)?\s*([\s\S]*?)\s*```", text):
        inner = m.group(1).strip()
        if inner:
            blocks.append(inner)
    return blocks


def _candidate_text_snippets(text: str) -> List[str]:
    """从模型原文中拆出若干候选片段（全文、各代码块、去掉前缀说明后）。"""
    raw = (text or "").strip().lstrip("\ufeff")
    if not raw:
        return []
    ordered: List[str] = []
    seen: set = set()

    def add(s: str) -> None:
        s = (s or "").strip()
        if s and s not in seen:
            seen.add(s)
            ordered.append(s)

    add(raw)
    for b in _gather_markdown_fences(raw):
        add(b)
    for marker in ("[", "{"):
        idx = raw.find(marker)
        if idx > 0:
            add(raw[idx:])
    return ordered


def _parse_value_to_case_list(parsed: Any) -> Optional[List[Any]]:
    if isinstance(parsed, list):
        return parsed
    if isinstance(parsed, dict):
        coerced = _coerce_to_case_list(parsed)
        if isinstance(coerced, list):
            return coerced
    return None


def _parse_fragment_to_value(frag: str) -> Optional[Any]:
    try:
        return _json_loads_lenient(frag)
    except Exception:
        return _try_ast_literal(frag)


def _try_parse_array_fragment(frag: str) -> Optional[List[Dict]]:
    parsed = _parse_fragment_to_value(frag)
    if parsed is None:
        return None
    lst = _parse_value_to_case_list(parsed)
    if isinstance(lst, list) and lst:
        finalized = _finalize_case_rows(lst)
        if finalized:
            return finalized
    return None


def _try_parse_object_fragment(frag: str) -> Optional[List[Dict]]:
    parsed = _parse_fragment_to_value(frag)
    if parsed is None:
        return None
    lst = _parse_value_to_case_list(parsed)
    if isinstance(lst, list) and lst:
        finalized = _finalize_case_rows(lst)
        if finalized:
            return finalized
    if isinstance(parsed, dict) and _row_match_score(parsed) >= 1:
        finalized = _finalize_case_rows([parsed])
        if finalized:
            return finalized
    return None


def _parse_ndjson_lines(text: str) -> Optional[List[Dict]]:
    rows: List[Dict] = []
    for line in text.splitlines():
        line = line.strip()
        if not line or line.startswith("//"):
            continue
        if not line.startswith("{"):
            continue
        frag = _extract_balanced_region(line, 0, "{", "}")
        if not frag:
            continue
        try:
            obj = _json_loads_lenient(frag)
        except Exception:
            obj = _try_ast_literal(frag)
        if isinstance(obj, dict) and _row_match_score(obj) >= 1:
            rows.append(_map_row_aliases(obj))
    return rows if len(rows) >= 1 else None


def _parse_sequence_of_objects(text: str) -> Optional[List[Dict]]:
    """模型连续输出多个 {...} 且无外层数组时的兜底。"""
    rows: List[Dict] = []
    for i, ch in enumerate(text):
        if ch != "{":
            continue
        frag = _extract_balanced_region(text, i, "{", "}")
        if not frag or len(frag) < 10:
            continue
        try:
            obj = _json_loads_lenient(frag)
        except Exception:
            obj = _try_ast_literal(frag)
        if isinstance(obj, dict) and _row_match_score(obj) >= 1:
            rows.append(_map_row_aliases(obj))
    if not rows:
        return None
    dedup: List[Dict] = []
    seen: set = set()
    for r in rows:
        key = tuple(r.get(c, "") for c in DEFAULT_COLUMNS[:4])
        if key in seen:
            continue
        seen.add(key)
        dedup.append(r)
    return dedup if dedup else None


def parse_cases_from_content(content: str) -> List[Dict]:
    """
    从模型返回文本中解析用例列表，尽量兼容各厂商模型的非严格 JSON：
    markdown 代码块、尾随逗号、弯引号、Python 字面量、包裹对象、连续 {...}、NDJSON 等。
    """
    snippets = _candidate_text_snippets(content)
    if not snippets:
        return []

    for snip in snippets:
        # 1) 首个平衡 JSON 数组 + 全文内所有候选数组片段
        first_arr = _extract_first_json_array(snip)
        for frag in [x for x in [first_arr] if x] + _each_balanced_substring(snip, "[", "]"):
            got = _try_parse_array_fragment(frag)
            if got:
                return got

        # 2) 多个独立 {...}（换行或紧凑拼接），须在“只取首个对象”之前，否则会丢后续用例
        seq_got = _parse_sequence_of_objects(snip)
        if seq_got and len(seq_got) >= 2:
            return seq_got

        # 3) 每行一个 JSON 对象
        got = _parse_ndjson_lines(snip)
        if got:
            return got

        # 4) 单个 JSON 对象（含 {"cases":[...]} 等）及文中其它 {...} 片段
        if seq_got and len(seq_got) == 1:
            return seq_got
        first_obj = _extract_first_json_object(snip)
        for frag in [x for x in [first_obj] if x] + _each_balanced_substring(snip, "{", "}", min_len=20):
            got = _try_parse_object_fragment(frag)
            if got:
                return got

        if seq_got:
            return seq_got

    return []


def _extract_first_json_array(text: str) -> str:
    """
    从模型返回文本中提取首个 JSON 数组字符串，兼容：
    - ```json ... ```
    - 前后夹杂解释文字
    - 直接返回数组
    """
    s = (text or "").strip()
    if not s:
        return ""

    # 去掉 markdown 代码块包裹
    fence_match = re.search(r"```(?:json)?\s*([\s\S]*?)\s*```", s, re.IGNORECASE)
    if fence_match:
        s = fence_match.group(1).strip()

    # 直接就是 JSON 数组
    if s.startswith("[") and s.endswith("]"):
        return s

    # 提取第一个平衡的 JSON 数组
    start = s.find("[")
    if start == -1:
        return ""
    depth = 0
    in_string = False
    escape = False
    for i in range(start, len(s)):
        ch = s[i]
        if in_string:
            if escape:
                escape = False
            elif ch == "\\":
                escape = True
            elif ch == '"':
                in_string = False
            continue
        else:
            if ch == '"':
                in_string = True
                continue
            if ch == "[":
                depth += 1
            elif ch == "]":
                depth -= 1
                if depth == 0:
                    return s[start : i + 1]
    return ""


def _extract_first_json_object(text: str) -> str:
    s = (text or "").strip()
    if not s:
        return ""
    fence_match = re.search(r"```(?:json)?\s*([\s\S]*?)\s*```", s, re.IGNORECASE)
    if fence_match:
        s = fence_match.group(1).strip()
    if s.startswith("{") and s.endswith("}"):
        return s
    start = s.find("{")
    if start == -1:
        return ""
    depth = 0
    in_string = False
    escape = False
    for i in range(start, len(s)):
        ch = s[i]
        if in_string:
            if escape:
                escape = False
            elif ch == "\\":
                escape = True
            elif ch == '"':
                in_string = False
            continue
        else:
            if ch == '"':
                in_string = True
                continue
            if ch == "{":
                depth += 1
            elif ch == "}":
                depth -= 1
                if depth == 0:
                    return s[start : i + 1]
    return ""


def _coerce_to_case_list(parsed_json: Any) -> Optional[List[Any]]:
    if isinstance(parsed_json, list):
        return parsed_json
    if isinstance(parsed_json, dict):
        key_candidates = (
            "cases",
            "test_cases",
            "testcases",
            "data",
            "items",
            "rows",
            "result",
            "测试用例",
            "用例列表",
            "用例",
            "list",
            "output",
            "content",
            "case_list",
            "testCaseList",
            "TestCases",
            "casesList",
            "测试用例列表",
            "生成结果",
            "testcases_list",
        )
        for key in key_candidates:
            v = parsed_json.get(key)
            if isinstance(v, list):
                return v
        lower_orig = {str(k).strip().lower(): k for k in parsed_json.keys()}
        for key in key_candidates:
            lk = key.lower()
            orig = lower_orig.get(lk)
            if orig is None:
                continue
            v = parsed_json.get(orig)
            if isinstance(v, list):
                return v
    return None


def _repair_with_model(
    api_key: str,
    base_url: str,
    model_name: str,
    raw_text: str,
    endpoint_mode: str = "auto",
    proxy_url: str = "",
    skip_ssl_verify: bool = False,
) -> str:
    max_chars = 14000
    if len(raw_text) > max_chars:
        raw_text = (
            raw_text[: max_chars // 2]
            + "\n\n...(中间已省略，请仅根据可见片段修复)...\n\n"
            + raw_text[-max_chars // 2 :]
        )
    cols = json.dumps(DEFAULT_COLUMNS, ensure_ascii=False)
    repair_prompt = f"""
请将下面内容中的测试用例整理为可机器解析的结构（二选一，不要 markdown、不要解释）：
1) 纯 JSON 数组：[{{...}},...]
2) 或单个 JSON 对象包裹：{{"cases":[{{...}},...]}}
每个用例对象须包含字段（键名必须一致）：{cols}

原始内容：
---
{raw_text}
---
""".strip()
    return _request_model_with_fallback(
        api_key=api_key,
        base_url=base_url,
        model_name=model_name,
        system_prompt="你是JSON修复器。只输出一段合法JSON：要么是数组[...]，要么是对象{\"cases\":[...]}。不要代码围栏。",
        user_prompt=repair_prompt,
        endpoint_mode=endpoint_mode,
        proxy_url=proxy_url,
        skip_ssl_verify=skip_ssl_verify,
        temperature=0,
        max_tokens=8192,
    )


def _normalize_openai_base_url(base_url: str) -> str:
    """与 OpenAI SDK / 智谱文档对齐：非空时去掉首尾空白并保证以 / 结尾。"""
    u = (base_url or "").strip()
    if not u:
        return ""
    return u if u.endswith("/") else u + "/"


def _build_client(
    api_key: str,
    base_url: str,
    proxy_url: str = "",
    skip_ssl_verify: bool = False,
) -> OpenAI:
    # 显式配置了代理时关闭 trust_env，避免系统 HTTP(S)_PROXY 与手工代理叠加导致连到错误的中继
    use_proxy = bool((proxy_url or "").strip())
    http_client = httpx.Client(
        timeout=60,
        verify=(False if skip_ssl_verify else True),
        proxy=(proxy_url.strip() or None),
        trust_env=(False if use_proxy else True),
    )
    kwargs = {"api_key": api_key, "timeout": 60, "max_retries": 1, "http_client": http_client}
    norm_base = _normalize_openai_base_url(base_url)
    if norm_base:
        kwargs["base_url"] = norm_base
    return OpenAI(**kwargs)


class ModelAPIError(RuntimeError):
    def __init__(self, status_code: int, response_text: str, url: str):
        self.status_code = status_code
        self.response_text = response_text
        self.url = url
        super().__init__(f"HTTP {status_code}: {response_text[:1000]}")


def _resolve_endpoint_mode(endpoint_mode: str, base_url: str, model_name: str) -> str:
    mode = (endpoint_mode or "auto").strip().lower()
    if mode in {"responses", "chat_completions"}:
        return mode

    # GPT-5 系列和未带 /v1 的网关地址优先使用 Responses API。
    base = (base_url or "").strip().rstrip("/").lower()
    model = (model_name or "").strip().lower()
    if model.startswith(("gpt-5", "o1", "o3", "o4")) or not base.endswith("/v1"):
        return "responses"
    return "chat_completions"


def _endpoint_url(base_url: str, endpoint_mode: str) -> str:
    base = (base_url or "").strip().rstrip("/")
    if not base:
        base = "https://api.openai.com/v1"
    lower = base.lower()
    suffix = "/responses" if endpoint_mode == "responses" else "/chat/completions"
    if lower.endswith(suffix):
        return base
    if lower.endswith("/v1"):
        return base + suffix
    return base + "/v1" + suffix


def _build_http_client(
    proxy_url: str = "",
    skip_ssl_verify: bool = False,
) -> httpx.Client:
    use_proxy = bool((proxy_url or "").strip())
    return httpx.Client(
        timeout=90,
        verify=(False if skip_ssl_verify else True),
        proxy=(proxy_url.strip() or None),
        trust_env=(False if use_proxy else True),
    )


def _extract_response_text(payload: Any) -> str:
    """兼容 Responses、Chat Completions 以及部分第三方网关的返回格式。"""
    if isinstance(payload, str):
        return payload.strip()
    if not isinstance(payload, dict):
        return ""

    output_text = payload.get("output_text")
    if isinstance(output_text, str) and output_text.strip():
        return output_text.strip()

    choices = payload.get("choices")
    if isinstance(choices, list) and choices:
        first = choices[0]
        if isinstance(first, dict):
            message = first.get("message")
            if isinstance(message, dict):
                content = message.get("content")
                if isinstance(content, str) and content.strip():
                    return content.strip()
                if isinstance(content, list):
                    parts = [
                        p.get("text", "")
                        for p in content
                        if isinstance(p, dict) and isinstance(p.get("text"), str)
                    ]
                    if parts:
                        return "".join(parts).strip()
            text = first.get("text")
            if isinstance(text, str) and text.strip():
                return text.strip()

    output = payload.get("output")
    if isinstance(output, list):
        parts: List[str] = []
        for item in output:
            if not isinstance(item, dict):
                continue
            content = item.get("content")
            if isinstance(content, str):
                parts.append(content)
            elif isinstance(content, list):
                for part in content:
                    if isinstance(part, dict) and isinstance(part.get("text"), str):
                        parts.append(part["text"])
        if parts:
            return "".join(parts).strip()

    for key in ("content", "text", "result", "data"):
        value = payload.get(key)
        if isinstance(value, str) and value.strip():
            return value.strip()
    return ""


def _request_model(
    api_key: str,
    base_url: str,
    model_name: str,
    system_prompt: str,
    user_prompt: str,
    endpoint_mode: str = "auto",
    proxy_url: str = "",
    skip_ssl_verify: bool = False,
    temperature: Optional[float] = None,
    max_tokens: Optional[int] = None,
) -> str:
    resolved_mode = _resolve_endpoint_mode(endpoint_mode, base_url, model_name)
    url = _endpoint_url(base_url, resolved_mode)
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }
    if resolved_mode == "responses":
        payload: Dict[str, Any] = {
            "model": model_name,
            "input": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
        }
        if max_tokens is not None:
            payload["max_output_tokens"] = max_tokens
    else:
        payload = {
            "model": model_name,
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
        }
        if temperature is not None:
            payload["temperature"] = temperature
        if max_tokens is not None:
            payload["max_tokens"] = max_tokens

    try:
        with _build_http_client(proxy_url, skip_ssl_verify) as http_client:
            response = http_client.post(url, headers=headers, json=payload)
    except Exception:
        raise

    response_text = response.text or ""
    if response.status_code >= 400:
        raise ModelAPIError(response.status_code, response_text, url)
    try:
        response_payload = response.json()
    except ValueError:
        response_payload = response_text

    content = _extract_response_text(response_payload)
    if content:
        return content

    raise RuntimeError(
        f"接口返回成功但无法解析文本（端点：{url}，返回片段：{response_text[:500]}）"
    )


def _request_model_with_fallback(
    api_key: str,
    base_url: str,
    model_name: str,
    system_prompt: str,
    user_prompt: str,
    endpoint_mode: str = "auto",
    proxy_url: str = "",
    skip_ssl_verify: bool = False,
    temperature: Optional[float] = None,
    max_tokens: Optional[int] = None,
) -> str:
    try:
        return _request_model(
            api_key,
            base_url,
            model_name,
            system_prompt,
            user_prompt,
            endpoint_mode,
            proxy_url,
            skip_ssl_verify,
            temperature,
            max_tokens,
        )
    except ModelAPIError as error:
        # 自动模式下，部分旧网关不提供 Responses API，遇到 404/405 时回退旧接口。
        if (endpoint_mode or "auto").strip().lower() == "auto" and error.status_code in {404, 405}:
            return _request_model(
                api_key,
                base_url,
                model_name,
                system_prompt,
                user_prompt,
                "chat_completions",
                proxy_url,
                skip_ssl_verify,
                temperature,
                max_tokens,
            )
        raise


def _diagnose_exception(e: Exception, base_url: str, model_name: str) -> str:
    if isinstance(e, ModelAPIError):
        if e.status_code == 401:
            return "鉴权失败(401)：API Key 无效、过期或复制不完整。"
        if e.status_code == 403:
            return "权限不足(403)：当前 Key 无该模型权限或项目权限不足。"
        if e.status_code == 404:
            return f"资源不存在(404)：请检查接口端点、Base URL 或模型名（当前：{model_name}）。"
        if e.status_code == 429:
            return "限流/配额不足(429)：请检查额度、并发限制或稍后重试。"
        return (
            f"模型接口返回 HTTP {e.status_code}。\n"
            f"- 当前 Base URL: {base_url or '默认'}\n"
            f"- 当前模型: {model_name}\n"
            f"- 返回信息: {e.response_text[:1000]}"
        )
    if isinstance(e, APIConnectionError):
        cause = str(getattr(e, "__cause__", "")) or str(e)
        extra_405 = ""
        if "405" in cause or "Method Not Allowed" in cause:
            extra_405 = (
                "\n- 检测到 HTTP 405：多见于代理不支持 HTTPS 隧道（CONNECT）、或代理地址/端口填错"
                "（例如指向了内网里的普通 Web 服务、与本工具使用的 Streamlit 端口 8501–8599 冲突）。"
                "可先清空代理直连验证；确认代理为支持 HTTPS 的正向/透明代理后再填回。"
            )
        return (
            "连接失败(APIConnectionError)：无法连通模型服务。\n"
            f"- 当前 Base URL: {base_url or '默认'}\n"
            f"- 当前模型: {model_name}\n"
            f"- 底层错误: {cause}\n"
            "- 请检查：网络是否可访问外网、公司代理/VPN、防火墙是否拦截。\n"
            "- 可在“网络高级设置”中配置代理，必要时临时开启“跳过SSL证书校验”排查。"
            f"{extra_405}"
        )
    if isinstance(e, AuthenticationError):
        return "鉴权失败(401)：API Key 无效、过期或复制不完整。"
    if isinstance(e, PermissionDeniedError):
        return "权限不足(403)：当前 Key 无该模型权限或项目权限不足。"
    if isinstance(e, NotFoundError):
        return f"资源不存在(404)：请检查 Base URL 或模型名（当前：{model_name}）。"
    if isinstance(e, RateLimitError):
        return "限流/配额不足(429)：请检查额度、并发限制或稍后重试。"
    if isinstance(e, BadRequestError):
        return f"请求参数错误(400)：{str(e)}"
    return f"模型调用失败：{str(e)}"


def test_model_connection(
    api_key: str,
    base_url: str,
    model_name: str,
    endpoint_mode: str = "auto",
    proxy_url: str = "",
    skip_ssl_verify: bool = False,
) -> str:
    api_key = (api_key or "").strip()
    if not api_key:
        return "请先填写 API Key。"
    try:
        _request_model_with_fallback(
            api_key=api_key,
            base_url=base_url,
            model_name=model_name,
            system_prompt="你是一个连接测试助手。只回复 pong。",
            user_prompt="ping",
            endpoint_mode=endpoint_mode,
            proxy_url=proxy_url,
            skip_ssl_verify=skip_ssl_verify,
            temperature=0,
            max_tokens=5,
        )
        return "连接测试成功：鉴权、网络、模型调用均正常。"
    except Exception as e:
        return _diagnose_exception(e, base_url=base_url, model_name=model_name)


def call_model_generate_cases(
    source_text: str,
    case_count_hint: int,
    include_p2: bool,
    model_name: str,
    api_key: str,
    base_url: str,
    endpoint_mode: str = "auto",
    proxy_url: str = "",
    skip_ssl_verify: bool = False,
) -> List[Dict]:
    api_key = (api_key or "").strip() or os.getenv("OPENAI_API_KEY", "").strip()
    if not api_key:
        raise RuntimeError("请先在侧边栏填写 API Key，或配置 OPENAI_API_KEY 环境变量。")

    prompt = build_prompt(source_text, case_count_hint, include_p2)

    try:
        content = _request_model_with_fallback(
            api_key=api_key,
            base_url=base_url,
            model_name=model_name,
            system_prompt="你是资深测试架构师，只输出 JSON 数组。",
            user_prompt=prompt,
            endpoint_mode=endpoint_mode,
            proxy_url=proxy_url,
            skip_ssl_verify=skip_ssl_verify,
            temperature=0.2,
        )
    except Exception as e:
        raise RuntimeError(_diagnose_exception(e, base_url=base_url, model_name=model_name)) from e
    if not content:
        raise RuntimeError("模型未返回内容，请重试。")

    data = parse_cases_from_content(content)
    if not data:
        # 二次纠偏：让模型把非标准输出转成纯 JSON 数组
        repaired_text = _repair_with_model(
            api_key=api_key,
            base_url=base_url,
            model_name=model_name,
            raw_text=content,
            endpoint_mode=endpoint_mode,
            proxy_url=proxy_url,
            skip_ssl_verify=skip_ssl_verify,
        )
        data = parse_cases_from_content(repaired_text)
    if not data:
        raise RuntimeError("模型返回格式仍不可解析（非标准JSON）。建议切换模型或稍后重试。")

    # 不足条数时自动补齐，最多补两轮
    all_rows = normalize_rows(data)
    all_rows = dedupe_cases(all_rows)
    rounds = 0
    while len(all_rows) < case_count_hint and rounds < 2:
        rounds += 1
        need = case_count_hint - len(all_rows)
        supplement_prompt = build_supplement_prompt(
            source_text=source_text,
            existing_rows=all_rows,
            need_count=need,
            include_p2=include_p2,
        )
        try:
            supplement_content = _request_model_with_fallback(
                api_key=api_key,
                base_url=base_url,
                model_name=model_name,
                system_prompt="你是资深测试架构师，只输出 JSON 数组。",
                user_prompt=supplement_prompt,
                endpoint_mode=endpoint_mode,
                proxy_url=proxy_url,
                skip_ssl_verify=skip_ssl_verify,
                temperature=0.2,
            )
            supplement_data = parse_cases_from_content(supplement_content)
            if not supplement_data:
                repaired_text = _repair_with_model(
                    api_key=api_key,
                    base_url=base_url,
                    model_name=model_name,
                    raw_text=supplement_content,
                    endpoint_mode=endpoint_mode,
                    proxy_url=proxy_url,
                    skip_ssl_verify=skip_ssl_verify,
                )
                supplement_data = parse_cases_from_content(repaired_text)
            if supplement_data:
                merged = all_rows + normalize_rows(supplement_data)
                all_rows = dedupe_cases(merged)
                all_rows = normalize_rows(all_rows)
            else:
                break
        except Exception:
            break

    return normalize_rows(all_rows)


def dataframe_to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="测试用例")
    output.seek(0)
    return output.read()


st.set_page_config(page_title="测试用例生成器", layout="wide")
st.title("测试用例生成器")
st.caption("支持导入 txt/docx/pdf 或粘贴文本，调用模型自动生成并导出 Excel。")

config = load_local_config()
saved_keys = config.get("provider_keys", {}) if isinstance(config, dict) else {}
saved_provider = config.get("provider_select", "Kimi(月之暗面)") if isinstance(config, dict) else "Kimi(月之暗面)"
saved_provider = saved_provider if saved_provider in PROVIDER_PRESETS else "Kimi(月之暗面)"
saved_base_urls = config.get("provider_base_urls", {}) if isinstance(config, dict) else {}
saved_models = config.get("provider_models", {}) if isinstance(config, dict) else {}
saved_endpoint_modes = config.get("provider_endpoint_modes", {}) if isinstance(config, dict) else {}
saved_proxy_url = config.get("proxy_url", "") if isinstance(config, dict) else ""
saved_skip_ssl_verify = bool(config.get("skip_ssl_verify", False)) if isinstance(config, dict) else False

if "provider_keys" not in st.session_state:
    st.session_state.provider_keys = saved_keys
if "provider_select" not in st.session_state:
    st.session_state.provider_select = saved_provider
if "last_provider" not in st.session_state:
    st.session_state.last_provider = st.session_state.provider_select
if "api_key_input" not in st.session_state:
    st.session_state.api_key_input = st.session_state.provider_keys.get(st.session_state.provider_select, "")
if "provider_base_urls" not in st.session_state:
    st.session_state.provider_base_urls = saved_base_urls
if "provider_models" not in st.session_state:
    st.session_state.provider_models = saved_models
if "provider_endpoint_modes" not in st.session_state:
    st.session_state.provider_endpoint_modes = saved_endpoint_modes
if "base_url_input" not in st.session_state:
    st.session_state.base_url_input = st.session_state.provider_base_urls.get(
        st.session_state.provider_select,
        PROVIDER_PRESETS[st.session_state.provider_select]["base_url"],
    )
if "model_name_input" not in st.session_state:
    st.session_state.model_name_input = st.session_state.provider_models.get(
        st.session_state.provider_select,
        PROVIDER_PRESETS[st.session_state.provider_select]["model"],
    )
if "endpoint_mode_input" not in st.session_state:
    st.session_state.endpoint_mode_input = st.session_state.provider_endpoint_modes.get(
        st.session_state.provider_select,
        "auto",
    )
if "proxy_url_input" not in st.session_state:
    st.session_state.proxy_url_input = saved_proxy_url
if "skip_ssl_verify_input" not in st.session_state:
    st.session_state.skip_ssl_verify_input = saved_skip_ssl_verify


def on_provider_change():
    prev = st.session_state.get("last_provider")
    new_provider = st.session_state.get("provider_select")
    if prev:
        st.session_state.provider_keys[prev] = st.session_state.get("api_key_input", "")
        st.session_state.provider_base_urls[prev] = st.session_state.get("base_url_input", "")
        st.session_state.provider_models[prev] = st.session_state.get("model_name_input", "")
        st.session_state.provider_endpoint_modes[prev] = st.session_state.get("endpoint_mode_input", "auto")
    st.session_state.api_key_input = st.session_state.provider_keys.get(new_provider, "")
    st.session_state.base_url_input = st.session_state.provider_base_urls.get(
        new_provider, PROVIDER_PRESETS[new_provider]["base_url"]
    )
    st.session_state.model_name_input = st.session_state.provider_models.get(
        new_provider, PROVIDER_PRESETS[new_provider]["model"]
    )
    st.session_state.endpoint_mode_input = st.session_state.provider_endpoint_modes.get(new_provider, "auto")
    st.session_state.last_provider = new_provider
    save_local_config(
        {
            "provider_select": st.session_state.provider_select,
            "provider_keys": st.session_state.provider_keys,
            "provider_base_urls": st.session_state.provider_base_urls,
            "provider_models": st.session_state.provider_models,
            "provider_endpoint_modes": st.session_state.provider_endpoint_modes,
            "proxy_url": st.session_state.get("proxy_url_input", ""),
            "skip_ssl_verify": st.session_state.get("skip_ssl_verify_input", False),
        }
    )


def on_api_key_change():
    provider = st.session_state.get("provider_select")
    st.session_state.provider_keys[provider] = st.session_state.get("api_key_input", "")
    save_local_config(
        {
            "provider_select": st.session_state.provider_select,
            "provider_keys": st.session_state.provider_keys,
            "provider_base_urls": st.session_state.provider_base_urls,
            "provider_models": st.session_state.provider_models,
            "provider_endpoint_modes": st.session_state.provider_endpoint_modes,
            "proxy_url": st.session_state.get("proxy_url_input", ""),
            "skip_ssl_verify": st.session_state.get("skip_ssl_verify_input", False),
        }
    )


def on_base_url_change():
    provider = st.session_state.get("provider_select")
    st.session_state.provider_base_urls[provider] = st.session_state.get("base_url_input", "")
    save_local_config(
        {
            "provider_select": st.session_state.provider_select,
            "provider_keys": st.session_state.provider_keys,
            "provider_base_urls": st.session_state.provider_base_urls,
            "provider_models": st.session_state.provider_models,
            "provider_endpoint_modes": st.session_state.provider_endpoint_modes,
            "proxy_url": st.session_state.get("proxy_url_input", ""),
            "skip_ssl_verify": st.session_state.get("skip_ssl_verify_input", False),
        }
    )


def on_model_name_change():
    provider = st.session_state.get("provider_select")
    st.session_state.provider_models[provider] = st.session_state.get("model_name_input", "")
    save_local_config(
        {
            "provider_select": st.session_state.provider_select,
            "provider_keys": st.session_state.provider_keys,
            "provider_base_urls": st.session_state.provider_base_urls,
            "provider_models": st.session_state.provider_models,
            "provider_endpoint_modes": st.session_state.provider_endpoint_modes,
            "proxy_url": st.session_state.get("proxy_url_input", ""),
            "skip_ssl_verify": st.session_state.get("skip_ssl_verify_input", False),
        }
    )


def on_endpoint_mode_change():
    provider = st.session_state.get("provider_select")
    st.session_state.provider_endpoint_modes[provider] = st.session_state.get("endpoint_mode_input", "auto")
    save_local_config(
        {
            "provider_select": st.session_state.provider_select,
            "provider_keys": st.session_state.provider_keys,
            "provider_base_urls": st.session_state.provider_base_urls,
            "provider_models": st.session_state.provider_models,
            "provider_endpoint_modes": st.session_state.provider_endpoint_modes,
            "proxy_url": st.session_state.get("proxy_url_input", ""),
            "skip_ssl_verify": st.session_state.get("skip_ssl_verify_input", False),
        }
    )


def on_network_change():
    save_local_config(
        {
            "provider_select": st.session_state.provider_select,
            "provider_keys": st.session_state.provider_keys,
            "provider_base_urls": st.session_state.provider_base_urls,
            "provider_models": st.session_state.provider_models,
            "provider_endpoint_modes": st.session_state.provider_endpoint_modes,
            "proxy_url": st.session_state.get("proxy_url_input", ""),
            "skip_ssl_verify": st.session_state.get("skip_ssl_verify_input", False),
        }
    )


with st.sidebar:
    st.subheader("配置")
    provider = st.selectbox(
        "API 提供方",
        list(PROVIDER_PRESETS.keys()),
        key="provider_select",
        on_change=on_provider_change,
    )
    api_key_input = st.text_input(
        "API Key",
        key="api_key_input",
        type="password",
        on_change=on_api_key_change,
        help="按模型分别记忆，并保存到本机当前用户目录。",
    )
    base_url_input = st.text_input(
        "Base URL",
        key="base_url_input",
        on_change=on_base_url_change,
        help="OpenAI 官方默认 https://api.openai.com/v1；智谱/Kimi 建议使用默认值",
    )
    model_name = st.text_input("模型名称", key="model_name_input", on_change=on_model_name_change)
    endpoint_mode = st.selectbox(
        "接口端点",
        list(ENDPOINT_PRESETS.values()),
        format_func=lambda mode: next(
            label for label, value in ENDPOINT_PRESETS.items() if value == mode
        ),
        key="endpoint_mode_input",
        on_change=on_endpoint_mode_change,
        help="你的模型配置显示为 /v1/responses 时请选择 Responses；不确定时使用自动识别。",
    )
    if provider == "Kimi(月之暗面)":
        st.caption("Kimi 建议：Base URL 使用 https://api.moonshot.cn/v1，模型如 moonshot-v1-8k。")
    if provider == "GPT(OpenAI)":
        st.caption("GPT 建议：Base URL 使用 https://api.openai.com/v1，模型如 gpt-4.1 / gpt-4o。")
    case_count_hint = st.number_input("期望最少用例数", min_value=10, max_value=500, value=20, step=5)
    include_p2 = st.checkbox("大版本模式（包含 P2 全量）", value=True)
    with st.expander("网络高级设置", expanded=False):
        st.caption(
            "访问 https 模型接口需要代理支持 HTTPS（CONNECT）。若出现 405，请核对代理端口是否真为代理服务，"
            "或先留空直连测试。"
        )
        st.text_input(
            "代理地址(可选)",
            key="proxy_url_input",
            on_change=on_network_change,
            placeholder="例如: http://127.0.0.1:7890",
        )
        st.checkbox(
            "跳过SSL证书校验(仅排查用)",
            key="skip_ssl_verify_input",
            on_change=on_network_change,
        )
    test_conn_btn = st.button("测试连接", use_container_width=True)

if test_conn_btn:
    with st.spinner("正在测试模型连接..."):
        result = test_model_connection(
            api_key=st.session_state.get("api_key_input", ""),
            base_url=st.session_state.get("base_url_input", ""),
            model_name=st.session_state.get("model_name_input", ""),
            endpoint_mode=st.session_state.get("endpoint_mode_input", "auto"),
            proxy_url=st.session_state.get("proxy_url_input", ""),
            skip_ssl_verify=bool(st.session_state.get("skip_ssl_verify_input", False)),
        )
        if result.startswith("连接测试成功"):
            st.success(result)
        else:
            st.error(result)

left, right = st.columns(2)
with left:
    uploaded = st.file_uploader("上传需求文档", type=["txt", "docx", "pdf"])
with right:
    pasted_text = st.text_area("或直接粘贴需求文本", height=180, placeholder="把需求流程、页面说明、交互规则粘贴到这里...")

source_text = ""
if uploaded is not None:
    try:
        source_text = parse_uploaded_file(uploaded)
        st.success(f"已读取文件：{uploaded.name}")
    except Exception as e:
        st.error(f"文件解析失败：{e}")

if pasted_text and pasted_text.strip():
    source_text = pasted_text.strip()

if source_text:
    with st.expander("识别到的需求文本预览", expanded=False):
        st.text_area("预览", source_text[:8000], height=260)

gen_btn = st.button("生成测试用例", type="primary", use_container_width=True)

if "generated_df" not in st.session_state:
    st.session_state.generated_df = None

if gen_btn:
    if not source_text.strip():
        st.warning("请先上传文档或粘贴需求文本。")
    else:
        with st.spinner("模型正在生成测试用例，请稍候..."):
            try:
                rows = call_model_generate_cases(
                    source_text=source_text,
                    case_count_hint=int(case_count_hint),
                    include_p2=include_p2,
                    model_name=model_name.strip() or DEFAULT_MODEL,
                    api_key=api_key_input,
                    base_url=base_url_input,
                    endpoint_mode=st.session_state.get("endpoint_mode_input", "auto"),
                    proxy_url=st.session_state.get("proxy_url_input", ""),
                    skip_ssl_verify=bool(st.session_state.get("skip_ssl_verify_input", False)),
                )
                df = pd.DataFrame(rows, columns=DEFAULT_COLUMNS)
                st.session_state.generated_df = df
                st.success(f"生成完成，共 {len(df)} 条测试用例。")
            except Exception as e:
                st.error(f"生成失败：{e}")

df_result = st.session_state.generated_df
if df_result is not None and not df_result.empty:
    st.subheader("生成结果")
    st.dataframe(df_result, use_container_width=True, height=500)

    excel_bytes = dataframe_to_excel_bytes(df_result)
    st.download_button(
        label="下载 Excel（.xlsx）",
        data=excel_bytes,
        file_name="测试用例生成结果.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    tsv_text = df_result.to_csv(index=False, sep="\t")
    st.download_button(
        label="下载 TSV（可粘贴 Excel）",
        data=tsv_text.encode("utf-8-sig"),
        file_name="测试用例生成结果.tsv",
        mime="text/tab-separated-values",
        use_container_width=True,
    )
