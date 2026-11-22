# -*- coding: utf-8 -*-
"""
批量文件重命名工具 · Web 后端（完整版）
======================================
纯 Python 标准库实现的 REST API + 静态文件服务器。

Copyright (C) 2026 批量文件重命名工具 作者

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program.  If not, see <https://www.gnu.org/licenses/>.

功能：
  - 文件夹管理（含 Windows 目录选择对话框）
  - 加载文件夹文件列表（递归 / 仅文件夹 / 过滤系统文件）
  - 命名规则实时预览（前缀+编号 / 编号+前缀 / 前缀+编号+后缀 / 分类+编号）
  - 添加 / 替换 / 正则替换 / 元数据占位符
  - 真实执行重命名（带撤销日志，可一键回退）
  - 模板保存 / 加载 / 删除；多方案管理
  - Excel 导出（openpyxl）
  - 图片内容识别（本地 CLIP 模型包 / 本地 Ollama 多模态 / 云端 AI API / 规则降级）
  - AI 内容命名（依据文件内容生成文件名，云端或本地 Ollama）

运行：
  python server.py [port]     默认端口 8890
  启动后访问 http://localhost:8890
"""
import os
import re
import io
import sys
import json
import time
import base64
import shutil
import datetime
import threading
import subprocess
import traceback
import urllib.request
import urllib.error
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse, unquote

# ============================================================
# 基础配置
# ============================================================
# PyInstaller 打包时，资源文件（webui.html、qr）位于 _MEIPASS 临时目录；
# 运行数据（data/）则固定放在 exe 同目录，保证用户数据可持久化。
if getattr(sys, "frozen", False):
    APP_DIR = os.path.dirname(sys.executable)          # exe 所在目录（用户数据）
    RES_DIR = getattr(sys, "_MEIPASS", APP_DIR)        # 打包资源目录
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))
    RES_DIR = APP_DIR

BASE_DIR = RES_DIR
DATA_DIR = os.path.join(APP_DIR, "data")
UNDO_DIR = os.path.join(DATA_DIR, "_undo_logs")
TEMPLATE_DIR = os.path.join(DATA_DIR, "templates")
SCHEME_DIR = os.path.join(DATA_DIR, "schemes")
AI_CONFIG_FILE = os.path.join(DATA_DIR, "ai_config.json")
LOG_DIR = os.path.join(DATA_DIR, "logs")
WEB_LOG_FILE = os.path.join(LOG_DIR, "web.log")
for _d in (UNDO_DIR, TEMPLATE_DIR, SCHEME_DIR, LOG_DIR):
    os.makedirs(_d, exist_ok=True)

AI_READ_TIMEOUT = 90
AI_LOCAL_TIMEOUT = 30  # 本地 Ollama 单次命名超时（秒），缩短以便「停止」更快响应
MAX_RETRIES = 2
LOCAL_DEFAULT_BASE = "http://localhost:11434"

CATEGORIES = ["人物", "动物", "风景", "建筑", "食物", "文档", "其他"]

IMAGE_EXT_MAP = {".jpg": "jpeg", ".jpeg": "jpeg", ".png": "png", ".bmp": "bmp",
                 ".webp": "webp", ".gif": "gif", ".tiff": "tiff"}

DOC_EXTS = {".doc", ".docx", ".pdf", ".txt", ".md", ".xlsx", ".xls", ".pptx", ".csv", ".html", ".htm"}
ARCHIVE_EXTS = {".zip", ".rar", ".7z", ".tar", ".gz", ".bz2", ".xz", ".iso"}
VIDEO_EXTS = {".mp4", ".avi", ".mkv", ".mov", ".wmv", ".flv", ".webm", ".m4v"}
AUDIO_EXTS = {".mp3", ".wav", ".flac", ".aac", ".ogg", ".m4a", ".wma"}
SYSTEM_FILENAMES = {"desktop.ini", "thumbs.db", "ehthumbs.db"}

TEXT_EXTS = {".txt", ".md", ".log", ".ini", ".cfg", ".conf", ".yml", ".yaml", ".toml"}
CODE_EXTS = {".py", ".js", ".ts", ".tsx", ".jsx", ".vue", ".svelte", ".c", ".cpp", ".h", ".hpp",
             ".java", ".cs", ".rb", ".php", ".pl", ".sh", ".bat", ".ps1", ".sql",
             ".css", ".scss", ".less", ".go", ".rs", ".swift", ".kt", ".scala", ".r", ".lua", ".dart"}
OFFICE_EXTS = {".docx", ".doc", ".xlsx", ".xls", ".pptx", ".rtf"}
BOOK_EXTS = {".epub", ".mobi", ".azw3", ".fb2"}
SUPPORTED_CONTENT_EXTS = TEXT_EXTS | CODE_EXTS | OFFICE_EXTS | ARCHIVE_EXTS | BOOK_EXTS | \
                         {".json", ".xml", ".csv", ".html", ".htm", ".pdf"}

# 云端 AI 服务商（OpenAI 兼容 Chat Completion 接口）
PROVIDERS = {
    "OpenAI": {"endpoint": "https://api.openai.com/v1/chat/completions",
               "default_model": "gpt-4o-mini",
               "models": ["gpt-4o", "gpt-4o-mini"],
               "vision": True},
    "DeepSeek": {"endpoint": "https://api.deepseek.com/v1/chat/completions",
                 "default_model": "deepseek-chat",
                 "models": ["deepseek-chat", "deepseek-v3", "deepseek-r1"],
                 "vision": False},
    "智谱 GLM": {"endpoint": "https://open.bigmodel.cn/api/paas/v4/chat/completions",
                 "default_model": "glm-4v-flash",
                 "models": ["glm-4v-flash", "glm-4v-plus", "glm-4-flash", "glm-4-plus"],
                 "vision": True},
    "通义千问": {"endpoint": "https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions",
                 "default_model": "qwen-vl-plus",
                 "models": ["qwen-vl-plus", "qwen-vl-max", "qwen-turbo-latest", "qwen-plus", "qwen-max"],
                 "vision": True},
    "硅基流动": {"endpoint": "https://api.siliconflow.cn/v1/chat/completions",
                 "default_model": "Qwen/Qwen3-VL-8B-Instruct",
                 "models": ["Qwen/Qwen3-VL-8B-Instruct", "Qwen/Qwen3-VL-32B-Instruct",
                            "Qwen/Qwen2.5-72B-Instruct", "deepseek-ai/DeepSeek-V3",
                            "zai-org/GLM-4.5V", "Qwen/Qwen-Image"],
                 "vision": True},
    "腾讯混元": {"endpoint": "https://api.hunyuan.cloud.tencent.com/v1/chat/completions",
                 "default_model": "hunyuan-vision",
                 "models": ["hunyuan-vision", "hunyuan-standard", "hunyuan-lite"],
                 "vision": True},
    "自定义": {"endpoint": "", "default_model": "", "models": [], "vision": True},
}
PROVIDER_NAMES = list(PROVIDERS.keys())

AI_CATEGORIZE_FOR_TAGS = "请将这张图片分类到以下类别之一，只返回类别名称：\n\n{categories}\n"
AI_SELF_DESCRIBE_PROMPT = ("【任务】请识别这张图片的主要内容和特征，生成一段简短的中文描述。"
                           "\n【描述要求】1.主要物体/场景 2.场景氛围 3.其他显著特征"
                           "\n【约束】只返回纯文本描述，不要换行，不要额外说明，保持简洁。\n直接返回描述即可。")
AI_TITLE_PROMPT = "请用简短的词语或短语概括这张图片的主要内容，作为图片标题（不超过20字），只返回标题本身。"

# 默认 AI 配置
DEFAULT_AI_CONFIG = {
    # 识别模式: rule / clip / ollama / api
    "recognition_mode": "rule",
    # 云端 API
    "api_provider": "OpenAI",
    "api_key": "",
    "api_model": "gpt-4o-mini",
    "api_endpoint": "",
    # Ollama
    "ollama_base": LOCAL_DEFAULT_BASE,
    "ollama_model": "",
    # 分类设置
    "classify_mode": "固定分类",          # 固定分类 / 自定义分类 / AI自由识别
    "custom_categories": "人物,动物,风景,建筑,食物,文档,其他",
    "free_style": "图片描述",              # 图片描述 / 图片标题
    "use_multiple_tags": False,
    "tag_separator": "-",
    "max_tag_count": 5,
    "batch_size": 8,
    "max_concurrent": 2,
    # AI 内容命名
    "ai_content_mode": "云端API",          # 云端API / 本地Ollama
    "ai_content_prompt": "你是一个文件命名助手。请为下面的文件或文件夹生成一个简洁贴切的新名字（建议8-15字，中文优先）。\n"
                         "判断依据：如果「内容或清单」是正文，则为文件起名；如果内容是「文件夹内文件清单」，则为文件夹起名。\n"
                         "原名称: {filename}\n内容或清单: {content}\n"
                         "只返回新名字本身，不要扩展名、编号、解释或引号。",
    "ai_content_limit": "auto",           # auto / 500 / 1000 / 2000 / 5000
    "ai_batch_opt": "batch",              # single / batch
    # AI 内容命名页独立云端配置（与图片识别页互不影响，留空时回退到 api_*）
    "ai_naming_provider": "",
    "ai_naming_key": "",
    "ai_naming_model": "",
    "ai_naming_endpoint": "",
    "local_model": "",
    "local_api_base": LOCAL_DEFAULT_BASE,
    "local_temperature": 0.7,
    "local_max_tokens": 512,
}


# ============================================================
# AI 配置读写
# ============================================================
def load_ai_config():
    try:
        with open(AI_CONFIG_FILE, "r", encoding="utf-8") as f:
            saved = json.load(f)
        cfg = dict(DEFAULT_AI_CONFIG)
        cfg.update(saved)
        return cfg
    except Exception:
        return dict(DEFAULT_AI_CONFIG)


def save_ai_config(cfg):
    with open(AI_CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


# ============================================================
# 核心逻辑（移植自桌面端）
# ============================================================

def sanitize_filename(name):
    if not name:
        return "文件"
    illegal_chars = '\\/:*?"<>|'
    cleaned = name
    for ch in illegal_chars:
        cleaned = cleaned.replace(ch, "_")
    # 控制字符（\n \t \x00 等）Windows 文件名不允许，之前靠正则替换插入的换行会让改名报「失败: 语法」
    cleaned = re.sub(r"[\x00-\x1f\x7f]", "_", cleaned)
    cleaned = cleaned.strip().rstrip(". ")
    cleaned = re.sub("_+", "_", cleaned).strip("_")
    reserved = {"con", "prn", "aux", "nul", *(f"com{i}" for i in range(1, 10)),
                *(f"lpt{i}" for i in range(1, 10))}
    if os.path.splitext(cleaned)[0].lower() in reserved:
        cleaned = "_" + cleaned
    if len(cleaned) > 255:
        cleaned = cleaned[:255]
    return cleaned or "文件"


MAX_FULL_PATH = 259   # Windows 未开启长路径支持时完整路径上限（含盘符）


def fit_filename(folder, base, ext):
    """把「base+ext」裁剪到 Windows 能落盘的长度，返回裁剪后的 base。
    之前只把 base 截到 255 字符，没算扩展名与目录长度，前缀很长时真实改名会
    报「失败: 语法」（WinError 123 文件名/路径过长）。"""
    ext = ext or ""
    base = base or ""
    # 单个文件名（含扩展名）上限 255 字符
    max_base = max(1, 255 - len(ext))
    if len(base) > max_base:
        base = base[:max_base]
    # 完整路径上限：目录已经很深时继续裁剪文件名，避免整次改名失败
    try:
        full_len = len(os.path.join(folder or "", base + ext))
        if full_len > MAX_FULL_PATH:
            base = base[:max(1, len(base) - (full_len - MAX_FULL_PATH))]
    except Exception:
        pass
    return base or "文件"


def normalize_ext(raw):
    """把用户/表格给的扩展名规范化成「.md」「.tar.gz」这种形式；无效返回 ""。

    - 允许 "md" / ".md" / "*.md" / " .MD "；
    - 允许多段（.tar.gz），但最长 3 段、每段 ≤10 个字母数字；
    - 纯数字（如 "v1.2" 里的 ".2"）不算扩展名，避免把版本号当扩展名。
    """
    if raw is None:
        return ""
    e = str(raw).strip().strip('"').strip("'").strip()
    if not e:
        return ""
    e = e.lstrip("*").strip()
    e = re.sub(r"^[.\s]+", "", e)
    e = re.sub(r"[^A-Za-z0-9.]", "", e)
    parts = [p for p in e.split(".") if p]
    if not parts or len(parts) > 3:
        return ""
    if all(p.isdigit() for p in parts):
        return ""
    if any(len(p) > 10 for p in parts):
        return ""
    return "." + ".".join(parts)


def split_target_name(name, old_ext, is_folder):
    """把「用户给定的完整新名」拆成 (基础名, 扩展名)。

    规则：
      - 文件夹没有扩展名，整名都是基础名；
      - 名字末尾是合法扩展名（.md / .PDF / .tar.gz）时以它为准 —— Excel 导入映射改扩展名靠这条；
      - 否则整名当基础名，沿用原扩展名 —— 避免把「v1.2 报告」「2024.年终总结」里的点当扩展名。
    """
    name = (name or "").strip()
    if is_folder:
        return name, ""
    raw_ext = os.path.splitext(name)[1]
    ext = normalize_ext(raw_ext) if raw_ext else ""
    if ext:
        base = name[:len(name) - len(raw_ext)]
        if base:
            return base, ext
    return name, (old_ext or "")


def update_state_after_move(old_path, new_path):
    """文件/文件夹在磁盘上改名后，同步内存状态。

    涉及两处：
      - STATE.file_map：old_path / old / folder（folder 覆盖被改名目录的全部子孙条目）
      - STATE.files：条目自身的 old，以及以被改名目录为父目录的 folder

    之前 perform_rename 只更新了被改名的那一条，且用「改名后 os.path.isdir(old_path)」
    判断是否文件夹（此时 old_path 早已不存在，永远为假），因此文件夹改名后
    子条目的 folder 仍是旧路径，下一步操作一律「失败: 找不到」；
    perform_undo 则完全不更新内存，撤销后任何操作都会「找不到」。
    """
    old_base = os.path.normpath(old_path)
    new_base = os.path.normpath(new_path)
    old_norm = os.path.normcase(old_base)
    prefix = old_norm + os.sep

    def _map(p):
        if not p:
            return None
        norm = os.path.normcase(os.path.normpath(p))
        if norm == old_norm:
            return new_base
        if norm.startswith(prefix):
            rel = os.path.normpath(p)[len(old_base):]
            return os.path.normpath(new_base + rel)
        return None

    with STATE.lock:
        for entry in STATE.file_map.values():
            mapped = _map(entry.get("old_path"))
            if mapped:
                entry["old_path"] = mapped
                entry["old"] = os.path.basename(mapped)
            mapped_folder = _map(entry.get("folder"))
            if mapped_folder:
                entry["folder"] = mapped_folder
        for sf in STATE.files:
            folder = sf.get("folder") or ""
            if os.path.normcase(os.path.normpath(os.path.join(folder, sf.get("old") or ""))) == old_norm:
                sf["old"] = os.path.basename(new_base)
                sf["relpath"] = sf["old"]
                sf["ext"] = os.path.splitext(sf["old"])[1].lower()
                if not sf.get("_is_folder"):
                    sf["cat"] = guess_category(sf["old"])
                continue
            mapped_folder = _map(folder)
            if mapped_folder:
                sf["folder"] = mapped_folder


def replace_meta_placeholders(text, meta, index=None):
    if not text:
        return text
    try:
        result = text
        for key in ("创建日期", "修改日期", "文件大小", "文件扩展名",
                    "文件扩展名-带点", "作者", "所有者", "原文件名"):
            # 未解析的占位符替换为空，避免 {作者} 等字面量进入文件名
            val = (meta or {}).get(key) or ""
            result = result.replace("{%s}" % key, val)
            result = result.replace("【%s】" % key, val)
        if index is not None:
            result = re.sub(r"\{ID:(\d+)\}", lambda m: str(index + 1).zfill(int(m.group(1))), result)
            result = result.replace("{ID}", str(index + 1))
        return result
    except Exception:
        return text


def read_file_metadata(filepath, orig_name=None):
    meta = {}
    try:
        st = os.stat(filepath)
    except Exception:
        return meta
    # 逐字段容错：单个字段（如某些网络/共享盘上的创建时间）异常时，其余元数据仍可用
    try:
        meta["创建日期"] = datetime.datetime.fromtimestamp(st.st_ctime).strftime("%Y-%m-%d")
    except Exception:
        try:
            meta["创建日期"] = datetime.datetime.fromtimestamp(st.st_mtime).strftime("%Y-%m-%d")
        except Exception:
            meta["创建日期"] = ""
    try:
        meta["修改日期"] = datetime.datetime.fromtimestamp(st.st_mtime).strftime("%Y-%m-%d")
    except Exception:
        meta["修改日期"] = ""
    size = st.st_size
    if size < 1024:
        meta["文件大小"] = f"{size} B"
    elif size < 1048576:
        meta["文件大小"] = f"{size // 1024} KB"
    else:
        meta["文件大小"] = f"{size / 1048576:.1f} MB"
    ext_dot = os.path.splitext(filepath)[1].lower()
    meta["文件扩展名"] = ext_dot.lstrip(".") if ext_dot else "(无)"
    meta["文件扩展名-带点"] = ext_dot if ext_dot else "(无)"
    # 原文件名：优先使用首次加载的原始文件名（重命名后 old 已变，但占位符应保持原始名）
    base_name = os.path.basename(orig_name or filepath)
    meta["原文件名"] = os.path.splitext(base_name)[0]
    try:
        author, owner = _win_file_author(filepath)
        if author:
            meta["作者"] = author
        if owner:
            meta["所有者"] = owner
    except Exception:
        pass
    return meta


def _win_file_author(filepath):
    """Windows 下尽力获取文件作者 / 所有者（无依赖，失败返回空）"""
    author, owner = "", ""
    if os.name != "nt":
        return author, owner
    try:
        import ctypes
        from ctypes import wintypes

        def _sid_to_name(sid):
            try:
                import win32security  # pywin32 可能未安装
                name, _, _ = win32security.LookupAccountSid(None, sid)
                return name
            except Exception:
                return ""

        # 所有者：通过文件安全描述符获取（Windows API）
        try:
            SECURITY_INFORMATION = 0x00000001  # OWNER_SECURITY_INFORMATION
            h = ctypes.windll.kernel32.CreateFileW(
                filepath, 0x00000000, 0x00000001 | 0x00000002,
                None, 3, 0x00000080, None)  # OPEN_EXISTING | FILE_FLAG_BACKUP_SEMANTICS
            if h and h != -1:
                try:
                    psize = wintypes.DWORD()
                    ctypes.windll.advapi32.GetFileSecurityW(filepath, SECURITY_INFORMATION,
                                                            None, 0, ctypes.byref(psize))
                    buf = ctypes.create_string_buffer(psize.value)
                    ctypes.windll.advapi32.GetFileSecurityW(filepath, SECURITY_INFORMATION,
                                                            buf, psize.value, ctypes.byref(psize))
                    psid = ctypes.POINTER(wintypes.BYTE)()
                    ctypes.windll.advapi32.GetSecurityDescriptorOwner(buf, ctypes.byref(psid),
                                                                      ctypes.byref(wintypes.BOOL()))
                    if psid:
                        owner = _sid_to_name(psid) or owner
                finally:
                    ctypes.windll.kernel32.CloseHandle(h)
        except Exception:
            pass
    except Exception:
        pass
    return author, owner


def _py_replacement(repl):
    """将 JS 风格替换引用（$1、$&）转换为 Python 语法（\\g<1>、\\g<0>）。
    前端正则预设使用 $1 等写法，而 Python re.sub 仅支持 \\1 / \\g<1>，
    若不转换，「提取括号内容」等预设会在真实重命名时输出字面量 $1。"""
    if not repl:
        return repl
    repl = re.sub(r"\$(\d+)", lambda m: r"\g<%s>" % m.group(1), repl)
    repl = repl.replace("$&", r"\g<0>")
    return repl


def apply_regex(text, scheme):
    if scheme.get("regex_enabled") and scheme.get("regex_find"):
        if scheme.get("regex_lowercase"):
            text = text.lower()
        try:
            flags = 0 if scheme.get("regex_case_sensitive", True) else re.IGNORECASE
            text = re.sub(scheme["regex_find"], _py_replacement(scheme.get("regex_replace", "")), text, flags=flags)
        except re.error as e:
            # 正则写错时不要让用户以为已经生效
            STATE.log(f"正则替换未生效（语法错误：{e}）：{scheme.get('regex_find')!r}", "warning")
    return text


def build_new_name(scheme, index, meta=None, force_numbering=False):
    mode = scheme.get("mode", "prefix_num")
    numbering_on = force_numbering or bool(scheme.get("numbering_on", True))
    num = int(scheme.get("start_num", 1)) + index
    try:
        pad = int(scheme.get("pad_width") or 2) or 2
    except (TypeError, ValueError):
        pad = 2
    pad = min(20, max(1, pad))  # 可自定义 1-20 位
    padded = str(num).zfill(pad)
    prefix = replace_meta_placeholders(str(scheme.get("prefix", "文件") or "文件"), meta, index=index)
    suffix_raw = replace_meta_placeholders(str(scheme.get("suffix", "") or ""), meta, index=index)
    prefix = prefix.replace("{编号}", padded)
    suffix_raw = suffix_raw.replace("{编号}", padded)
    prefix = sanitize_filename(prefix)
    # 空后缀保持为空（sanitize_filename("") 会返回"文件"，需先判断）
    suffix = sanitize_filename(suffix_raw) if suffix_raw.strip() else ""
    if not prefix:
        prefix = "文件"
    if not numbering_on:
        # 编号功能关闭：仅应用前缀/后缀，不添加编号
        return f"{prefix}_{suffix}" if suffix else prefix
    if mode == "num_prefix":
        return f"{padded}_{prefix}"
    if mode == "prefix_num_suffix":
        return f"{prefix}_{padded}_{suffix}" if suffix else f"{prefix}_{padded}"
    return f"{prefix}_{padded}"


def get_unique_path(folder, new_base, ext, old_name=None):
    cand = f"{new_base}{ext}"
    # 目标名与原文件名一致（Windows 大小写不敏感）：允许同名（即文件自身），不算冲突，不追加 _1
    if old_name and cand.lower() == old_name.lower():
        return cand
    if not os.path.exists(os.path.join(folder, cand)):
        return cand
    counter = 1
    while True:
        cand2 = f"{new_base}_{counter}{ext}"
        if old_name and cand2.lower() == old_name.lower():
            counter += 1
            continue
        if not os.path.exists(os.path.join(folder, cand2)):
            return cand2
        counter += 1


def dedupe_candidate(folder, cand, old_name, used_set):
    """批内去重：目标名被本批其他文件占用（或磁盘上已存在别的同名文件）时追加 _N。

    原先只查 used_set，不看磁盘：目录里已有未参与本批的 X_1.txt 时，
    「预览」给出的 X_1.txt 实际不可用，真实改名会「失败: 已存在」。
    """
    base0, ext0 = os.path.splitext(cand)
    n = 1
    self_name = (old_name or "").lower()

    def _taken(name):
        low = name.lower()
        if low in used_set:
            return True
        if low == self_name:
            return False
        try:
            return os.path.exists(os.path.join(folder, name))
        except Exception:
            return False

    while _taken(cand):
        cand = f"{base0}_{n}{ext0}"
        n += 1
    return cand


def guess_category(path):
    ext = os.path.splitext(path)[1].lower()
    if ext in IMAGE_EXT_MAP:
        return "图片"
    if ext in DOC_EXTS:
        return "文档"
    if ext in ARCHIVE_EXTS:
        return "压缩包"
    if ext in VIDEO_EXTS:
        return "视频"
    if ext in AUDIO_EXTS:
        return "音频"
    return "其他"


def is_system_file(filepath):
    name = os.path.basename(filepath).lower()
    if name in SYSTEM_FILENAMES:
        return True
    if os.path.splitext(name)[1] in (".tmp", ".temp", ".ini"):
        return True
    try:
        import ctypes
        attrs = ctypes.windll.kernel32.GetFileAttributesW(filepath)
        if attrs != -1 and attrs & 6:
            return True
    except Exception:
        pass
    return False


# ============================================================
# 应用状态
# ============================================================
class AppState:
    def __init__(self):
        self.lock = threading.Lock()
        self._id_seq = 0          # 文件 id 全局单调递增：重新加载新文件夹后旧 id 不会「复活」成新文件
        self.folders = []
        self.file_map = {}      # id -> {folder, old, old_path}
        self.files = []         # 最近加载的快照
        self.logs = []
        self.recognize_stop = False   # 停止识别标志
        self.recognizing = False      # 是否正在识别
        self.rename_stop = False      # 停止重命名标志
        self.renaming = False         # 是否正在重命名
        self.ai_naming_stop = False   # 停止 AI 内容命名标志
        self.ai_naming = False        # 是否正在 AI 内容命名
        self.ai_naming_total = 0      # AI 命名总数（用于进度）
        self.ai_naming_done = 0       # AI 命名已完成数
        self.ai_naming_result = None  # AI 命名最终结果 {results, stopped, failed, error_detail}
        self.ai_naming_error = None   # AI 命名整体失败原因（轮询接口透出给前端）

    def next_id(self):
        with self.lock:
            fid = self._id_seq
            self._id_seq += 1
            return fid

    def log(self, msg, level="info"):
        """记录操作日志：内存(500条) + 持久化到 data/logs/web.log"""
        ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        entry = {"time": ts, "level": level, "msg": str(msg)}
        with self.lock:
            self.logs.append(entry)
            if len(self.logs) > 500:
                self.logs[:100] = []
            try:
                with open(WEB_LOG_FILE, "a", encoding="utf-8") as f:
                    f.write(f"[{ts}] [{level.upper():7}] {msg}\n")
            except Exception:
                pass

    def get_logs(self):
        with self.lock:
            return list(self.logs)

    def clear_logs(self):
        with self.lock:
            self.logs = []
            try:
                with open(WEB_LOG_FILE, "w", encoding="utf-8") as f:
                    f.write("")
            except Exception:
                pass
            return True


STATE = AppState()

# 程序退出控制（exe 无窗口模式）：浏览器关闭标签页 / 点击退出按钮时结束进程
_exit_lock = threading.Lock()
_exit_timer = None
_EXIT_DELAY_AUTO = 10.0    # 自动退出（关标签页）延迟：给页面刷新留窗口
_EXIT_DELAY_MANUAL = 1.5   # 按钮退出延迟：给前端收尾留时间


def _cancel_pending_exit():
    """取消待执行的退出（页面刷新 / 重新活跃时调用）"""
    global _exit_timer
    with _exit_lock:
        if _exit_timer is not None:
            try:
                _exit_timer.cancel()
            except Exception:
                pass
            _exit_timer = None
            STATE.log("已取消退出（页面重新活跃）", "info")


def _schedule_exit(delay, reason):
    """延迟后退出程序（exe 模式下彻底结束进程）"""
    global _exit_timer
    with _exit_lock:
        # 取消旧定时器（内联实现，避免嵌套获取非重入锁造成死锁）
        if _exit_timer is not None:
            try:
                _exit_timer.cancel()
            except Exception:
                pass
            _exit_timer = None
        STATE.log(f"收到退出请求（{reason}），{delay:.0f} 秒后关闭程序", "warning")

        def _do_exit():
            STATE.log("程序即将关闭", "info")
            # 正在重命名时先等它做完（最多 60 秒），避免留下半批改名
            for _ in range(120):
                if not STATE.renaming:
                    break
                time.sleep(0.5)
            # 强制结束进程（exe 无窗口模式下最可靠），避免 shutdown() 阻塞等待
            try:
                svr = _current_server
                if svr is not None:
                    # 在新线程中优雅停止 serve_forever，随后强制退出
                    threading.Thread(target=svr.shutdown, daemon=True).start()
            except Exception:
                pass
            time.sleep(0.2)
            os._exit(0)

        _exit_timer = threading.Timer(delay, _do_exit)
        _exit_timer.daemon = True
        _exit_timer.start()


_current_server = None  # 由 main() 注入，供退出流程调用 shutdown()


# ============================================================
# 文件扫描
# ============================================================
def normalize_folder(path):
    """统一文件夹写法（去空白、规范分隔符、去尾部反斜杠），
    否则「D:\\a」与「D:\\a\\」会被当成两个不同文件夹，重复添加且列表里出现两条。"""
    if path is None:
        return ""
    p = str(path).strip().strip('"')
    if not p:
        return ""
    try:
        return os.path.normpath(p)
    except Exception:
        return p


def _same_folder(a, b):
    try:
        return os.path.normcase(os.path.normpath(a)) == os.path.normcase(os.path.normpath(b))
    except Exception:
        return a == b


def scan_files(folders, recursive, only_folders, filter_system):
    entries = []
    # 每次重新加载都从空白状态开始：若不清空 file_map，上一批文件留下的旧 id
    # 依然指向旧文件夹里的真实文件，前端拿着过期 id 操作时会「成功改名另一个
    # 文件夹里的文件」（跨文件夹误改名）。id 由 STATE.next_id() 单调递增，
    # 因此清空后旧 id 不会再复活成新文件。
    with STATE.lock:
        STATE.file_map.clear()
        STATE.files = []

    def _make_entry(path, name, depth, is_folder):
        fid = STATE.next_id()
        try:
            st = os.stat(path)
            size, mtime = st.st_size, datetime.datetime.fromtimestamp(st.st_mtime).strftime("%Y-%m-%d %H:%M")
        except Exception:
            size, mtime = 0, ""
        ext = os.path.splitext(name)[1].lower()
        with STATE.lock:
            # orig：首次加载的原始文件名，重命名后 old 会更新为新名，但 {原文件名} 元数据始终引用 orig
            # is_folder：必须一起存，后续「文件夹改名要同步子条目」「文件夹名里的点不是扩展名」都依赖它
            STATE.file_map[fid] = {"folder": os.path.dirname(path), "old": name, "old_path": path,
                                   "orig": name, "is_folder": is_folder}
        return {"id": fid, "old": name, "ext": ext, "relpath": name,
                "size": size, "mtime": mtime, "cat": guess_category(name) if not is_folder else "文件夹",
                "folder": os.path.dirname(path), "_is_folder": is_folder}

    def _is_link_like(path):
        """junction / 符号链接：递归进入会跑出所选目录（甚至改到目录外的文件），一律不跟进"""
        try:
            if os.path.islink(path):
                return True
            st = os.lstat(path)
            return bool(getattr(st, "st_file_attributes", 0) & 0x400)  # FILE_ATTRIBUTE_REPARSE_POINT
        except Exception:
            return False

    def walk(folder, depth=0):
        try:
            names = sorted(os.listdir(folder))
        except Exception:
            return
        for name in names:
            path = os.path.join(folder, name)
            if _is_link_like(path):
                continue
            if filter_system and is_system_file(path):
                continue
            if os.path.isdir(path):
                if only_folders:
                    entries.append(_make_entry(path, name, depth, True))
                if recursive:
                    walk(path, depth + 1)
            elif os.path.isfile(path):
                if not only_folders:
                    entries.append(_make_entry(path, name, depth, False))

    for folder in folders:
        if os.path.isdir(folder):
            walk(folder, 0)
    with STATE.lock:
        STATE.files = entries
    return entries


# ============================================================
# 命名预览 / 重命名 / 撤销
# ============================================================
def compute_preview(ids, scheme):
    results = {}
    with STATE.lock:
        file_map = dict(STATE.file_map)
    scheme = scheme or {}
    scheme_indices = {}
    used_targets = {}  # folder -> 本批已占用的目标名集合（小写），用于批内去重
    for fid in ids:
        info = file_map.get(fid)
        if not info:
            results[fid] = {"new": "", "status": "未知文件"}
            continue
        try:
            folder, old = info["folder"], info["old"]
            is_folder = bool(info.get("is_folder"))
            if is_folder:
                # 文件夹本身没有扩展名，整名参与命名；否则「报告.v2」会被拆成
                # 基础名「报告」+ 扩展名「.v2」，预览/改名结果变成「前缀_1.v2」
                name, ext = old, ""
            else:
                name, ext = os.path.splitext(old)

            # 命名规则是否启用（前端 getScheme 传入 rules_on，兼容旧方案 rules_enabled）
            # 需同时满足：功能开关打开 且 已显式配置规则（前缀/后缀/非默认模式）；
            # 未显式配置时保留原文件名，让「添加&替换/正则」独立生效（与前端提示一致）
            # 统一修改扩展名（留空=保持原扩展名；文件夹不处理）
            target_ext = ext
            if not is_folder:
                new_ext = normalize_ext(scheme.get("new_ext"))
                if new_ext:
                    target_ext = new_ext

            rules_enabled = bool(scheme.get("rules_enabled", True))
            rules_on = bool(scheme.get("rules_on", True)) and rules_enabled
            numbering_on = bool(scheme.get("numbering_on", True))
            add_replace_on = bool(scheme.get("add_replace_on", True))
            old_name = name

            # 分类+编号模式：基于文件的 category 字段
            if scheme.get("use_category") and scheme.get("categories"):
                cat = scheme["categories"].get(str(fid)) or "其他"
                cat = sanitize_filename(cat)
                key = cat
                scheme_indices.setdefault(key, 0)
                num = int(scheme.get("start_num", 1)) + scheme_indices[key]
                try:
                    pad = int(scheme.get("pad_width") or 2) or 2
                except (TypeError, ValueError):
                    pad = 2
                pad = min(20, max(1, pad))
                padded = str(num).zfill(pad)
                new_base = f"{cat}_{padded}"
                scheme_indices[key] += 1
            # 按扩展名分组编号：同一扩展名使用连续编号（「按扩展名编号」为独立选项，不受编号开关影响）
            elif scheme.get("group_by_ext"):
                key = ext.lower() or "(无扩展名)"
                scheme_indices.setdefault(key, 0)
                meta = read_file_metadata(os.path.join(folder, old), orig_name=info.get("orig"))
                new_base = build_new_name(scheme, scheme_indices[key], meta=meta, force_numbering=True)
                scheme_indices[key] += 1
            # 命名规则已配置：前缀/编号生成新名
            elif rules_on:
                scheme_indices.setdefault("default", 0)
                meta = read_file_metadata(os.path.join(folder, old), orig_name=info.get("orig"))
                new_base = build_new_name(scheme, scheme_indices["default"], meta=meta)
                scheme_indices["default"] += 1
            # 命名规则未配置：以原文件名为基础（此时「添加&替换/正则」独立生效）
            else:
                new_base = old_name

            # 添加内容（前/后）：直接拼在基础名前后，不自动加 "_" 等分隔符。
            # 需要分隔符时由用户在「添加内容」里自己写（如「2024_」）。
            # 只裁掉朝外那一侧的空格（末尾/开头空格 Windows 本来也会去掉），
            # 贴合一侧的空格保留，这样「旅行 」+「IMG」仍能得到「旅行 IMG」。
            add_text_raw = str(scheme.get("add_text") or "")
            if add_replace_on and add_text_raw.strip():
                if scheme.get("add_position") == "前":
                    new_base = f"{add_text_raw.lstrip()}{new_base}"
                else:
                    new_base = f"{new_base}{add_text_raw.rstrip()}"
            # 查找替换：作用于基础名（受「添加&替换」开关控制）
            # find_mode：包含（替换所有出现）/ 开头是（文件名以目标开头，只替换靠前的第一个）/ 结尾是（文件名以目标结尾，只替换靠后的最后一个）
            find_text = (scheme.get("find_text") or "").strip()
            if add_replace_on and find_text:
                fmode = scheme.get("find_mode") or "包含"
                repl = scheme.get("replace_text", "")
                if fmode == "开头是":
                    if new_base.startswith(find_text):
                        # 文件名以目标开头：仅替换第一个（靠前）出现，其余保留
                        new_base = new_base.replace(find_text, repl, 1)
                elif fmode == "结尾是":
                    if new_base.endswith(find_text):
                        # 文件名以目标结尾：仅替换最后一个（靠后）出现，其余保留
                        idx = new_base.rfind(find_text)
                        new_base = new_base[:idx] + repl + new_base[idx + len(find_text):]
                else:
                    if find_text in new_base:
                        new_base = new_base.replace(find_text, repl)

            new_base = apply_regex(new_base, scheme)
            new_base = sanitize_filename(new_base) or "文件"
            # 裁到可落盘长度（要算上扩展名与目录路径，否则真实改名报「失败: 语法」）
            new_base = fit_filename(folder, new_base, target_ext)
            # old_name：目标名与原文件名一致时允许同名（自身不算冲突）
            new_name = get_unique_path(folder, new_base, target_ext, old_name=old)
            # 批内去重：本批其他文件已占用该目标名（且非自身）时自动加 _N，保证预览与执行一致
            used_set = used_targets.setdefault(folder, set())
            cand = dedupe_candidate(folder, new_name, old, used_set)
            if cand.lower() != old.lower():
                used_set.add(cand.lower())
            # 名称没有任何变化时提前告知（真实重命名会跳过，不再虚报「成功」）
            results[fid] = {"new": cand,
                            "status": "跳过: 名称未变化" if cand == old else "待重命名"}
        except Exception as e:
            results[fid] = {"new": "", "status": f"错误: {e}"}
    return results


def save_undo_log(operations):
    ts = time.strftime("%Y-%m-%d %H:%M:%S")
    ts_file = time.strftime("%Y%m%d_%H%M%S")
    data = {"timestamp": ts, "operation_count": len(operations), "operations": operations}
    with open(os.path.join(UNDO_DIR, "_undo_latest.json"), "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    with open(os.path.join(UNDO_DIR, f"_undo_{ts_file}.json"), "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def has_undo():
    return os.path.exists(os.path.join(UNDO_DIR, "_undo_latest.json"))


def perform_undo():
    latest = os.path.join(UNDO_DIR, "_undo_latest.json")
    if not os.path.exists(latest):
        return {"success": 0, "failed": 0, "skipped": 0}
    with open(latest, "r", encoding="utf-8") as f:
        data = json.load(f)
    stats = {"success": 0, "failed": 0, "skipped": 0, "details": []}
    for op in reversed(data.get("operations", [])):
        current = os.path.join(op["folder"], op["new"])
        target = os.path.join(op["folder"], op["old"])
        if not os.path.exists(current):
            stats["skipped"] += 1
            stats["details"].append({"file": op.get("new"), "status": "跳过: 原文件已不在"})
            continue
        try:
            os.rename(current, target)
            stats["success"] += 1
            # 撤销也要同步内存状态，否则撤销成功后前端再点预览/重命名一律「失败: 找不到」
            update_state_after_move(current, target)
        except Exception as e:
            stats["failed"] += 1
            stats["details"].append({"file": op.get("new"), "status": f"失败: {e}"})
    # 撤销记录一次性消费：成功后归档，避免「再次点撤销」拿到同一批记录显示全跳过、
    # 也避免 has_undo 永远为 True
    if stats["success"] > 0:
        try:
            done_path = os.path.join(UNDO_DIR, f"_undo_undone_{time.strftime('%Y%m%d_%H%M%S')}.json")
            shutil.move(latest, done_path)
        except Exception:
            try:
                os.remove(latest)
            except Exception:
                pass
    STATE.log(f"撤销完成！成功:{stats['success']} 失败:{stats['failed']} 跳过:{stats['skipped']}",
              "success" if stats["failed"] == 0 else "warning")
    return stats


def perform_rename(ids, scheme):
    fixed_names = (scheme or {}).get("fixed_names") or {}
    with STATE.lock:
        file_map = dict(STATE.file_map)
    # AI 命名结果直通模式：使用前端已定稿的目标名，不再按规则重新计算
    if fixed_names:
        preview = {}
        for _fid, _name in fixed_names.items():
            try:
                _fid = int(_fid)
            except (TypeError, ValueError):
                pass
            if not _name:
                continue
            # AI 直通同样做冲突防护：与自身同名放行，与其他文件撞名自动加 _1（与规则预览一致）
            info = file_map.get(_fid)
            if info:
                # 目标名可能带扩展名（Excel 导入就是靠它改扩展名），也可能只是「v1.2 报告」这种带点的名字：
                # 只有末尾是合法扩展名时才替换，否则沿用原扩展名
                old_ext = os.path.splitext(info["old"])[1]
                _base, _ext = split_target_name(_name, old_ext, bool(info.get("is_folder")))
                _base = fit_filename(info["folder"], sanitize_filename(_base) or "文件", _ext)
                _name = get_unique_path(info["folder"], _base, _ext, old_name=info["old"])
            preview[_fid] = {"new": _name}
    else:
        preview = compute_preview(ids, scheme)
    stats = {"total": len(ids), "success": 0, "failed": 0, "skipped": 0, "stopped": False}
    results = {}  # fid -> {"status": "成功"|"失败"|"跳过", "old":…, "new":…}
    operations = []
    used_targets = {}  # 执行层批内去重（覆盖 AI 直通路径，防止同批同名互相覆盖）
    for fid in ids:
        if STATE.rename_stop:
            STATE.log("重命名已被用户停止", "warning")
            stats["stopped"] = True
            break
        info = file_map.get(fid)
        new = preview.get(fid, {}).get("new", "")
        if not info:
            # 前端列表与后端状态不一致（如后端重启/重新加载后旧页面仍操作）——提示重新加载，避免静默跳过
            stats["skipped"] += 1
            results[fid] = {"status": "跳过: 文件列表已失效，请重新加载", "old": "", "new": ""}
            continue
        if not new:
            stats["skipped"] += 1
            results[fid] = {"status": "跳过: 未生成新文件名", "old": info["old"], "new": ""}
            continue
        folder = info["folder"]
        # 批内去重 + 磁盘判重：目标名已被本批其他文件占用，或磁盘上已存在同名的
        # 未参与本批文件时自动加 _N（否则预览会说可用、执行时才报「已存在」）
        used_set = used_targets.setdefault(folder, set())
        new = dedupe_candidate(folder, new, info["old"], used_set)
        if new.lower() != info["old"].lower():
            used_set.add(new.lower())
        old_path = os.path.join(folder, info["old"])
        new_path = os.path.join(folder, new)
        if new == info["old"]:
            # 新名与原名完全一致：没有可做的改名，之前会执行「自己改自己」并虚报成功，
            # 还会写进撤销记录，导致撤销结果与用户预期不符
            stats["skipped"] += 1
            results[fid] = {"status": "跳过: 名称未变化", "old": info["old"], "new": new}
            continue
        if not os.path.exists(old_path):
            stats["failed"] += 1
            results[fid] = {"status": "失败: 找不到", "old": info["old"], "new": new}
            continue
        try:
            os.rename(old_path, new_path)
            stats["success"] += 1
            operations.append({"folder": info["folder"], "old": info["old"], "new": new})
            results[fid] = {"status": "成功", "old": info["old"], "new": new}
            # 同步更新内存状态（含文件夹改名后的全部子孙条目），
            # 否则后续预览/重命名会基于旧路径「失败: 找不到」
            update_state_after_move(old_path, new_path)
            # 立即落盘撤销记录：批量改名途中关闭标签页/点退出会触发进程结束，
            # 若只在循环结束后写盘，已改好的那部分文件将永远无法撤销
            try:
                save_undo_log(operations)
            except Exception:
                pass
        except PermissionError:
            stats["failed"] += 1
            results[fid] = {"status": "失败: 权限", "old": info["old"], "new": new}
        except OSError as e:
            winerr = getattr(e, "winerror", None)
            msg = "已存在" if winerr == 183 else ("语法" if winerr == 123 else str(e))
            STATE.log(f"{info['old']} 重命名失败：{msg}", "warning")
            stats["failed"] += 1
            results[fid] = {"status": f"失败: {msg}", "old": info["old"], "new": new}
        except Exception as e:
            STATE.log(f"{info['old']} 重命名失败：{e}", "warning")
            stats["failed"] += 1
            results[fid] = {"status": f"失败: {e}", "old": info["old"], "new": new}
    if operations:
        save_undo_log(operations)
    STATE.log(f"处理完成！成功:{stats['success']} 失败:{stats['failed']} 跳过:{stats['skipped']}")
    stats["results"] = results
    return stats


# ============================================================
# AI：文件内容读取 / 文件名清理
# ============================================================
def read_document_content(filepath):
    ext = os.path.splitext(filepath)[1].lower()
    name = os.path.basename(filepath)
    try:
        if ext in TEXT_EXTS | CODE_EXTS | {".csv", ".json", ".xml", ".html", ".htm"}:
            with open(filepath, "r", encoding="utf-8", errors="ignore") as f:
                content = f.read(20000)
            if ext in (".html", ".htm"):
                content = re.sub(r"<[^>]+>", " ", content)
            return content.strip()
        if ext == ".docx":
            try:
                import docx
                return "\n".join(p.text for p in docx.Document(filepath).paragraphs).strip()
            except ImportError:
                return "python-docx 未安装"
            except Exception:
                return f"读取文件内容失败 {name}"
        if ext == ".doc":
            with open(filepath, "rb") as f:
                raw = f.read(20000)
            text = raw.decode("latin-1", errors="ignore")
            text = re.sub(r"[^\u0020-\u007E\u4e00-\u9fff\u3000-\u303f\uff00-\uffef\n]", "", text)
            return text.strip() if len(text.strip()) >= 10 else "无法读取 .doc 文件"
        if ext == ".xlsx":
            try:
                from openpyxl import load_workbook
                wb = load_workbook(filepath, read_only=True, data_only=True)
                try:
                    parts = [" | ".join(str(c) for c in row if c is not None)
                             for ws in wb.worksheets for row in ws.iter_rows(values_only=True) if any(c is not None for c in row)]
                finally:
                    wb.close()
                return "\n".join(parts).strip()
            except ImportError:
                return "openpyxl 未安装"
            except Exception:
                return f"读取文件内容失败 {name}"
        if ext == ".pptx":
            try:
                from pptx import Presentation
                prs = Presentation(filepath)
                parts = [shape.text for slide in prs.slides for shape in slide.shapes
                         if getattr(shape, "has_text_frame", False) and shape.text]
                return "\n".join(parts).strip()
            except ImportError:
                return "python-pptx 未安装"
            except Exception:
                return f"读取文件内容失败 {name}"
        if ext == ".pdf":
            try:
                import fitz
                doc = fitz.open(filepath)
                try:
                    text = "\n".join(page.get_text() for page in doc)
                finally:
                    doc.close()
                if text and text.strip():
                    return text.strip()
            except Exception:
                pass
            try:
                from pdfminer.high_level import extract_text
                t = extract_text(filepath)
                if t and t.strip():
                    return t.strip()
            except Exception:
                pass
            return "PDF 为扫描件无文字层，请安装 OCR 引擎"
        if ext == ".zip":
            import zipfile
            with zipfile.ZipFile(filepath) as zf:
                names = zf.namelist()[:200]
            return "压缩包内容清单:\n" + "\n".join(names)
        if ext in (".rar", ".7z"):
            try:
                r = subprocess.run(["7z", "l", "-ba", filepath], capture_output=True, text=True, timeout=15)
                return "压缩包内容:\n" + (r.stdout[:4000] if r.stdout else "无法列出")
            except Exception:
                return "压缩包内容无法读取"
        if ext in BOOK_EXTS:
            return "电子书文件，内容无法直接读取"
        return ""
    except Exception as e:
        return f"读取失败: {e}"


def read_folder_content(folder_path):
    try:
        names = sorted(os.listdir(folder_path))
        files = [n for n in names if os.path.isfile(os.path.join(folder_path, n))]
        return "文件夹内文件清单:\n" + "\n".join(files[:200]) if files else ""
    except Exception:
        return ""


def clean_ai_filename(raw):
    """清理 AI 返回的文件名，去除解释性文字和非法字符"""
    if not raw:
        return ""
    raw = raw.strip().strip('"').strip("'").strip("`").strip("。.")
    lines = [l.strip() for l in raw.split("\n") if l.strip()]
    # 只把「明显是解释的话」当噪声，避免把「说明文档」「生成报告」这类正常名字误杀
    explain_prefixes = ("好的", "以下是", "抱歉", "很抱歉", "当然", "ok", "OK", "---")

    def _clean(text):
        text = re.sub(r"^\d+[\.\、:：\-）\)]\s*", "", text).strip()
        text = re.sub(r'[\\/:*?"<>|\r\n\t]', "", text).strip()
        return text.strip("。.,，、").strip()[:50]

    for line in lines:
        is_explain = line.startswith(explain_prefixes) or line.endswith(("：", ":"))
        if is_explain:
            # 「好的，文件名：旅行照片」→ 取冒号后面的部分
            parts = re.split(r"[:：]", line)
            if len(parts) > 1:
                cand = _clean(parts[-1])
                if cand and not re.match(r"^\d", cand):
                    return cand
            continue
        cand = _clean(line)
        if cand:
            return cand
    # 全是解释性文字：返回空，让调用方明确报「模型未返回有效文件名」，不要拿整句当文件名
    return ""


def ensure_extension(new_name, old_name):
    ext = os.path.splitext(old_name)[1]
    if ext:
        if not new_name.lower().endswith(ext.lower()):
            return new_name + ext
        return new_name
    return new_name


# ============================================================
# AI：通用 Chat Completion 调用
# ============================================================
def chat_completion(endpoint, model, key, messages, max_tokens=512, temperature=0.7, timeout=AI_READ_TIMEOUT):
    body = {"model": model, "messages": messages, "max_tokens": max_tokens,
            "temperature": temperature, "stream": False}
    req = urllib.request.Request(endpoint, data=json.dumps(body).encode("utf-8"), method="POST")
    req.add_header("Content-Type", "application/json")
    if key:
        req.add_header("Authorization", f"Bearer {key}")
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        data = json.loads(resp.read().decode("utf-8"))
    choices = data.get("choices")
    if not choices:
        raise ValueError("响应缺少 choices 字段")
    msg = choices[0].get("message") or {}
    return (msg.get("content") or "").strip()


def ai_content_generate_name(filename, content, cfg):
    prompt = cfg.get("ai_content_prompt", DEFAULT_AI_CONFIG["ai_content_prompt"])
    # 不用 str.format：自定义提示词里若出现别的花括号（如 {文件名}）会抛 KeyError，
    # 导致 AI 命名整批失败。这里只做字面量替换。
    prompt = (str(prompt)
              .replace("{filename}", str(filename))
              .replace("{ext}", os.path.splitext(filename)[1])
              .replace("{content}", str(content[:2000])))
    if cfg.get("ai_content_mode") == "本地Ollama":
        return _ai_content_local_call(prompt, cfg)
    return _ai_content_api_call(prompt, cfg)


def _ai_content_api_call(prompt, cfg):
    # AI 内容命名页独立配置优先，留空时回退到图片识别页的云端配置（兼容旧版）
    provider = cfg.get("ai_naming_provider") or cfg.get("api_provider") or "OpenAI"
    info = PROVIDERS.get(provider, {})
    endpoint = (cfg.get("ai_naming_endpoint") or cfg.get("api_endpoint") or "").strip() or info.get("endpoint", "")
    model = (cfg.get("ai_naming_model") or cfg.get("api_model") or "").strip() or info.get("default_model", "")
    key = (cfg.get("ai_naming_key") or cfg.get("api_key") or "").strip()
    if not endpoint or not model:
        raise Exception("云端 API 未配置，请先在「AI 内容命名」页填写服务商和 API Key")
    if not key:
        raise Exception("请先填写 API Key")
    try:
        return chat_completion(endpoint, model, key, [{"role": "user", "content": prompt}])
    except Exception as e:
        raise Exception(f"API {provider} 调用失败: {e}")


def _ai_content_local_call(prompt, cfg):
    base = (cfg.get("local_api_base") or "").strip().rstrip("/") or LOCAL_DEFAULT_BASE
    model = (cfg.get("local_model") or "").strip()
    if not model:
        raise Exception("本地模型未配置，请先「检测」或填写模型名称")
    endpoint = f"{base}/v1/chat/completions"
    try:
        return chat_completion(endpoint, model, "", [{"role": "user", "content": prompt}],
                               max_tokens=cfg.get("local_max_tokens", 512),
                               temperature=cfg.get("local_temperature", 0.7),
                               timeout=AI_LOCAL_TIMEOUT)
    except Exception as e:
        raise Exception(f"本地模型调用失败（{base}）: {e}\n请确认 Ollama/LM Studio 已启动")


def ai_content_batch_call(batch, cfg):
    """批量调用模型为一批文件生成名字，返回 {index: name}（index 从 1 起）。
    解析失败/照抄原名的项不会出现在返回中。"""
    lines = []
    for i, b in enumerate(batch):
        fname = b["old"]
        content = (b.get("content") or "")[:500]
        lines.append(f"{i+1}. 原文件名: {fname}\n   内容: {content}")
    prompt = (f"请为以下{len(batch)}个文件根据其内容分别生成【新文件名】，要求：\n"
              f"1. 新文件名必须与内容相关，简洁（不超过15字），不要扩展名\n"
              f"2. 禁止照抄原文件名，必须生成不同的新名字\n"
              f"3. 严格按格式输出，每行一个，格式: 编号 → 新文件名\n"
              + "\n".join(lines) + "\n\n只返回结果，每行格式如: 1 → 项目周报")
    if cfg.get("ai_content_mode") == "本地Ollama":
        raw = _ai_content_local_call(prompt, cfg)
    else:
        raw = _ai_content_api_call(prompt, cfg)
    parsed = {}
    for line in raw.split("\n"):
        # 兼容多种编号分隔：1. / 1、 / 1: / 1： / 1- / 1) / 1． / 1 →
        m = re.match(r"^\s*(?:【?(\d+)】?|(\d+))\s*[.、:：\-)）．→]\s*(.+?)\s*$", line)
        if m:
            idx = int(m.group(1) or m.group(2))
            name = clean_ai_filename(m.group(3))
            # 过滤：照抄原文件名的无效输出
            if 1 <= idx <= len(batch) and name:
                orig = os.path.splitext(batch[idx - 1]["old"])[0]
                if name != orig:
                    parsed[idx] = name
    return parsed


def run_ai_naming(ids, cfg):
    """AI 内容命名主入口：返回 {results, err, stopped}，支持中途停止与进度"""
    with STATE.lock:
        file_map = dict(STATE.file_map)
    STATE.ai_naming_stop = False
    STATE.ai_naming = True
    STATE.ai_naming_total = len(ids)
    STATE.ai_naming_done = 0
    STATE.ai_naming_result = None
    STATE.ai_naming_error = None
    try:
        limit_map = {"auto": 2000, "500": 500, "1000": 1000, "2000": 2000, "5000": 5000}
        limit = limit_map.get(cfg.get("ai_content_limit", "auto"), 2000)
        batch_opt = cfg.get("ai_batch_opt", "batch")
        batch = []
        for fid in ids:
            info = file_map.get(fid)
            if not info:
                continue
            full = os.path.join(info["folder"], info["old"])
            if info.get("_is_folder") or os.path.isdir(full):
                content = read_folder_content(full)
            else:
                content = read_document_content(full)
            batch.append({"fid": fid, "old": info["old"], "content": (content or "")[:limit]})
        if not batch:
            return {}, "没有可处理的文件", False
        STATE.ai_naming_total = len(batch)
        results = {}
        batch_mode = batch_opt == "batch" and len(batch) > 1
        if batch_mode:
            # 自动分批：避免单次请求过大导致 Ollama/云端超时甚至崩溃
            # 前端设置项叫 batch_size；旧代码只读 ai_batch_size，导致用户的「批量大小」永远无效
            try:
                chunk_size = int(cfg.get("ai_batch_size") or cfg.get("batch_size") or 10)
            except (TypeError, ValueError):
                chunk_size = 10
            chunk_size = max(1, chunk_size)
            for start in range(0, len(batch), chunk_size):
                chunk = batch[start:start + chunk_size]
                if STATE.ai_naming_stop:
                    STATE.log("AI 内容命名已被用户停止", "warning")
                    break
                names = {}
                try:
                    names = ai_content_batch_call(chunk, cfg)
                except Exception as e:
                    # 本批整体失败（如超时）：降级为逐文件处理，不让整批丢失
                    STATE.log(f"批量命名第 {start // chunk_size + 1} 批失败，已降级逐文件处理：{str(e)[:120]}", "warning")
                    for j, b in enumerate(chunk):
                        if STATE.ai_naming_stop:
                            break
                        try:
                            raw = ai_content_generate_name(b["old"], b["content"], cfg)
                            n = clean_ai_filename(raw)
                            if n:
                                names[j + 1] = n
                        except Exception as e2:
                            pass
                for j, b in enumerate(chunk):
                    name = names.get(j + 1)
                    if name:
                        results[b["fid"]] = ensure_extension(name, b["old"])
                    else:
                        # 解析失败不再静默：标记错误供前端展示
                        results[b["fid"]] = f"__ERR__:模型未返回有效文件名（第 {start + j + 1} 个）"
                STATE.ai_naming_done = min(start + len(chunk), len(batch))
        else:
            for b in batch:
                if STATE.ai_naming_stop:
                    STATE.log("AI 内容命名已被用户停止", "warning")
                    break
                try:
                    raw = ai_content_generate_name(b["old"], b["content"], cfg)
                    new_name = clean_ai_filename(raw)
                    if new_name:
                        results[b["fid"]] = ensure_extension(new_name, b["old"])
                    else:
                        results[b["fid"]] = f"__ERR__:模型未返回有效文件名"
                except Exception as e:
                    results[b["fid"]] = f"__ERR__:{e}"
                STATE.ai_naming_done += 1
        STATE.ai_naming_result = results
        return results, None, STATE.ai_naming_stop
    finally:
        STATE.ai_naming = False


# ============================================================
# AI：图片内容识别
# ============================================================
def build_recognition_prompt(cfg):
    categories = [c.strip() for c in cfg.get("custom_categories", "").split(",") if c.strip()] or CATEGORIES
    mode = cfg.get("classify_mode", "固定分类")
    if mode == "AI自由识别":
        style = cfg.get("free_style", "图片描述")
        if style == "图片标题":
            return AI_TITLE_PROMPT
        return AI_SELF_DESCRIBE_PROMPT
    return AI_CATEGORIZE_FOR_TAGS.format(categories="、".join(categories))


def parse_recognition_content(content, cfg):
    categories = [c.strip() for c in cfg.get("custom_categories", "").split(",") if c.strip()] or CATEGORIES
    mode = cfg.get("classify_mode", "固定分类")
    max_tags = int(cfg.get("max_tag_count", 5))
    if mode == "AI自由识别":
        style = cfg.get("free_style", "图片描述")
        if style == "图片标题":
            return clean_ai_filename(content) or "其他", []
        # 自定义分类：格式 "分类-标签1-标签2"
        parts = [p.strip() for p in content.split(cfg.get("tag_separator", "-")) if p.strip()]
        if len(parts) >= 2:
            return sanitize_filename(parts[0])[:30] or "其他", parts[1:][:max_tags]
        return clean_ai_filename(content) or "其他", []
    # 固定分类 / 自定义分类：匹配最相近类别
    best_cat, best_score = None, 0
    cleaned = re.sub(r"[，.\"'\s]", "", content)
    for c in categories:
        score = cleaned.count(c)
        if score > best_score:
            best_score, best_cat = score, c
    return sanitize_filename(best_cat or "其他"), []


def rule_recognize(ids):
    with STATE.lock:
        file_map = dict(STATE.file_map)
    results = {}
    for fid in ids:
        info = file_map.get(fid)
        if not info:
            continue
        results[fid] = {"category": guess_category(info["old"])}
    return results


def _image_b64(path):
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode("utf-8")


def ollama_recognize(ids, cfg):
    base = (cfg.get("ollama_base") or "").strip().rstrip("/") or LOCAL_DEFAULT_BASE
    model = (cfg.get("ollama_model") or "").strip()
    if not model:
        raise Exception("Ollama 模型未配置，请先「检测并接入已有模型」")
    # 前置校验：模型必须支持图片识别，避免 35 张全部失败
    ok, msg = check_ollama_vision(model, base)
    if not ok:
        raise Exception(msg)
    with STATE.lock:
        file_map = dict(STATE.file_map)
    prompt = build_recognition_prompt(cfg)
    results = {}
    logged_errors = set()
    for fid in ids:
        if STATE.recognize_stop:
            STATE.log("Ollama 识别已被用户停止", "warning")
            break
        info = file_map.get(fid)
        if not info:
            continue
        path = os.path.join(info["folder"], info["old"])
        ext = IMAGE_EXT_MAP.get(os.path.splitext(info["old"])[1].lower(), "jpeg")
        try:
            b64 = _image_b64(path)
            body = {"model": model, "prompt": prompt, "images": [b64], "stream": False}
            req = urllib.request.Request(f"{base}/api/generate", data=json.dumps(body).encode("utf-8"),
                                         headers={"Content-Type": "application/json"}, method="POST")
            with urllib.request.urlopen(req, timeout=AI_READ_TIMEOUT) as resp:
                data = json.loads(resp.read().decode("utf-8"))
            content = (data.get("response") or "").strip()
            cat, tags = parse_recognition_content(content, cfg)
            results[fid] = {"category": cat, "tags": tags}
        except Exception as e:
            msg = str(e)
            # 只记录前 3 种不同的错误，避免刷屏
            key = msg[:120]
            if key not in logged_errors and len(logged_errors) < 3:
                logged_errors.add(key)
                STATE.log(f"Ollama 识别失败（{info['old']}）：{msg[:150]}", "warning")
            results[fid] = {"category": "其他", "tags": [], "error": msg}
    return results


def cloud_api_recognize(ids, cfg):
    provider = cfg.get("api_provider") or "OpenAI"
    info = PROVIDERS.get(provider, {})
    endpoint = (cfg.get("api_endpoint") or "").strip() or info.get("endpoint", "")
    model = (cfg.get("api_model") or "").strip() or info.get("default_model", "")
    key = (cfg.get("api_key") or "").strip()
    if not endpoint or not model:
        raise Exception("云端 API 未配置，请先填写服务商和 API Key")
    if not key:
        raise Exception("请先填写 API Key")
    with STATE.lock:
        file_map = dict(STATE.file_map)
    prompt = build_recognition_prompt(cfg)
    results = {}
    logged_errors = set()
    for fid in ids:
        if STATE.recognize_stop:
            STATE.log("云端 API 识别已被用户停止", "warning")
            break
        info = file_map.get(fid)
        if not info:
            continue
        path = os.path.join(info["folder"], info["old"])
        ext = IMAGE_EXT_MAP.get(os.path.splitext(info["old"])[1].lower(), "jpeg")
        try:
            b64 = _image_b64(path)
            image_url = f"data:image/{ext};base64,{b64}"
            messages = [{"role": "user", "content": [
                {"type": "text", "text": prompt},
                {"type": "image_url", "image_url": {"url": image_url}}]}]
            content = chat_completion(endpoint, model, key, messages, max_tokens=256)
            cat, tags = parse_recognition_content(content, cfg)
            results[fid] = {"category": cat, "tags": tags}
        except Exception as e:
            msg = str(e)
            key = msg[:120]
            if key not in logged_errors and len(logged_errors) < 3:
                logged_errors.add(key)
                STATE.log(f"云端 API 识别失败（{info['old']}）：{msg[:150]}", "warning")
            results[fid] = {"category": "其他", "tags": [], "error": msg}
    return results


def model_pack_recognize(ids):
    """使用本地 model_pack 子进程（run_clip.py）进行 CLIP 识别"""
    pack_dir = os.path.join(BASE_DIR, "resources", "model_pack")
    python_exe = os.path.join(pack_dir, "python", "Scripts", "python.exe")
    run_clip = os.path.join(pack_dir, "run_clip.py")
    if not os.path.exists(python_exe) or not os.path.exists(run_clip):
        python_exe = os.path.join(pack_dir, "python", "python.exe")
    if not os.path.exists(python_exe) or not os.path.exists(run_clip):
        raise Exception("未找到模型包（resources/model_pack），无法使用本地 CLIP 识别")
    with STATE.lock:
        file_map = dict(STATE.file_map)
    valid_ids = [fid for fid in ids if fid in file_map]
    image_paths = [os.path.join(file_map[fid]["folder"], file_map[fid]["old"]) for fid in valid_ids]
    if not image_paths:
        raise Exception("没有可识别的图片")
    try:
        # 用 Popen 启动子进程，支持中途停止
        proc = subprocess.Popen([python_exe, run_clip] + image_paths,
                                stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                text=True, cwd=pack_dir)
        try:
            out, err = proc.communicate(timeout=300)
        except subprocess.TimeoutExpired:
            proc.kill()
            proc.wait()
            raise Exception("模型包识别超时（300 秒）")
    except Exception as e:
        raise Exception(f"模型包执行失败: {e}")
    if proc.returncode != 0:
        raise Exception(f"模型包执行失败: {(err or '')[:300]}")
    try:
        raw_results = json.loads(out.strip())
    except json.JSONDecodeError:
        raise Exception(f"模型包输出解析失败: {out[:200]}")
    results = {}
    if STATE.recognize_stop:
        STATE.log("本地 CLIP 识别已被用户停止", "warning")
        return results
    for fid, r in zip(valid_ids, raw_results):
        results[fid] = {"category": r.get("category", "其他"),
                        "similarity": r.get("similarity", 0.0),
                        "tags": []}
    return results


def local_clip_recognize(ids):
    """直接导入 open_clip 进行识别（需要本机 Python 环境已装依赖）"""
    try:
        import torch
        import open_clip
        from PIL import Image
    except ImportError:
        return None  # 无依赖，返回 None 触发降级
    with STATE.lock:
        file_map = dict(STATE.file_map)
    device = "cuda" if torch.cuda.is_available() else "cpu"
    try:
        model_name = "ViT-B-32"
        # 不传 pretrained 会得到随机权重，分类结果毫无意义
        model, _, preprocess = open_clip.create_model_and_transforms(model_name, pretrained="openai")
        model = model.to(device)
        model.eval()
        tokenizer = open_clip.get_tokenizer(model_name)
        prompts = CLIP_PROMPTS_LOCAL
        text_tokens = tokenizer(prompts).to(device)
        with torch.no_grad():
            text_features = model.encode_text(text_tokens)
            text_features /= text_features.norm(dim=-1, keepdim=True)
    except Exception:
        return None
    results = {}
    for fid in ids:
        if STATE.recognize_stop:
            STATE.log("本地 CLIP 识别已被用户停止", "warning")
            break
        info = file_map.get(fid)
        if not info:
            continue
        path = os.path.join(info["folder"], info["old"])
        try:
            image = Image.open(path).convert("RGB")
            img_input = preprocess(image).unsqueeze(0).to(device)
            with torch.no_grad():
                img_features = model.encode_image(img_input)
                img_features /= img_features.norm(dim=-1, keepdim=True)
                sim = (img_features @ text_features.T).squeeze(0)
                best = sim.argmax().item()
            results[fid] = {"category": CATEGORIES[best], "similarity": float(sim[best].item()), "tags": []}
        except Exception as e:
            results[fid] = {"category": "其他", "tags": [], "error": str(e)}
    return results


CLIP_PROMPTS_LOCAL = [
    "a photo of a person, people, human face, portrait",
    "a photo of an animal, pet, wildlife, bird, cat, dog",
    "a photo of landscape, nature, scenery, mountain, beach, forest, sunset, outdoor",
    "a photo of a building, architecture, house, city, urban, street",
    "a photo of food, meal, dish, cuisine, drink",
    "a photo of a document, paper, text, book, page, screenshot, whiteboard",
    "a photo of something else, miscellaneous, other",
]


def recognize_images(ids, cfg):
    """识别分发：rule / clip / ollama / api，返回 {results, mode, note}"""
    STATE.recognize_stop = False
    STATE.recognizing = True
    try:
        mode = cfg.get("recognition_mode", "rule")
        if mode == "rule":
            return rule_recognize(ids), "rule", None
        if mode == "ollama":
            results = ollama_recognize(ids, cfg)
            return results, "ollama", None
        if mode == "api":
            results = cloud_api_recognize(ids, cfg)
            return results, "api", None
        if mode == "clip":
            # 优先本机 open_clip，其次 model_pack 子进程
            try:
                r = local_clip_recognize(ids)
                if r is not None:
                    return r, "clip", None
            except Exception:
                pass
            try:
                r = model_pack_recognize(ids)
                return r, "clip(model_pack)", None
            except Exception as e:
                # 最后降级到规则分类
                results = rule_recognize(ids)
                return results, "rule", f"CLIP 不可用，已降级为规则分类：{e}"
        # 未知模式降级
        return rule_recognize(ids), "rule", None
    finally:
        STATE.recognizing = False


# ============================================================
# 模板 / 方案管理
# ============================================================
def list_templates():
    out = []
    for fn in sorted(os.listdir(TEMPLATE_DIR)):
        if fn.endswith(".json"):
            name = fn[:-5]
            try:
                with open(os.path.join(TEMPLATE_DIR, fn), "r", encoding="utf-8") as f:
                    data = json.load(f)
                out.append({"name": name, "data": data})
            except Exception:
                out.append({"name": name, "data": {}})
    return out


def save_template(name, data):
    if not re.match(r"^[\w\u4e00-\u9fff\- ]+$", name):
        return False, "模板名称只能包含中文、字母、数字、下划线和横线"
    try:
        with open(os.path.join(TEMPLATE_DIR, f"{name}.json"), "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        return True, "已保存"
    except Exception as e:
        return False, str(e)


def _safe_json_path(base_dir, name):
    """把用户给的名字限制成 base_dir 下的单个 .json 文件，拒绝路径穿越/绝对路径。

    之前 delete_template("..\\..\\x") 能删到目录外任意 .json（save_* 有白名单，delete_* 没有）。
    """
    if not isinstance(name, str) or not name:
        return None
    if not re.match(r"^[\w\u4e00-\u9fff\- ]+$", name):
        return None
    p = os.path.join(base_dir, f"{name}.json")
    try:
        if os.path.dirname(os.path.abspath(p)) != os.path.abspath(base_dir):
            return None
    except Exception:
        return None
    return p


def delete_template(name):
    p = _safe_json_path(TEMPLATE_DIR, name)
    if p and os.path.exists(p):
        os.remove(p)
        return True
    return False


def list_schemes():
    out = []
    for fn in sorted(os.listdir(SCHEME_DIR)):
        if fn.endswith(".json"):
            name = fn[:-5]
            try:
                with open(os.path.join(SCHEME_DIR, fn), "r", encoding="utf-8") as f:
                    data = json.load(f)
                out.append({"name": name, "data": data})
            except Exception:
                out.append({"name": name, "data": {}})
    return out


def save_scheme(name, data):
    if not re.match(r"^[\w\u4e00-\u9fff\- ]+$", name):
        return False, "方案名称只能包含中文、字母、数字、下划线和横线"
    try:
        with open(os.path.join(SCHEME_DIR, f"{name}.json"), "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        return True, "已保存"
    except Exception as e:
        return False, str(e)


def delete_scheme(name):
    p = _safe_json_path(SCHEME_DIR, name)
    if p and os.path.exists(p):
        os.remove(p)
        return True
    return False


# ============================================================
# Ollama 检测
# ============================================================
def _model_has_vision(name, caps):
    """判断模型是否支持图片识别：优先信任 Ollama 官方 capabilities 标注；
    仅当 capabilities 为空（旧版 Ollama 未上报）时才按名称启发式兜底。"""
    if caps:
        return "vision" in caps
    return any(v in name.lower() for v in ("vl", "vision", "minicpm"))


def detect_ollama(base=None):
    """检测 Ollama 模型，返回 (models, vision_models, error)。
    vision_models 为具备视觉（多模态）能力的模型名列表，仅这些模型可用于图片识别。"""
    base = (base or LOCAL_DEFAULT_BASE).strip().rstrip("/")
    models, vision_models = [], []
    try:
        req = urllib.request.Request(f"{base}/api/tags", method="GET")
        with urllib.request.urlopen(req, timeout=5) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        for m in data.get("models", []):
            name = m.get("name", "")
            if not name:
                continue
            models.append(name)
            caps = m.get("capabilities") or []
            if _model_has_vision(name, caps):
                vision_models.append(name)
    except Exception:
        return [], [], f"无法连接 Ollama（{base}）"
    return models, vision_models, None


def check_ollama_vision(model, base=None):
    """确认指定 Ollama 模型支持图片识别（vision）。返回 (ok, message)。
    ok=False 时 message 为可读的中文错误提示。"""
    base = (base or LOCAL_DEFAULT_BASE).strip().rstrip("/")
    try:
        req = urllib.request.Request(f"{base}/api/tags", method="GET")
        with urllib.request.urlopen(req, timeout=5) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        found, has_vision = False, False
        for m in data.get("models", []):
            if m.get("name", "") == model:
                found = True
                has_vision = _model_has_vision(model, m.get("capabilities") or [])
                break
        if not found:
            return False, f"Ollama 中不存在模型「{model}」，请点击「检测并接入已有模型」重新选择"
        if not has_vision:
            return False, (f"模型「{model}」不支持图片识别（无 vision 能力）。\n"
                           f"请选择视觉模型，如 qwen2.5vl、llava 等（检测时会优先列出支持图片的模型）")
        return True, ""
    except Exception as e:
        return False, f"无法连接 Ollama（{base}）：{e}"


# ============================================================
# Excel 导出
# ============================================================
def export_excel(filename, rows):
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment
    except ImportError:
        return None, "缺少 openpyxl，请运行 pip install openpyxl"
    wb = Workbook()
    ws = wb.active
    ws.title = "重命名清单"
    headers = ["原文件名", "新文件名", "状态", "分类", "所在文件夹"]
    fill = PatternFill("solid", fgColor="14B8A6")
    for ci, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=ci, value=h)
        cell.font = Font(bold=True, color="04211C")
        cell.fill = fill
        cell.alignment = Alignment(horizontal="center")
    for ri, row in enumerate(rows, 2):
        for ci, k in enumerate(("old", "new", "status", "cat", "folder"), 1):
            ws.cell(row=ri, column=ci, value=row.get(k, ""))
    for ci in range(1, 6):
        ws.column_dimensions[chr(64 + ci)].width = 30
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue(), None


def import_excel_bytes(data):
    """解析上传的 Excel 文件（二进制），提取「原文件名 → 新文件名」映射"""
    try:
        from openpyxl import load_workbook
    except ImportError:
        return None, "缺少 openpyxl，请运行 pip install openpyxl"
    try:
        import io as _io
        wb = load_workbook(_io.BytesIO(data), read_only=True, data_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
    except Exception as e:
        return None, f"Excel 解析失败：{e}"
    if not rows:
        return [], "Excel 为空"
    mappings = []
    seen = set()
    for row in rows:
        # 跳过表头（含"原文件名"等字样）和空行
        if not row or all(c is None or str(c).strip() == "" for c in row):
            continue
        old = str(row[0] or "").strip() if row[0] is not None else ""
        new = str(row[1] or "").strip() if len(row) > 1 and row[1] is not None else ""
        if not old:
            continue
        # 表头检测
        if old in ("原文件名", "文件名", "old", "old_name", "源文件名") and len(rows) > 1:
            continue
        if new in ("新文件名", "new", "new_name", "目标文件名"):
            continue
        key = (old, new)
        if key in seen:
            continue
        seen.add(key)
        mappings.append({"old": old, "new": new})
    return mappings, None


def export_recognition_results(results, fmt="json"):
    """导出识别结果：json / csv / txt，成功返回 (data, None)"""
    try:
        if fmt == "csv":
            import csv as _csv
            buf = io.StringIO()
            w = _csv.writer(buf)
            w.writerow(["原文件名", "分类", "标签", "相似度"])
            for r in results:
                w.writerow([r["old"], r["category"], ",".join(r.get("tags", [])), r.get("similarity", "")])
            return buf.getvalue().encode("utf-8-sig"), None
        if fmt == "txt":
            lines = [f"{r['old']}\t{r['category']}\t{','.join(r.get('tags', []))}" for r in results]
            return ("\n".join(lines) + "\n").encode("utf-8"), None
        # json
        return json.dumps(results, ensure_ascii=False, indent=2).encode("utf-8"), None
    except Exception as e:
        return None, f"导出失败：{e}"


# ============================================================
# HTTP 服务器
# ============================================================
def json_resp(handler, data, status=200):
    body = json.dumps(data, ensure_ascii=False).encode("utf-8")
    handler.send_response(status)
    handler.send_header("Content-Type", "application/json; charset=utf-8")
    handler.send_header("Content-Length", str(len(body)))
    handler.send_header("Cache-Control", "no-store")
    handler.end_headers()
    handler.wfile.write(body)


def read_json(handler):
    try:
        length = int(handler.headers.get("Content-Length", 0) or 0)
    except (TypeError, ValueError):
        length = 0
    if length <= 0:
        return {}
    try:
        raw = handler.rfile.read(length)
    except Exception:
        return {}
    try:
        data = json.loads(raw.decode("utf-8"))
    except Exception:
        return {}
    return data if isinstance(data, dict) else {}


class Handler(BaseHTTPRequestHandler):
    server_version = "BatchRenamerWeb/2.0"

    def log_message(self, fmt, *args):
        pass

    # ---------- 请求来源校验 ----------
    def _allowed_hosts(self):
        port = self.server.server_address[1]
        return {f"127.0.0.1:{port}", f"localhost:{port}", f"[::1]:{port}"}

    def _check_request_origin(self):
        """拦掉跨站请求。

        服务只监听 127.0.0.1，但浏览器里任意网页都能向它发「简单请求」：
        不加校验时，用户开着本工具去访问别的网页，对方就能批量改名/撤销/退出，
        甚至通过 /api/open/file 启动任意程序。
        - 有 Origin/Referer 时必须是本机同源；
        - POST 必须是 application/json（会触发预检，普通表单/文本请求发不出来）。
        """
        origin = self.headers.get("Origin") or ""
        if origin and origin != "null":
            host = urlparse(origin).netloc
            if host not in self._allowed_hosts():
                return False
        ref = self.headers.get("Referer") or ""
        if ref:
            try:
                if urlparse(ref).netloc not in self._allowed_hosts():
                    return False
            except Exception:
                return False
        return True

    def _known_path(self, p):
        """路径是否属于当前已加载的文件/文件夹（限制 open/reveal 的作用范围）"""
        try:
            target = os.path.normcase(os.path.normpath(str(p)))
        except Exception:
            return False
        with STATE.lock:
            entries = list(STATE.file_map.values())
        for e in entries:
            op = e.get("old_path")
            if op and os.path.normcase(os.path.normpath(op)) == target:
                return True
        return False

    def _send_file(self, path, content_type):
        try:
            with open(path, "rb") as f:
                body = f.read()
        except OSError:
            json_resp(self, {"error": "文件不存在"}, 404)
            return
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-cache")
        self.end_headers()
        self.wfile.write(body)

    def _qr_path(self, filename):
        """定位打赏二维码：兼容源码结构 resources/qr 与打包结构 qr（PyInstaller）"""
        for base in (BASE_DIR, RES_DIR):
            for p in (os.path.join(base, "resources", "qr", filename),
                      os.path.join(base, "qr", filename)):
                if os.path.exists(p):
                    return p
        return os.path.join(BASE_DIR, "resources", "qr", filename)

    # ---------- GET ----------
    def do_GET(self):
        parsed = urlparse(self.path)
        path = unquote(parsed.path)
        try:
            if path in ("/", "/index.html", "/webui.html"):
                self._send_file(os.path.join(BASE_DIR, "webui.html"), "text/html; charset=utf-8")
            elif path == "/qr/wechat.png":
                self._send_file(self._qr_path("wechat.png"), "image/png")
            elif path == "/qr/alipay.png":
                self._send_file(self._qr_path("alipay.png"), "image/png")
            elif path == "/api/health":
                json_resp(self, {"status": "ok", "folders": len(STATE.folders), "files": len(STATE.files)})
            elif path == "/api/folders":
                json_resp(self, {"folders": list(STATE.folders)})
            elif path == "/api/files":
                json_resp(self, {"files": list(STATE.files)})
            elif path == "/api/templates":
                json_resp(self, {"templates": list_templates()})
            elif path == "/api/schemes":
                json_resp(self, {"schemes": list_schemes()})
            elif path == "/api/logs":
                json_resp(self, {"logs": STATE.get_logs()})
            elif path == "/api/logs/clear":
                # 写操作只允许 POST（GET 可被 <img>/链接等跨站触发）
                json_resp(self, {"error": "请使用 POST"}, 405)
            elif path == "/api/undo":
                json_resp(self, {"has_undo": has_undo()})
            elif path == "/api/providers":
                json_resp(self, {"providers": PROVIDERS})
            elif path == "/api/categories":
                json_resp(self, {"categories": CATEGORIES})
            elif path == "/api/ai/config":
                json_resp(self, {"config": load_ai_config()})
            else:
                json_resp(self, {"error": f"未知接口 {path}"}, 404)
        except Exception as e:
            traceback.print_exc()
            json_resp(self, {"error": str(e)}, 500)

    # ---------- POST ----------
    def do_POST(self):
        parsed = urlparse(self.path)
        path = unquote(parsed.path)
        if not self._check_request_origin():
            json_resp(self, {"error": "跨站请求被拒绝"}, 403)
            return
        ctype = (self.headers.get("Content-Type") or "").lower()
        if not ctype.startswith("application/json"):
            # 浏览器对 application/json 会先发预检，跨站网页因此发不出能生效的 POST
            json_resp(self, {"error": "请求需要 Content-Type: application/json"}, 415)
            return
        data = read_json(self)
        # 除退出相关接口外，任意活跃请求视为程序仍在被使用，取消待退出
        if path not in ("/api/exit", "/api/exit/cancel"):
            _cancel_pending_exit()
        try:
            if path == "/api/exit":
                mode = (data or {}).get("mode", "manual")
                delay = _EXIT_DELAY_MANUAL if mode == "manual" else _EXIT_DELAY_AUTO
                _schedule_exit(delay, "按钮退出" if mode == "manual" else "浏览器标签页关闭")
                json_resp(self, {"ok": True, "delay": delay})
            elif path == "/api/exit/cancel":
                _cancel_pending_exit()
                json_resp(self, {"ok": True})
            elif path == "/api/folders/pick":
                self._handle_pick_folder()
            elif path == "/api/folders/add":
                folder = normalize_folder(data.get("path", ""))
                if folder and os.path.isdir(folder) and not any(_same_folder(f, folder) for f in STATE.folders):
                    STATE.folders.append(folder)
                    STATE.log(f"已添加文件夹: {folder}")
                    json_resp(self, {"folders": list(STATE.folders)})
                elif folder and any(_same_folder(f, folder) for f in STATE.folders):
                    json_resp(self, {"error": "该文件夹已在列表中", "folders": list(STATE.folders)}, 400)
                else:
                    json_resp(self, {"error": "无效的文件夹路径"}, 400)
            elif path == "/api/folders/remove":
                folder = normalize_folder(data.get("path", ""))
                hit = next((f for f in STATE.folders if _same_folder(f, folder)), None)
                if hit is not None:
                    STATE.folders.remove(hit)
                    json_resp(self, {"folders": list(STATE.folders)})
                else:
                    json_resp(self, {"error": "文件夹不存在"}, 404)
            elif path == "/api/folders/clear":
                n = len(STATE.folders)
                STATE.folders = []
                STATE.log(f"已清空文件夹列表（{n} 个）", "warning")
                json_resp(self, {"folders": []})
            elif path == "/api/files/load":
                self._handle_load_files(data)
            elif path == "/api/preview":
                self._handle_preview(data)
            elif path == "/api/rename":
                self._handle_rename(data)
            elif path == "/api/rename/stop":
                STATE.rename_stop = True
                STATE.log("已请求停止重命名", "warning")
                json_resp(self, {"ok": True, "renaming": STATE.renaming})
            elif path == "/api/rename/status":
                json_resp(self, {"renaming": STATE.renaming, "stop_requested": STATE.rename_stop})
            elif path == "/api/excel/import":
                import base64 as _b64
                try:
                    raw = _b64.b64decode(data.get("file", "") or "")
                except Exception:
                    raw = b""
                if not raw:
                    json_resp(self, {"error": "文件内容为空或解码失败"}, 400)
                    return
                mappings, err = import_excel_bytes(raw)
                if err:
                    json_resp(self, {"error": err}, 400)
                    return
                STATE.log(f"从 Excel 导入 {len(mappings)} 条重命名映射", "success")
                json_resp(self, {"ok": True, "mappings": mappings})
            elif path == "/api/recognize/export":
                fmt = data.get("format", "json")
                entries = []
                for f in STATE.files:
                    entries.append({"old": f.get("old", ""), "category": f.get("cat", "其他"),
                                    "tags": f.get("tags", []), "similarity": f.get("similarity", "")})
                payload, ferr = export_recognition_results(entries, fmt)
                if ferr:
                    json_resp(self, {"error": ferr}, 400)
                    return
                STATE.log(f"导出识别结果（{fmt}，{len(entries)} 条）", "success")
                self.send_response(200)
                cts = {"json": "application/json; charset=utf-8", "csv": "text/csv; charset=utf-8",
                       "txt": "text/plain; charset=utf-8"}.get(fmt, "application/octet-stream")
                self.send_header("Content-Type", cts)
                self.send_header("Content-Disposition", f'attachment; filename="recognition.{fmt}"; filename*=UTF-8\'\'%E8%AF%86%E5%88%AB%E7%BB%93%E6%9E%9C.{fmt}')
                self.send_header("Content-Length", str(len(payload)))
                self.end_headers()
                self.wfile.write(payload)
            elif path == "/api/open/file":
                p = data.get("path", "")
                if p and os.path.exists(p) and not self._known_path(p):
                    json_resp(self, {"error": "只能打开当前列表中的文件"}, 403)
                elif p and os.path.exists(p):
                    try:
                        os.startfile(p)  # noqa
                        json_resp(self, {"ok": True})
                    except Exception as e:
                        json_resp(self, {"error": f"无法打开：{e}"}, 500)
                else:
                    json_resp(self, {"error": "文件不存在"}, 404)
            elif path == "/api/open/reveal":
                p = data.get("path", "")
                if p and os.path.exists(p) and not self._known_path(p):
                    json_resp(self, {"error": "只能定位当前列表中的文件"}, 403)
                elif p and os.path.exists(p):
                    try:
                        subprocess.Popen(["explorer", "/select,", p])
                        json_resp(self, {"ok": True})
                    except Exception as e:
                        json_resp(self, {"error": f"无法定位：{e}"}, 500)
                else:
                    json_resp(self, {"error": "文件不存在"}, 404)
            elif path == "/api/logs/clear":
                STATE.clear_logs()
                json_resp(self, {"ok": True, "logs": []})
            elif path == "/api/undo":
                if STATE.renaming:
                    json_resp(self, {"error": "正在重命名，请等待完成后再撤销"}, 409)
                    return
                json_resp(self, perform_undo())
            elif path == "/api/templates/save":
                ok, msg = save_template(data.get("name", ""), data.get("data", {}))
                STATE.log(f"保存模板「{data.get('name','')}」：{msg}", "success" if ok else "error")
                json_resp(self, {"ok": ok, "message": msg, "templates": list_templates()},
                          200 if ok else 400)
            elif path == "/api/templates/delete":
                delete_template(data.get("name", ""))
                STATE.log(f"删除模板「{data.get('name','')}」", "warning")
                json_resp(self, {"ok": True, "templates": list_templates()})
            elif path == "/api/schemes/save":
                ok, msg = save_scheme(data.get("name", ""), data.get("data", {}))
                STATE.log(f"保存方案「{data.get('name','')}」：{msg}", "success" if ok else "error")
                json_resp(self, {"ok": ok, "message": msg, "schemes": list_schemes()},
                          200 if ok else 400)
            elif path == "/api/schemes/delete":
                delete_scheme(data.get("name", ""))
                STATE.log(f"删除方案「{data.get('name','')}」", "warning")
                json_resp(self, {"ok": True, "schemes": list_schemes()})
            elif path == "/api/export/excel":
                self._handle_export_excel(data)
            elif path == "/api/recognize":
                self._handle_recognize(data)
            elif path == "/api/recognize/stop":
                STATE.recognize_stop = True
                STATE.log("已请求停止识别", "warning")
                json_resp(self, {"ok": True, "recognizing": STATE.recognizing})
            elif path == "/api/recognize/status":
                json_resp(self, {"recognizing": STATE.recognizing, "stop_requested": STATE.recognize_stop})
            elif path == "/api/ai/config":
                incoming = data.get("config")
                if not isinstance(incoming, dict) or not incoming:
                    # 请求体畸形时若直接保存会把用户已存的 API Key/服务商整体清空
                    json_resp(self, {"error": "配置内容为空或格式错误，未保存"}, 400)
                    return
                cfg = load_ai_config()
                cfg.update(incoming)
                save_ai_config(cfg)
                STATE.log("AI 配置已保存", "success")
                json_resp(self, {"ok": True, "config": load_ai_config()})
            elif path == "/api/ai/detect-ollama":
                models, vision_models, err = detect_ollama(data.get("base"))
                json_resp(self, {"models": models, "vision_models": vision_models, "error": err})
            elif path == "/api/ai/naming":
                self._handle_ai_naming(data)
            elif path == "/api/ai/naming/stop":
                STATE.ai_naming_stop = True
                STATE.log("已请求停止 AI 内容命名", "warning")
                json_resp(self, {"ok": True, "ai_naming": STATE.ai_naming})
            elif path == "/api/ai/naming/status":
                json_resp(self, {"running": STATE.ai_naming, "stop_requested": STATE.ai_naming_stop,
                                 "total": STATE.ai_naming_total, "done": STATE.ai_naming_done,
                                 "results": STATE.ai_naming_result,
                                 "error": STATE.ai_naming_error})
            elif path == "/api/ai/test":
                self._handle_test_api(data)
            else:
                json_resp(self, {"error": f"未知接口 {path}"}, 404)
        except Exception as e:
            traceback.print_exc()
            json_resp(self, {"error": str(e)}, 500)

    # ---------- 具体处理 ----------
    def _handle_pick_folder(self):
        try:
            import tkinter as tk
            from tkinter import filedialog
            root = tk.Tk()
            root.withdraw()
            root.attributes("-topmost", True)
            folder = filedialog.askdirectory(title="选择目标文件夹")
            root.destroy()
            folder = normalize_folder(folder)
            if folder and os.path.isdir(folder):
                if not any(_same_folder(f, folder) for f in STATE.folders):
                    STATE.folders.append(folder)
                    STATE.log(f"已添加文件夹: {folder}")
            json_resp(self, {"folders": list(STATE.folders)})
        except Exception as e:
            json_resp(self, {"error": f"无法打开目录选择框: {e}"}, 500)

    def _handle_load_files(self, data):
        folders = data.get("folders") or STATE.folders
        if not folders:
            json_resp(self, {"error": "请先添加文件夹"}, 400)
            return
        cleaned = []
        for f in folders:
            nf = normalize_folder(f)
            if nf and os.path.isdir(nf) and not any(_same_folder(x, nf) for x in cleaned):
                cleaned.append(nf)
        if not cleaned:
            json_resp(self, {"error": "文件夹不存在或不可访问"}, 400)
            return
        STATE.folders = cleaned
        files = scan_files(STATE.folders, data.get("recursive", True),
                           data.get("only_folders", False), data.get("filter_system", True))
        STATE.log(f"已加载 {len(files)} 个文件/文件夹")
        json_resp(self, {"files": files})

    def _handle_preview(self, data):
        ids = data.get("ids", [])
        scheme = data.get("scheme", {})
        if not ids:
            json_resp(self, {"error": "没有要预览的文件"}, 400)
            return
        results = compute_preview(ids, scheme)
        prefix = (scheme or {}).get("prefix", "") or ""
        STATE.log(f"生成预览：{len(ids)} 个文件，前缀「{prefix}」，模式 { (scheme or {}).get('mode','') }", "info")
        json_resp(self, {"results": results})

    def _handle_rename(self, data):
        ids = data.get("ids", [])
        scheme = data.get("scheme", {})
        if not ids:
            json_resp(self, {"error": "没有要重命名的文件"}, 400)
            return
        with STATE.lock:
            busy = STATE.renaming
            if not busy:
                STATE.renaming = True
        if busy:
            # 两个标签页/连点同时提交时，两批改名会互相踩：同一文件一成功一「找不到」，
            # 且撤销日志只有最后一批
            json_resp(self, {"error": "已有重命名任务正在进行，请等待完成后再操作"}, 409)
            return
        STATE.log(f"开始重命名：{len(ids)} 个文件，前缀「{(scheme or {}).get('prefix','')}」", "info")
        STATE.rename_stop = False
        try:
            stats = perform_rename(ids, scheme)
        finally:
            STATE.renaming = False
        json_resp(self, stats)

    def _handle_recognize(self, data):
        ids = data.get("ids", [])
        if not ids:
            json_resp(self, {"error": "没有要识别的文件"}, 400)
            return
        cfg = load_ai_config()
        # 允许请求覆盖配置（前端临时切换模式）
        if data.get("config"):
            cfg.update(data.get("config"))
        try:
            results, mode, note = recognize_images(ids, cfg)
            # 回写识别结果到 STATE.files，供「导出识别结果」使用
            with STATE.lock:
                for f in STATE.files:
                    r = results.get(f["id"])
                    if r:
                        f["cat"] = r.get("category", f.get("cat"))
                        f["tags"] = r.get("tags", [])
                        if r.get("similarity") is not None:
                            f["similarity"] = r.get("similarity")
            # 收集识别失败信息，向前端透出原因
            errors=[r.get("error") for r in results.values() if r.get("error")]
            if errors and not note:
                note=f"有 {len(errors)} 个文件识别失败，示例：{errors[0][:100]}"
            stats = {}
            for r in results.values():
                cat = r.get("category", "其他")
                stats[cat] = stats.get(cat, 0) + 1
            STATE.log(f"识别完成：{len(results)} 个文件（模式: {mode}）")
            json_resp(self, {"results": results, "stats": stats, "mode": mode, "note": note})
        except Exception as e:
            traceback.print_exc()
            json_resp(self, {"error": str(e)}, 500)

    def _json_safe(self, obj, code=200):
        """从工作线程向客户端写回 JSON（线程安全）"""
        try:
            json_resp(self, obj, code)
        except Exception:
            pass

    def _handle_ai_naming(self, data):
        ids = data.get("ids", [])
        if not ids:
            json_resp(self, {"error": "没有要命名的文件"}, 400)
            return
        if STATE.ai_naming:
            json_resp(self, {"error": "已有 AI 命名任务正在进行，请等待完成或先点击停止"}, 400)
            return
        cfg = load_ai_config()
        if data.get("config"):
            cfg.update(data.get("config"))

        # 先同步占位，避免前端第一次轮询时任务线程还没启动、被误判为「已完成（0 个）」
        STATE.ai_naming = True
        STATE.ai_naming_stop = False
        STATE.ai_naming_total = len(ids)
        STATE.ai_naming_done = 0
        STATE.ai_naming_result = None
        STATE.ai_naming_error = None

        # 独立线程执行命名：主线程立即返回，避免阻塞 HTTP 导致「停止」请求无法到达。
        # 结果一律通过 /api/ai/naming/status 轮询获取；这里不能再往已经结束的
        # 响应 socket 写第二份 JSON（主线程已响应过，多写会污染 keep-alive 连接）。
        def worker():
            try:
                results, err, stopped = run_ai_naming(ids, cfg)
                if err:
                    STATE.ai_naming_error = err
                    STATE.log(f"AI 内容命名未执行：{err}", "error")
                    return
                ok = sum(1 for v in results.values() if v and not str(v).startswith("__ERR__"))
                errs = [str(v).split(":", 1)[1] if ":" in str(v) else str(v)
                        for v in results.values() if str(v).startswith("__ERR__")]
                STATE.log(f"AI 内容命名完成：{ok} 个成功" + (f"，{len(errs)} 个失败" if errs else "") +
                          ("（已停止）" if stopped else ""))
            except Exception as e:
                traceback.print_exc()
                STATE.ai_naming_error = f"AI 内容命名异常：{e}"
                STATE.log(STATE.ai_naming_error, "error")

        threading.Thread(target=worker, daemon=True).start()
        json_resp(self, {"started": True, "total": len(ids)})

    def _handle_test_api(self, data):
        """测试云端 API 连接（支持前端实时表单配置覆盖）。
        scope=ai_naming 时测试「AI 内容命名」页独立配置，否则测试「图片识别」页配置。"""
        cfg = load_ai_config()
        if data.get("config"):
            cfg.update(data.get("config"))
        if data.get("scope") == "ai_naming":
            provider = cfg.get("ai_naming_provider") or cfg.get("api_provider") or "OpenAI"
            info = PROVIDERS.get(provider, {})
            endpoint = (cfg.get("ai_naming_endpoint") or cfg.get("api_endpoint") or "").strip() or info.get("endpoint", "")
            model = (cfg.get("ai_naming_model") or cfg.get("api_model") or "").strip()
            key = (cfg.get("ai_naming_key") or cfg.get("api_key") or "").strip()
        else:
            provider = cfg.get("api_provider") or "OpenAI"
            info = PROVIDERS.get(provider, {})
            endpoint = (cfg.get("api_endpoint") or "").strip() or info.get("endpoint", "")
            model = (cfg.get("api_model") or "").strip()
            key = (cfg.get("api_key") or "").strip()
        if not endpoint or not model:
            json_resp(self, {"ok": False, "message": "请先配置服务商和模型"})
            return
        if not key:
            json_resp(self, {"ok": False, "message": "请先填写 API Key"})
            return
        try:
            resp = chat_completion(endpoint, model, key,
                                   [{"role": "user", "content": "hi"}], max_tokens=1, timeout=15)
            json_resp(self, {"ok": True, "message": f"连接成功：{resp[:80]}"})
        except Exception as e:
            json_resp(self, {"ok": False, "message": f"连接失败：{e}"})

    def _handle_export_excel(self, data):
        filename = data.get("filename", "重命名清单.xlsx")
        if not filename.endswith(".xlsx"):
            filename += ".xlsx"
        payload, err = export_excel(filename, data.get("rows", []))
        if err:
            json_resp(self, {"error": err}, 400)
            return
        self.send_response(200)
        self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
        self.send_header("Content-Length", str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)


def _existing_instance(port):
    """探测该端口上是否已经有本程序在运行（避免重复启动后两个实例抢端口）"""
    try:
        req = urllib.request.Request(f"http://127.0.0.1:{port}/api/health", method="GET")
        with urllib.request.urlopen(req, timeout=1.5) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        return isinstance(data, dict) and data.get("status") == "ok"
    except Exception:
        return False


def main():
    port = 8890
    if len(sys.argv) > 1:
        try:
            port = int(sys.argv[1])
        except ValueError:
            print(f"端口参数无效：{sys.argv[1]}，已回退到 8890")
            port = 8890
    # HTTPServer 默认 allow_reuse_address=True，同一端口可以被两个实例同时绑定，
    # 请求会随机落到其中一个实例上（表现为「预览/重命名时有时无」「列表突然空了」）
    if _existing_instance(port):
        print(f"检测到程序已在 http://localhost:{port} 运行，直接打开浏览器即可（不再重复启动）")
        _auto_open_browser(port, delay=0.2)
        time.sleep(1.5)
        return
    try:
        server = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    except OSError as e:
        print(f"端口 {port} 无法启动：{e}")
        print(f"请换一个端口重试，例如：python server.py {port + 1}")
        return
    global _current_server
    _current_server = server
    STATE.log(f"Web 服务已启动，端口 {port}，数据目录 {DATA_DIR}", "success")
    print("=" * 56)
    print("  批量文件重命名工具 · Web 控制台（完整版）")
    print(f"  访问地址: http://localhost:{port}")
    print(f"  数据目录: {DATA_DIR}")
    print("  按 Ctrl+C 停止服务")
    print("=" * 56)
    # 启动成功后自动打开浏览器（若本机有 GUI 环境）
    _auto_open_browser(port)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n服务已停止")
        server.shutdown()


def _auto_open_browser(port, delay=1.2):
    """服务启动后延迟自动打开默认浏览器（后台线程，失败静默）"""
    def _open():
        import time as _t
        _t.sleep(delay)
        try:
            import webbrowser
            webbrowser.open(f"http://localhost:{port}")
        except Exception:
            pass
    try:
        import threading
        threading.Thread(target=_open, daemon=True).start()
    except Exception:
        pass


if __name__ == "__main__":
    main()
