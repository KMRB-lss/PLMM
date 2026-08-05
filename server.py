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
    cleaned = cleaned.strip().rstrip(". ")
    cleaned = re.sub("_+", "_", cleaned).strip("_")
    reserved = {"con", "prn", "aux", "nul", *(f"com{i}" for i in range(1, 10)),
                *(f"lpt{i}" for i in range(1, 10))}
    if os.path.splitext(cleaned)[0].lower() in reserved:
        cleaned = "_" + cleaned
    if len(cleaned) > 255:
        cleaned = cleaned[:255]
    return cleaned or "文件"


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


def read_file_metadata(filepath):
    meta = {}
    try:
        st = os.stat(filepath)
        meta["创建日期"] = datetime.datetime.fromtimestamp(st.st_ctime).strftime("%Y-%m-%d")
        meta["修改日期"] = datetime.datetime.fromtimestamp(st.st_mtime).strftime("%Y-%m-%d")
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
        meta["原文件名"] = os.path.splitext(os.path.basename(filepath))[0]
        try:
            author, owner = _win_file_author(filepath)
            if author:
                meta["作者"] = author
            if owner:
                meta["所有者"] = owner
        except Exception:
            pass
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
        except re.error:
            pass
    return text


def build_new_name(scheme, index, meta=None):
    mode = scheme.get("mode", "prefix_num")
    num = int(scheme.get("start_num", 1)) + index
    padded = str(num).zfill(int(scheme.get("pad_width", 2)))
    prefix = replace_meta_placeholders(str(scheme.get("prefix", "文件") or "文件"), meta, index=index)
    suffix = replace_meta_placeholders(str(scheme.get("suffix", "") or ""), meta, index=index)
    prefix = prefix.replace("{编号}", padded)
    suffix = suffix.replace("{编号}", padded)
    prefix = sanitize_filename(prefix)
    suffix = sanitize_filename(suffix)
    if not prefix:
        prefix = "文件"
    if mode == "num_prefix":
        return f"{padded}_{prefix}"
    if mode == "prefix_num_suffix":
        return f"{prefix}_{padded}_{suffix}" if suffix else f"{prefix}_{padded}"
    return f"{prefix}_{padded}"


def get_unique_path(folder, new_base, ext):
    if not os.path.exists(os.path.join(folder, f"{new_base}{ext}")):
        return f"{new_base}{ext}"
    counter = 1
    while True:
        cand = f"{new_base}_{counter}{ext}"
        if not os.path.exists(os.path.join(folder, cand)):
            return cand
        counter += 1


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
def scan_files(folders, recursive, only_folders, filter_system):
    entries = []
    id_counter = [0]

    def _make_entry(path, name, depth, is_folder):
        fid = id_counter[0]
        id_counter[0] += 1
        try:
            st = os.stat(path)
            size, mtime = st.st_size, datetime.datetime.fromtimestamp(st.st_mtime).strftime("%Y-%m-%d %H:%M")
        except Exception:
            size, mtime = 0, ""
        ext = os.path.splitext(name)[1].lower()
        with STATE.lock:
            STATE.file_map[fid] = {"folder": os.path.dirname(path), "old": name, "old_path": path}
        return {"id": fid, "old": name, "ext": ext, "relpath": name,
                "size": size, "mtime": mtime, "cat": guess_category(name) if not is_folder else "文件夹",
                "folder": os.path.dirname(path), "_is_folder": is_folder}

    def walk(folder, depth=0):
        try:
            names = sorted(os.listdir(folder))
        except Exception:
            return
        for name in names:
            path = os.path.join(folder, name)
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
    for fid in ids:
        info = file_map.get(fid)
        if not info:
            results[fid] = {"new": "", "status": "未知文件"}
            continue
        try:
            folder, old = info["folder"], info["old"]
            name, ext = os.path.splitext(old)

            # 命名规则是否显式配置（前端 getScheme 传入 rules_enabled）
            rules_on = bool(scheme.get("rules_enabled"))
            old_name = name

            # 分类+编号模式：基于文件的 category 字段
            if scheme.get("use_category") and scheme.get("categories"):
                cat = scheme["categories"].get(str(fid)) or "其他"
                cat = sanitize_filename(cat)
                key = cat
                scheme_indices.setdefault(key, 0)
                num = int(scheme.get("start_num", 1)) + scheme_indices[key]
                padded = str(num).zfill(int(scheme.get("pad_width", 2)))
                new_base = f"{cat}_{padded}"
                scheme_indices[key] += 1
            # 按扩展名分组编号：同一扩展名使用连续编号
            elif scheme.get("group_by_ext"):
                key = ext.lower() or "(无扩展名)"
                scheme_indices.setdefault(key, 0)
                meta = read_file_metadata(os.path.join(folder, old))
                new_base = build_new_name(scheme, scheme_indices[key], meta=meta)
                scheme_indices[key] += 1
            # 命名规则已配置：前缀/编号生成新名
            elif rules_on:
                scheme_indices.setdefault("default", 0)
                meta = read_file_metadata(os.path.join(folder, old))
                new_base = build_new_name(scheme, scheme_indices["default"], meta=meta)
                scheme_indices["default"] += 1
            # 命名规则未配置：以原文件名为基础（此时「添加&替换/正则」独立生效）
            else:
                new_base = old_name

            # 添加内容（前/后）：作用于基础名
            add_text = (scheme.get("add_text") or "").strip()
            if add_text:
                new_base = f"{add_text}_{new_base}" if scheme.get("add_position") == "前" else f"{new_base}_{add_text}"
            # 查找替换：作用于基础名
            find_text = (scheme.get("find_text") or "").strip()
            if find_text and find_text in new_base:
                new_base = new_base.replace(find_text, scheme.get("replace_text", ""))

            new_base = apply_regex(new_base, scheme)
            new_base = sanitize_filename(new_base) or "文件"
            results[fid] = {"new": get_unique_path(folder, new_base, ext), "status": "待重命名"}
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
    stats = {"success": 0, "failed": 0, "skipped": 0}
    for op in reversed(data.get("operations", [])):
        current = os.path.join(op["folder"], op["new"])
        target = os.path.join(op["folder"], op["old"])
        if not os.path.exists(current):
            stats["skipped"] += 1
            continue
        try:
            os.rename(current, target)
            stats["success"] += 1
        except Exception:
            stats["failed"] += 1
    STATE.log(f"撤销完成！成功:{stats['success']} 失败:{stats['failed']} 跳过:{stats['skipped']}",
              "success" if stats["failed"] == 0 else "warning")
    return stats


def perform_rename(ids, scheme):
    preview = compute_preview(ids, scheme)
    stats = {"total": len(ids), "success": 0, "failed": 0, "skipped": 0, "stopped": False}
    results = {}  # fid -> {"status": "成功"|"失败"|"跳过", "old":…, "new":…}
    operations = []
    with STATE.lock:
        file_map = dict(STATE.file_map)
    for fid in ids:
        if STATE.rename_stop:
            STATE.log("重命名已被用户停止", "warning")
            stats["stopped"] = True
            break
        info = file_map.get(fid)
        new = preview.get(fid, {}).get("new", "")
        if not info or not new:
            stats["skipped"] += 1
            results[fid] = {"status": "跳过", "old": "", "new": ""}
            continue
        old_path = os.path.join(info["folder"], info["old"])
        new_path = os.path.join(info["folder"], new)
        if not os.path.exists(old_path):
            stats["failed"] += 1
            results[fid] = {"status": "失败: 找不到", "old": info["old"], "new": new}
            continue
        try:
            os.rename(old_path, new_path)
            stats["success"] += 1
            operations.append({"folder": info["folder"], "old": info["old"], "new": new})
            results[fid] = {"status": "成功", "old": info["old"], "new": new}
            # 同步更新内存状态：使后续预览/重命名基于真实磁盘文件名，避免「找不到文件」
            with STATE.lock:
                entry = STATE.file_map.get(fid)
                if entry is not None:
                    entry["old"] = new
                    entry["old_path"] = new_path
                for sf in STATE.files:
                    if sf.get("id") == fid:
                        sf["old"] = new
                        sf["relpath"] = new
                        sf["ext"] = os.path.splitext(new)[1].lower()
                        if not sf.get("_is_folder"):
                            sf["cat"] = guess_category(new)
                        break
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
                parts = [" | ".join(str(c) for c in row if c is not None)
                         for ws in wb.worksheets for row in ws.iter_rows(values_only=True) if any(c is not None for c in row)]
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
                text = "\n".join(page.get_text() for page in doc)
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
    skip_prefixes = ("请求", "注意", "要求", "说明", "这里", "生成", "以下", "根据", "好的", "ok", "OK", "---")
    for l in lines:
        if l.startswith(skip_prefixes) or re.match(r"^\d+[\.\、:：]\s*", l):
            l = re.sub(r"^\d+[\.\、:：]\s*", "", l).strip()
        if l:
            cleaned = re.sub(r'[\\/:*?"<>|\r\n]', "", l).strip().strip("。")
            return cleaned[:50]
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
    prompt = prompt.format(filename=filename, ext=os.path.splitext(filename)[1], content=content[:2000])
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
                               temperature=cfg.get("local_temperature", 0.7))
    except Exception as e:
        raise Exception(f"本地模型调用失败（{base}）: {e}\n请确认 Ollama/LM Studio 已启动")


def ai_content_batch_call(batch, cfg):
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
        m = re.match(r"^\s*(\d+)\s*[.、:：→\-]\s*(.+?)\s*$", line)
        if m:
            idx = int(m.group(1))
            name = clean_ai_filename(m.group(2))
            # 过滤：照抄原文件名的无效输出
            if 1 <= idx <= len(batch) and name:
                orig = os.path.splitext(batch[idx - 1]["old"])[0]
                if name != orig:
                    parsed[idx] = name
    return [parsed.get(i + 1, "") for i in range(len(batch))]


def run_ai_naming(ids, cfg):
    """AI 内容命名主入口：返回 {results, err, stopped}，支持中途停止"""
    with STATE.lock:
        file_map = dict(STATE.file_map)
    STATE.ai_naming_stop = False
    STATE.ai_naming = True
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
        results = {}
        if batch_opt == "batch" and len(batch) > 1:
            if STATE.ai_naming_stop:
                STATE.log("AI 内容命名已被用户停止", "warning")
                return results, None, True
            names = ai_content_batch_call(batch, cfg)
            for b, new_name in zip(batch, names):
                if new_name:
                    results[b["fid"]] = ensure_extension(new_name, b["old"])
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
                except Exception as e:
                    results[b["fid"]] = f"__ERR__:{e}"
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
    image_paths = [os.path.join(file_map[fid]["folder"], file_map[fid]["old"]) for fid in ids if fid in file_map]
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
    for fid, r in zip(ids, raw_results):
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
        model, _, preprocess = open_clip.create_model_and_transforms(model_name)
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


def delete_template(name):
    p = os.path.join(TEMPLATE_DIR, f"{name}.json")
    if os.path.exists(p):
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
    p = os.path.join(SCHEME_DIR, f"{name}.json")
    if os.path.exists(p):
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
    length = int(handler.headers.get("Content-Length", 0))
    if length <= 0:
        return {}
    raw = handler.rfile.read(length)
    try:
        return json.loads(raw.decode("utf-8"))
    except Exception:
        return {}


class Handler(BaseHTTPRequestHandler):
    server_version = "BatchRenamerWeb/2.0"

    def log_message(self, fmt, *args):
        pass

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
                STATE.clear_logs()
                json_resp(self, {"ok": True, "logs": []})
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
                folder = data.get("path", "").strip()
                if folder and os.path.isdir(folder) and folder not in STATE.folders:
                    STATE.folders.append(folder)
                    STATE.log(f"已添加文件夹: {folder}")
                    json_resp(self, {"folders": list(STATE.folders)})
                else:
                    json_resp(self, {"error": "无效的文件夹路径或已存在"}, 400)
            elif path == "/api/folders/remove":
                folder = data.get("path", "")
                if folder in STATE.folders:
                    STATE.folders.remove(folder)
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
                if p and os.path.exists(p):
                    try:
                        os.startfile(p)  # noqa
                        json_resp(self, {"ok": True})
                    except Exception as e:
                        json_resp(self, {"error": f"无法打开：{e}"}, 500)
                else:
                    json_resp(self, {"error": "文件不存在"}, 404)
            elif path == "/api/open/reveal":
                p = data.get("path", "")
                if p and os.path.exists(p):
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
                save_ai_config(data.get("config", {}))
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
            if folder and os.path.isdir(folder):
                if folder not in STATE.folders:
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
        STATE.folders = [f for f in folders if os.path.isdir(f)]
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
        STATE.log(f"开始重命名：{len(ids)} 个文件，前缀「{(scheme or {}).get('prefix','')}」", "info")
        STATE.rename_stop = False
        STATE.renaming = True
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

    def _handle_ai_naming(self, data):
        ids = data.get("ids", [])
        if not ids:
            json_resp(self, {"error": "没有要命名的文件"}, 400)
            return
        cfg = load_ai_config()
        if data.get("config"):
            cfg.update(data.get("config"))
        try:
            results, err, stopped = run_ai_naming(ids, cfg)
            if err:
                json_resp(self, {"error": err}, 400)
                return
            ok = sum(1 for v in results.values() if v and not str(v).startswith("__ERR__"))
            errs = [str(v).split(":", 1)[1] if ":" in str(v) else str(v)
                    for v in results.values() if str(v).startswith("__ERR__")]
            STATE.log(f"AI 内容命名完成：{ok} 个成功" + (f"，{len(errs)} 个失败" if errs else "") +
                      ("（已停止）" if stopped else ""))
            json_resp(self, {"results": results, "stopped": stopped,
                             "failed": len(errs), "error_detail": (errs[0] if errs else None)})
        except Exception as e:
            traceback.print_exc()
            json_resp(self, {"error": str(e)}, 500)

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


def main():
    port = 8890
    if len(sys.argv) > 1:
        port = int(sys.argv[1])
    server = ThreadingHTTPServer(("127.0.0.1", port), Handler)
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
