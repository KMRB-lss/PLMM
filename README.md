<div align="center">

# 📁 批量文件重命名工具

**Batch Renamer** — 一个纯本地运行、基于浏览器的批量文件重命名工具

![License](https://img.shields.io/badge/License-GPL--3.0-blue.svg)
![Python](https://img.shields.io/badge/Python-3.8%2B-green.svg)
![Platform](https://img.shields.io/badge/Platform-Windows-0078D6.svg)
![Dependencies](https://img.shields.io/badge/dependencies-标准库%2B可选-lightgrey.svg)

无需安装 GUI 框架 · 单文件后端 · 浏览器即用

</div>

---

## ✨ 功能特性

| 类别 | 功能 |
|------|------|
| 📁 **文件夹管理** | Windows 目录选择框 / 手动输入路径 / 多文件夹 / 递归子目录 / 仅文件夹 / 过滤系统文件 |
| ✏️ **命名规则** | 前缀+编号 / 编号+前缀 / 前缀+编号+后缀 / 分类+编号 / 按扩展名编号，实时预览 |
| 🔧 **文本处理** | 添加内容（前/后）、查找替换、正则替换（场景预设 + 参考文档）、元数据占位符点击插入 |
| 🗂 **方案与模板** | 保存 / 加载 / 删除 / 应用到勾选 / 全部设为该方案 |
| 📊 **Excel** | 导入「原文件名→新文件名」映射、导出重命名清单 |
| 👁 **重命名** | 生成预览 → 开始重命名（真实磁盘操作，支持中途停止，可一键撤销） |
| 🖼 **图片识别** | 规则分类 / 本地 CLIP / Ollama 多模态 / 云端 AI API（7 家服务商），支持多标签、自动导出、按分类筛选 |
| 🤖 **AI 内容命名** | 读取 txt / md / pdf / docx / xlsx / pptx / zip 等文件内容，云端 API 或本地 Ollama 生成文件名 |
| 🖱 **预览列表** | 路径列、右键菜单（打开/定位/复制）、4 种搜索模式、7 字段排序、勾选文件夹 |
| 💖 **更多** | 打赏收款码、操作日志、内置使用说明 / 正则参考 / 云端配置指南 |

## 🚀 快速开始

### 环境要求
- **Windows 7+**（工具调用 Windows 原生功能）
- **Python 3.8+**（[下载地址](https://www.python.org/downloads/)）

### 安装依赖

```bash
# 核心依赖（Excel 导入导出、图片处理）
pip install -r requirements.txt

# 可选：本地 CLIP 图片识别（torch 体积较大，约 2GB+）
pip install -r requirements-clip.txt
```

> 不安装 CLIP 依赖时，识别模式会自动降级为规则分类，其余功能不受影响。
> Ollama 与云端 API 识别无需额外安装。

### 启动

**方式一（推荐）**：双击 **`启动Web版.bat`** —— 服务启动后**自动打开浏览器**界面

**方式二**：命令行启动
```bash
python server.py 8890
```

**方式三**：指定端口
```bash
python server.py 9000
```

### 本地 CLIP 模型包

本地 CLIP 图片识别需要模型包 `resources/model_pack/`（约 750MB，体积大不入 Git）。
请从项目的 **Release 页面** 下载 `model_pack.zip`，解压到项目根目录，得到 `resources/model_pack/` 即可。
不使用本地 CLIP 时无需下载，识别模式自动降级为规则分类。

### 使用

浏览器打开 **http://localhost:8890** 即可。

四步完成重命名：
1. 📁 选择文件夹
2. ⟳ 加载文件列表
3. 👁 生成预览（核对新文件名）
4. ▶ 开始重命名（出错可 ↩ 撤销）

## 📦 项目结构

```
批量文件重命名工具/
│   server.py               # Web 后端（HTTP 服务 + 全部业务逻辑，单文件）
│   webui.html              # Web 前端界面（单文件，无构建步骤）
│   启动Web版.bat           # Windows 一键启动脚本
│   requirements.txt        # Python 依赖
│   LICENSE                 # GPL-3.0 许可证
│   LICENSING.md            # 第三方依赖许可证清单
│
├── resources/
│   ├── qr/                 # 打赏收款码
│   └── model_pack/         # 本地 CLIP 模型包（可选，不入库）
│
├── docs/                   # 文档（如 Ollama 使用指南）
└── data/                   # 运行时数据（自动生成，不入库）
    ├── ai_config.json      # AI 配置（含 API Key，请勿外泄）
    ├── _undo_logs/         # 撤销记录
    ├── schemes/            # 命名方案
    ├── templates/          # 规则模板
    └── logs/               # 操作日志
```

## 🧩 图片识别与 AI 命名

### 四种识别模式

| 模式 | 需要 | 联网 | 说明 |
|------|------|------|------|
| ⚙️ 规则分类 | 无 | 否 | 按扩展名自动分类，即刻可用 |
| 🖥 本地 CLIP | torch + open_clip | 否 | 语义识别，需 `resources/model_pack` 模型包 |
| 🦙 Ollama | 本地 Ollama | 否 | 自动检测视觉模型（minicpm-o、qwen2.5vl 等） |
| ☁️ 云端 API | API Key | 是 | OpenAI / DeepSeek / 智谱 GLM / 通义千问 / 硅基流动 / 腾讯混元 / 自定义 |

> 各服务商注册与 API Key 获取步骤见界面内置「📖 配置指南」。
> 图片识别详细教程见 [docs/Ollama使用指南.md](docs/Ollama使用指南.md)。

## 🛠 开发

```bash
# 语法检查
python -m py_compile server.py

# 启动开发服务
python server.py 8890
```

- 前端为单文件 `webui.html`，无构建步骤，改完刷新即生效
- 提交规范见 [CONTRIBUTING.md](CONTRIBUTING.md)
- 参与贡献请先阅读 [Issue 模板](.github/ISSUE_TEMPLATE/) 与 [PR 模板](.github/PULL_REQUEST_TEMPLATE.md)

## 📄 许可证

本项目采用 **GNU General Public License v3.0（GPL-3.0）**，详见 [LICENSE](LICENSE)。
第三方依赖许可证清单见 [LICENSING.md](LICENSING.md)。

> **GPL-3.0 是 copyleft（传染性）许可证**：任何基于本项目的修改或衍生作品，在分发时必须同样以 GPL-3.0 开源。

---

<div align="center">
如果这个工具对你有帮助，欢迎 <b>⭐ Star</b> 支持 💖
</div>
