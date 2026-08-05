<div align="center">

# 📁 批量文件重命名工具

**Batch Renamer** — 一个纯本地运行、基于浏览器的批量文件重命名工具

![License](https://img.shields.io/badge/License-GPL--3.0-blue.svg)
![Python](https://img.shields.io/badge/Python-3.9%2B-green.svg)
![Platform](https://img.shields.io/badge/Platform-Windows-0078D6.svg)
![Dependencies](https://img.shields.io/badge/dependencies-核心%2B可选-lightgrey.svg)
![CI](https://github.com/KMRB-lss/PLMM/actions/workflows/ci.yml/badge.svg)
![Release](https://img.shields.io/github/v/release/KMRB-lss/PLMM?color=orange)

无需安装 GUI 框架 · 单文件后端 · 浏览器即用 · 免安装 exe 可选

</div>

---

## ✨ 功能特性

| 类别 | 功能 |
|------|------|
| 📁 **文件夹管理** | Windows 目录选择框 / 手动输入路径 / 多文件夹 / 递归子目录 / 仅文件夹 / 过滤系统文件 |
| ✏️ **命名规则** | 前缀+编号 / 编号+前缀 / 前缀+编号+后缀 / 分类+编号 / 按扩展名编号，实时预览；**编号位数可自定义 1-20**（1 表示不补零）；**「命名规则」「编号功能」独立开关**，可只启用其中一部分 |
| 🔧 **文本处理** | 添加内容（前/后）、查找替换、正则替换（场景预设，`$1` 捕获引用在真实重命名中同样生效）、元数据占位符点击插入；**「添加&替换」独立开关**，与命名规则可组合、可单独使用 |
| 🗂 **模板管理** | 保存 / 加载 / 删除常用规则模板 |
| 📊 **Excel** | 导入「原文件名→新文件名」映射、导出重命名清单 |
| 👁 **重命名** | 生成预览 → 开始重命名（真实磁盘操作，支持中途停止，可一键撤销） |
| 🖼 **图片识别** | 规则分类 / 本地 CLIP / Ollama 多模态 / 云端 AI API（7 家服务商），支持多标签、自动导出、按分类筛选；Ollama 检测自动优先列出视觉模型（选中非视觉模型会明确报错）；**识别结果可直接作为新文件名**，也可按分类生成编号预览（两者独立） |
| 🤖 **AI 内容命名** | 读取 txt / md / pdf / docx / xlsx / pptx / zip 等文件内容，云端 API 或本地 Ollama 生成文件名；**支持文件夹**（按其内文件清单生成文件夹名），批量自动分批防超时、单批失败自动降级逐文件、实时进度显示、**停止生成即时响应**；**AI 命名结果直通重命名**，不会被命名规则覆盖 |
| 🖱 **预览列表** | **表头冻结**（滚动文件列表时列头始终可见）、**列宽可拖拽调整**、**悬停显示完整名称**、4 种搜索模式、7 字段排序、右键菜单（打开/定位/复制） |
| 🔗 **标签页联动** | 「生成预览」按钮随当前标签页变化：AI 内容命名页点击即触发 AI 生成，其余页为规则预览 |
| 💖 **更多** | 打赏收款码、操作日志、内置使用说明 / 正则参考 / 云端配置指南、一键退出程序（关闭标签页自动退出） |

## 🚀 快速开始

### 环境要求

**exe 方式**：仅需 **Windows 7+**，无需安装 Python。

**源码方式**：需要 **Windows 7+** + **Python 3.9+**（[下载地址](https://www.python.org/downloads/)）

### 方式零（推荐）：免安装 exe

从 [Releases 页面](https://github.com/KMRB-lss/PLMM/releases) 下载 `BatchFileRenamer_v1.0.0.exe`（约 55MB，已内置全部文档解析库，不含本地 CLIP），双击即可运行，**启动后自动打开浏览器界面**。
需要本地 CLIP 图片识别时，另下载 `model_pack.zip` 解压到 exe 同级目录 `resources/model_pack/` 即可。

> 运行数据（撤销记录、规则模板、AI 配置等）会自动创建在 **exe 同目录** 的 `data/` 文件夹中，请保持 exe 所在文件夹可写。
> 使用完毕可点击界面右上角 **「⏻ 退出程序」** 关闭；直接关闭浏览器标签页也会在约 10 秒后自动退出后台程序。

### 安装依赖

```bash
# 核心依赖（Excel 导入导出、图片处理、文档内容提取）
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

**只使用部分功能**：命名规则设置与添加&替换各有一个「启用」勾选框，编号功能有独立开关，默认全部勾选。取消勾选后，预览与重命名只应用勾选的功能——例如只想做添加&替换，取消「命名规则」即可，规则不会叠加。

## 📦 项目结构

```
批量文件重命名工具/
│   server.py               # Web 后端（HTTP 服务 + 全部业务逻辑，单文件）
│   webui.html              # Web 前端界面（单文件，无构建步骤）
│   启动Web版.bat           # Windows 一键启动脚本
│   requirements.txt        # Python 核心依赖（含文档解析库）
│   requirements-clip.txt   # 可选：本地 CLIP 识别依赖（torch 体积大）
│   LICENSE                 # GPL-3.0 许可证
│   LICENSING.md            # 第三方依赖许可证清单
│   CONTRIBUTING.md         # 贡献指南
│
├── .github/                # GitHub 模板与 CI
│   ├── workflows/ci.yml    # CI：语法 + 依赖检查
│   ├── ISSUE_TEMPLATE/     # Issue 模板（Bug / 功能建议）
│   └── PULL_REQUEST_TEMPLATE.md
│
├── resources/
│   ├── qr/                 # 打赏收款码
│   └── model_pack/         # 本地 CLIP 模型包（可选，不入库，Release 下载）
│
├── docs/                   # 文档（如 Ollama 使用指南）
└── data/                   # 运行时数据（自动生成，不入库）
    ├── ai_config.json      # AI 配置（含 API Key，请勿外泄）
    ├── _undo_logs/         # 撤销记录
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

### AI 内容命名

AI 内容命名会读取文件正文（txt / md / pdf / docx / xlsx / pptx / zip 等）交给模型生成文件名，支持文件夹（按其内文件清单生成文件夹名）。

- **两种模式**：云端 API（需 Key）与本地 Ollama（免费离线）
- **批量优化**：单文件独立生成（更贴合每个文件）或批量统一风格（更一致）
- **停止即时**：批量生成过程中可随时点「停止生成」，立即生效并保留已完成结果
- **结果直通**：AI 页点「生成预览」即触发 AI 命名，点「开始重命名」按 AI 名字执行，不会被命名规则覆盖

## 🛠 开发

```bash
# 安装开发依赖
pip install -r requirements.txt

# 语法检查
python -m py_compile server.py

# 启动开发服务
python server.py 8890
```

- 前端为单文件 `webui.html`，无构建步骤，改完刷新即生效
- 提交规范见 [CONTRIBUTING.md](CONTRIBUTING.md)
- 参与贡献请先阅读 [Issue 模板](.github/ISSUE_TEMPLATE/) 与 [PR 模板](.github/PULL_REQUEST_TEMPLATE.md)

### 构建免安装 exe（可选）

```bash
# 建议在干净 venv 中打包（避免混入无关大依赖）
pip install -r requirements.txt pyinstaller

pyinstaller --noconfirm --onefile --noconsole --name BatchFileRenamer \
  --add-data "webui.html;." --add-data "resources/qr;qr" server.py
```

> 产物约 55MB（已内置 python-docx / python-pptx / pymupdf / pdfminer.six 文档解析库，不含本地 CLIP）。若本机 site-packages 装有 torch 等大包，请先在干净 venv 中构建，否则 exe 会异常膨胀。

## 📄 许可证

本项目采用 **GNU General Public License v3.0（GPL-3.0）**，详见 [LICENSE](LICENSE)。
第三方依赖许可证清单见 [LICENSING.md](LICENSING.md)。

> **GPL-3.0 是 copyleft（传染性）许可证**：任何基于本项目的修改或衍生作品，在分发时必须同样以 GPL-3.0 开源。

---

<div align="center">
如果这个工具对你有帮助，欢迎 <b>⭐ Star</b> 支持 💖
</div>
