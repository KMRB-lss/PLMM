# 第三方库许可证清单

## 本项目许可证

项目主体代码（`server.py`、`webui.html` 等）为原创代码，采用 **GNU General Public License v3.0（GPL-3.0）**（见 `LICENSE` 文件）。

GPL-3.0 是 copyleft（传染性）许可证：允许商用、分发、修改，但**任何衍生作品在分发时必须同样以 GPL-3.0 开源**。若想将本项目代码整合进闭源/专有项目，需要获得作者额外的商业授权。

## 主要依赖（requirements.txt 核心）

| 库名 | 许可证 | 说明 |
|------|--------|------|
| Pillow | MIT-CMU | 图片处理 |
| openpyxl | MIT | Excel 导入/导出 |
| tqdm | MPL-2.0 / MIT 双许可 | 进度条 |
| python-docx | MIT | 读取 .docx 文档内容（AI 内容命名） |
| python-pptx | MIT | 读取 .pptx 演示文稿内容（AI 内容命名） |
| pymupdf | **AGPL-3.0**（或商业许可） | 读取 .pdf 文本（优先尝试，失败自动回退 pdfminer.six） |
| pdfminer.six | MIT | 读取 .pdf 文本（兜底） |

> **关于 PyMuPDF（AGPL-3.0）**：PyMuPDF 采用 AGPL-3.0 或商业双许可。AGPL 与 GPL-3.0 同为 copyleft，但含网络交互条款。本项目中 PyMuPDF 为可选增强组件（代码在 `try/except` 中导入，缺失时自动回退 pdfminer.six），仅用于本地读取 PDF 文本；本地使用不受影响。若需将本项目对外分发为服务，可在构建时排除 PyMuPDF（`pip install pdfminer.six` 即可，不装 pymupdf），PDF 文本提取仍可用。

## 可选依赖（requirements-clip.txt，本地 CLIP）

| 库名 | 许可证 | 说明 |
|------|--------|------|
| torch | BSD-3-Clause | 本地 CLIP 识别（可选） |
| open_clip_torch | MIT | 本地 CLIP 模型（可选） |

## model_pack 嵌入式依赖（可选，不入库）

以下库随 `resources/model_pack` 捆绑分发（Git 已忽略该目录，各自许可证文件位于 `site-packages/<包名>.dist-info/`）：

| 库名 | 许可证 |
|------|--------|
| numpy | BSD-3-Clause（含 0BSD/MIT/Zlib/CC0 子组件） |
| sympy | BSD |
| networkx | BSD-3-Clause |
| huggingface_hub | Apache 2.0 |
| torch | BSD-3-Clause |
| torchvision | BSD |
| open_clip_torch | MIT |
| pillow | MIT-CMU |
| timm | Apache 2.0 |
| typer / rich / click | MIT / MIT / BSD-3-Clause |
| httpx / httpcore / h11 | BSD-3-Clause |
| anyio | MIT |
| certifi | MPL-2.0 |
| filelock | MIT |
| fsspec | BSD-3-Clause |
| jinja2 | BSD-3-Clause |
| markdown_it_py / markupsafe | MIT / BSD-3-Clause |
| packaging | Apache-2.0 OR BSD-2-Clause |
| pyyaml | MIT |
| regex | Apache-2.0 AND CNRI-Python |
| safetensors | Apache 2.0 |
| tqdm | MPL-2.0 AND MIT |
| typing_extensions | PSF-2.0 |
| ftfy | Apache-2.0 |
| idna | BSD-3-Clause |
| pygments | BSD-2-Clause |
| wcwidth / mdurl / shellingham | MIT / MIT / ISC |
| annotated_doc | MIT |
| pip / setuptools | MIT |

**以上全部为宽松许可证（MIT / BSD / Apache-2.0 / MPL-2.0 / ISC / PSF）。其中 MIT / BSD / Apache-2.0 与 GPL-3.0 兼容，可自由并入 GPL 项目；MPL-2.0 与 GPL-3.0 兼容（MPL-2.0 官方提供 GPL 兼容条款）。因此将本项目以 GPL-3.0 授权**完全合规**，无需替换任何第三方依赖。**

> PyMuPDF 例外说明见上方主要依赖表格：其为 AGPL-3.0 双许可的可选增强组件，本地使用不受影响；对外分发服务时可排除（PDF 提取自动回退 pdfminer.six）。

## 资源目录

| 组件 | 许可证 | 说明 |
|------|--------|------|
| resources/qr/*.png | 自制 | 作者收款码图片，无版权问题 |

## 已清理的组件

| 组件 | 原许可证 | 处理 |
|------|---------|------|
| resources/antiword（antiword.exe） | **GPL-2.0** | ⚠️ 已删除。GPL 有传染性，Web 版不使用 |
| resources/tesseract / tesseract_lite | Apache 2.0 | 已删除。Web 版未调用 OCR（遗留资源） |
| resources/fonts | OFL-1.1 | 已删除。Web 版未引用该字体 |

## 兼容性结论

项目保留的全部第三方组件均为**宽松许可证**，与 **GPL-3.0** 完全兼容：

- **MIT / BSD / Apache-2.0**：与 GPL-3.0 兼容，可自由并入 GPL 项目
- **MPL-2.0**：文件级 copyleft，MPL-2.0 官方条款允许以 GPL-3.0 兼容方式组合分发
- **OFL（字体，已移除）**：与 GPL 兼容

**因此本项目采用 GPL-3.0 License 完全合规。**
