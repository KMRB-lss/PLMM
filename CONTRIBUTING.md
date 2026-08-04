# 感谢你愿意为本项目贡献代码！🎉

请遵循以下规范，让协作更顺畅：

## 开发环境
- Python 3.8+（建议 3.11）
- 前端为单文件 `webui.html`，无需构建步骤
- 后端为单文件 `server.py`，使用 Python 标准库 `http.server`

## 本地启动
```bash
python server.py 8890
# 浏览器打开 http://localhost:8890
```

## 提交规范
- 提交信息使用中文，遵循 `<类型>: <描述>` 格式
  - `feat:` 新功能
  - `fix:` 修复 Bug
  - `docs:` 文档
  - `refactor:` 重构
  - `style:` 样式
  - `perf:` 性能
- 示例：`fix: 修复预览列表列错位问题`

## 代码风格
- Python 遵循 PEP 8，注释使用中文
- JS 保持与现有代码一致的两空格缩进风格
- 修改 `webui.html` 后请用浏览器验证无 JS 报错

## 提交前检查
- [ ] `python -m py_compile server.py` 通过
- [ ] 浏览器打开页面无控制台报错
- [ ] 基本流程可用：加载文件夹 → 生成预览 → 开始重命名 → 撤销

## 分支模型
- `main` 为主分支，保持可发布状态
- 新功能 / 修复从 `main` 拉取特性分支，完成后提 Pull Request

## 许可证
本项目基于 **GNU GPL v3.0（GPL-3.0）** 开源，提交代码即视为同意在该许可证下发布。GPL-3.0 具有 copyleft 特性：基于本项目的衍生作品在分发时必须以相同许可证开源。
