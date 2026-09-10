# Bug 报告自动生成工具

## 功能
- 导入 `CSV / Excel(.xlsx/.xls)` 的 Bug 数据。
- 按固定模板输出《质量验收与缺陷分析报告》HTML，并保存为可被 Word 打开的 `.doc` 文件。
- 默认输出到 `F:\`；若没有 F 盘，则回退到桌面。
- 支持可选 AI 增强：可在网页端填写 `API Base URL / API Key / Model`，调用 OpenAI 兼容接口生成发布建议。

## 输入字段（必须包含）
- `标题`
- `任务ID`
- `执行者`
- `状态`
- `优先级`
- `参与者`

## 已实现规则
- 参与者提取：合并 `执行者` + `参与者` 并去重。
- 模块归类：按标题关键词归入：
  - 去水印/去字幕（含“去水印”“去字幕”“Logo”）
  - 数据埋点（含“埋点”）
  - 多语言适配（含“多语言”）
  - 视频翻译 / AI字幕 / 设置页 / 编辑页
  - 未命中归为“其他模块”
- 统计项：Bug 总数、解决率（已解决/总数）、高优先级占比。
- 核心深度分析：
  - 权重：`>5 -> 3`，`2~5 -> 2`，`1 -> 1`
  - 密度：`数量 / 权重`
  - 表格按 Bug 数从大到小排序
  - Top1 模块动态判定为“UI交互问题”或“功能逻辑问题”
- 附录排序（严格）：
  - 未关闭：`重新打开 > 待处理 > 开发中 > 暂不处理`
  - 已关闭：`已解决 > 已完成 > 开发完成`
  - 组内：同状态下高优先级在前，再按任务ID升序

## 使用
1. 安装 Python 3.9+。
2. （可选）安装 Excel 依赖（读取 `.xlsx/.xls` 时需要）：
   ```bash
   pip install pandas openpyxl
   ```
3. 运行：
   ```bash
   python generate_bug_report.py
   ```
4. 在弹窗中选择 Bug 文档并输入项目名，即可自动生成报告。

## 网页版

本地 Flask 版：

```bash
python app_web.py
```

打开 `http://127.0.0.1:8512` 使用。

Streamlit 版：

```bash
streamlit run streamlit_app.py
```

Streamlit Community Cloud 部署参数：

- Repository: `18160145171/yana_bug`
- Branch: `main`
- Main file path: `streamlit_app.py`

## 备注
- CSV 建议使用 `UTF-8 with BOM` 编码，避免中文乱码。
- 样式已按 Word 兼容方式输出（用于 `.doc` 打开）。
