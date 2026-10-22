# Excel 单元格转 PDF

上传 `.xlsx` / `.xlsm`，选择工作表和单元格（如 `A1, C3, D2:D8`），每个单元格生成一个单页 PDF，并打包成 ZIP 下载。

支持自定义宽高（毫米）、字号、边距、水平及垂直对齐；中文字体嵌入、自动换行和缩小字号。单次最多选择 5000 个单元格（含所有工作表）。PDF 按工作表目录和单元格地址命名。

## 运行

需要 Python 3.10 或更新版本。

```bash
python -m venv .venv
.venv/bin/pip install -r requirements.txt
.venv/bin/streamlit run app.py --server.headless true --browser.gatherUsageStats false
```

当前云环境已安装在 `/workspace/venvs/xslzpdf`，可在本仓库运行：

```bash
/workspace/venvs/xslzpdf/bin/streamlit run app.py --server.headless true --browser.gatherUsageStats false
```

## 使用

1. 上传文件，选择一个或多个工作表。
2. 输入单元格或区域；选中的所有表应用同一个范围。
3. 设置 PDF 宽高、字号、边距和对齐。
4. 点击生成，再下载 ZIP。

导出原始文字、数字和换行，不复制 Excel 格式、图片或边框；数字显示格式（例如百分比、货币）不保留。合并单元格仅左上角有值。公式使用上次保存的缓存值，不在服务器上计算，请先用 Excel 重新计算并保存。旧版 `.xls` 需另存为 `.xlsx`。不支持加密文件。

内容过长会自动缩小至最低 4pt；仍放不下则报告具体单元格，不截断内容。批量文件在内存中生成，建议大文件分批处理。上传文件不会写入项目目录，应用不包含第三方上传功能。

## 验证

```bash
/workspace/venvs/xslzpdf/bin/python -m unittest -v
```

测试覆盖单元格区域解析、越界与数量限制、中文内容和嵌入字体、页面尺寸、长文本、空值、零值、布尔值以及 ZIP 文件结构。
