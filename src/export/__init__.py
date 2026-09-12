"""DOCX 导出的内部模块划分。

`src/pdf_to_word_exporter.py` 保持为对外入口，继续导出全部导出函数与
工具；本包按职责拆分实现，避免单文件继续膨胀。

当前包含：

* `text_utils`：表格 HTML 解析、代码行识别与文本压缩，无内部依赖；
* `table_render`：PdfTable → Word 固定布局表格的渲染。
"""
