"""DOCX 导出的内部模块划分。

Worker 通过本包按职责选择结构化流式导出或高保真绝对定位导出；
各模块保持独立，避免单文件继续膨胀。

当前包含：

* `text_utils`：表格 HTML 解析、代码行识别与文本压缩，无内部依赖；
* `table_render`：PdfTable → Word 固定布局表格的渲染。
* `structured`：正文流式段落、原生表格、OMML 公式与内嵌图片；
* `fidelity`：按源坐标绝对定位的兼容导出路径。
"""
