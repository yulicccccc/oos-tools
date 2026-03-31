# 🍽️ 组件说明书：pages/ScanRDI.py (明档餐厅 / 前端 UI)

> **定位 (The Big Picture)**：ScanRDI 报告的专属前端界面。负责收集客人点单（用户输入），并端出成品菜（下载报告）。

## 1. 核心功能区块
* **UI 布局 (st.columns, st.text_input)**：搭建 1 到 5 步的全部输入表格。
* **呼叫后厨 (import scan_logic)**：点击 "Generate" 按钮时，把收集到的数据传给 scan_logic.py 进行底层处理。
* **上菜服务 (st.download_button)**：把后厨处理好的 Word 和 PDF 文件提供给用户下载。

## 2. 严格禁忌 (What NOT to do)
* ❌ 尽量不要在这里写超过 10 行的纯文本段落拼接（比如 Equipment Summary）。
* ❌ 复杂的正则表达式提取和长逻辑判断，请统统赶到 scan_logic.py 里去！
