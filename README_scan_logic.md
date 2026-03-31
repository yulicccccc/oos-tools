# 🔪 组件说明书：scan_logic.py (封闭式后厨 / 业务逻辑)

> **定位 (The Big Picture)**：ScanRDI 的专属备菜间。处理所有脏活、累活 (Heavy lifting) 和烧脑的逻辑。

## 1. 核心功能区块
* **邮件解析 (parse_email_text)**：用正则表达式疯狂抓取 ETX 编号、日期等信息。
* **长文生成 (generate_equipment_text, generate_narrative_and_details)**：根据输入条件，像拼积木一样拼凑出几十上百个单词的专业段落。
* **画表格 (create_table_pdf)**：用 ReportLab 从零开始画出补充表格的 PDF。

## 2. 严格禁忌 (What NOT to do)
* ❌ 绝对不要在这里出现任何画网页排版的代码（例如 st.columns, st.button）。后厨人员不允许跑到大堂去摆桌椅！
