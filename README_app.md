# 🛎️ 组件说明书：app.py (一楼大堂 / 迎宾前台)

> **定位 (The Big Picture)**：系统的正大门和主入口 (Main Entry Point)。负责迎接用户并提供全局导航。

## 1. 核心功能区块
* **大门招牌 (st.set_page_config)**：设置浏览器标签页名字为 "Microbiology Platform" 和 🦅 图标。
* **呼叫后勤 (pply_eagle_style)**：调用 utils 里的方法，刷上深蓝色侧边栏。
* **迎宾致辞 (st.title & st.markdown)**：展示 Eagle Trax 的欢迎语和可选模块指南。

## 2. 严格禁忌 (What NOT to do)
* ❌ 绝对不要在这里写任何表单输入框 (Input fields)。
* ❌ 绝对不要在这里写任何数据处理和报告生成的逻辑。前台不负责炒菜！
