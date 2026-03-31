# 🛠️ 组件说明书：utils.py (公共工具箱)

> **定位 (The Big Picture)**：本项目所有子页面（如 ScanRDI, Celsis 等）共享的核心逻辑、数据字典和基础 UI 配置。

## 1. 核心功能区块 (Core Modules)

* **a. UI 样式控制**
    * pply_eagle_style(): 统一隐藏默认菜单，设定 Eagle Analytical 专属的按钮圆角和样式。
* **b. 核心数据字典**
    * get_full_name(initial): 将分析师缩写（如 "QC", "DS"）转换为全名。新增员工只需在此更新。
* **c. 物理逻辑与算法**
    * get_room_logic(bsc_id): 根据 BSC 单双号法则（奇数 A，偶数 B），反向推导对应的 Cleanroom 和 Suite 编号。
* **d. 格式化小工具**
    * ordinal(n): 数字转序数词 (1 -> 1st)。
    * 
um_to_words(n): 数字转英文单词 (1 -> one)。

## 2. 修改指南 (Action Items)

* **新增员工**：在 get_full_name 字典中添加 "缩写": "全名" 键值对。
* **新增 BSC 机器**：在 get_room_logic 中将新机器编号归入对应的 Suite 组。

## 3. 严格禁忌 (What NOT to do)

* ❌ 绝对不要在这里放置只属于某个特定测试（如仅限于 ScanRDI）的长篇文字段落。
* ❌ 绝对不要在这里存放表单提交（Submit）的交互逻辑。
