# 🔦 组件说明书：get_fields.py (座位测绘仪)

> **定位 (The Big Picture)**：开发阶段使用的幕后工具 (Behind-the-scenes tool)。一次性脚本 (Throwaway script)。

## 1. 核心功能
* 当拿到一个新的、由老板提供的 PDF 模板时，运行此脚本。
* 它会扫描 PDF，提取所有“填空框”在系统内部的真实暗号（如 Text Field 50）。
* 将提取出的名字复制到主代码的 pdf_map 字典中。日常跑报告时**完全用不到**。
