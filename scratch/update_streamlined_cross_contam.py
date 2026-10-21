with open('.agents/AGENTS.md', 'rb') as f:
    b = f.read()

idx = b.find(b'1st Shift vs 2nd Shift')
start = b.rfind(b'b. ', 0, idx)
end = b.find(b'c. ', idx)

replacement_agents_b = '''b. **极简高效隔离证明法则 (Streamlined Segregation Phrasing - 聚焦不同人员与不同时间)**:
        - 🚨 **用户定稿极简原则**: 无需堆砌过多冗余琐碎信息（如无需专门交代 Prepper 姓名或强调外围套间层级名称）。
        - 核心证明聚焦三大要素：**班次（Shift）、检验处理分析员（Processing Analyst / 样本处理序号）、主操作设备（Primary Containment Equipment / BSC）**。
        - **法定标准精炼句式**:
          `"However, that sample was processed by [Analyst] ([N]th sample processed) during the [Shift]. The two testing events were segregated by shift, processing analyst and primary containment equipment."`\r\r\n\r\r\n     '''.encode('utf-8')

b_new = b[:start] + replacement_agents_b + b[end:]
with open('.agents/AGENTS.md', 'wb') as f:
    f.write(b_new)
print('Updated .agents/AGENTS.md successfully.')

with open('PRD.md', 'rb') as f:
    b_prd = f.read()

idx_prd = b_prd.find(b'1st Shift vs 2nd Shift')
start_prd = b_prd.rfind(b'2. ', 0, idx_prd)
end_prd = b_prd.find(b'3. ', idx_prd)

replacement_prd_b = '''2. **极简高效隔离证明法则 (Streamlined Segregation Phrasing - 聚焦不同人员与不同时间)**:
          - 🚨 **用户定稿极简原则**: 无需堆砌过多冗余琐碎信息（如无需专门交代 Prepper 姓名或强调外围套间层级名称）。
          - 核心证明聚焦三大要素：**班次（Shift）、检验处理分析员（Processing Analyst / 样本处理序号）、主操作设备（Primary Containment Equipment / BSC）**。
          - **法定标准精炼句式**:
            `"However, that sample was processed by [Analyst] ([N]th sample processed) during the [Shift]. The two testing events were segregated by shift, processing analyst and primary containment equipment."`\r\r\n\r\r\n        '''.encode('utf-8')

b_prd_new = b_prd[:start_prd] + replacement_prd_b + b_prd[end_prd:]
with open('PRD.md', 'wb') as f:
    f.write(b_prd_new)
print('Updated PRD.md successfully.')
