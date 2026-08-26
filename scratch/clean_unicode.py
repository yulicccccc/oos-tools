with open('em_logic.py', 'r', encoding='utf-8') as f:
    text = f.read()

text = text.replace('H₂O₂', 'H2O2').replace('’', "'").replace('–', '-')
with open('em_logic.py', 'w', encoding='utf-8') as f:
    f.write(text)
print('Successfully cleaned unicode characters in em_logic.py!')
