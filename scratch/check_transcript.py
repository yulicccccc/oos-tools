import json, sys

path = r'C:\Users\qchen\.gemini\antigravity\brain\b3f79cbb-02fb-4b92-abd0-5be59a2950f7\.system_generated\logs\transcript.jsonl'
with open(path, 'r', encoding='utf-8', errors='ignore') as f:
    for line in f:
        d = json.loads(line)
        if d.get('type') == 'USER_INPUT':
            print(f"=== STEP {d.get('step_index')} ===")
            content = d.get('content') or ''
            print(content[:500])
