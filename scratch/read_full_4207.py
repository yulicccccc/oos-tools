import json, sys
sys.stdout.reconfigure(encoding='utf-8')

transcript_path = r'C:\Users\qchen\.gemini\antigravity\brain\b48c93a2-1dda-4762-93a5-143172ae3c6f\.system_generated\logs\transcript_full.jsonl'
with open(transcript_path, 'r', encoding='utf-8') as f:
    for idx, line in enumerate(f):
        if idx == 4207:
            d = json.loads(line)
            print(d.get('content'))
