import json, sys
sys.stdout.reconfigure(encoding='utf-8')

transcript_path = r'C:\Users\qchen\.gemini\antigravity\brain\b48c93a2-1dda-4762-93a5-143172ae3c6f\.system_generated\logs\transcript.jsonl'
with open(transcript_path, 'r', encoding='utf-8') as f:
    for idx, line in enumerate(f):
        if 4200 <= idx <= 4215:
            d = json.loads(line)
            print(f"=== Line {idx} | Type: {d.get('type')} ===")
            c = str(d.get('content', ''))
            t = str(d.get('thinking', ''))
            if c:
                print("CONTENT:", c[:300].replace('\n', ' '))
            if t:
                print("THINKING:", t[:300].replace('\n', ' '))
