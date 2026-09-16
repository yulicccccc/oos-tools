import re

failures = [
    {'cat': 'personnel sampling (right touch)', 'obs': '1 CFU on RT', 'etx': 'ETX-260817-0447', 'id': 'Micrococcus luteus', 'time': 'daily', 'note': 'plate was desiccated on 5 day read.'},
    {'cat': 'settling plates', 'obs': '2 CFUs on sett 2', 'etx': 'ETX-260817-0507', 'id': 'Insufficient read Gram (+) rods', 'time': 'daily', 'note': 'plate was desiccated on 5 day read.'},
    {'cat': 'weekly active air sampling', 'obs': '1 CFU in ISO 8 115', 'etx': 'ETX-260817-0370', 'id': 'Gram (+) short rods', 'time': 'weekly', 'note': ''},
    {'cat': 'weekly surface sampling', 'obs': '2 CFUs on floor in 115A ISO 7', 'etx': 'ETX-260817-0366', 'id': 'Bacillus megaterium', 'time': 'weekly', 'note': ''},
]

pass_daily_clean = ['surface sampling']
pass_wk_clean = []

narr_parts = []
if pass_daily_clean:
    narr_parts.append('no microbial growth was observed in surface sampling')
if pass_wk_clean:
    narr_parts.append('Additionally, weekly active air and surface sampling showed no microbial growth')
narr = 'Upon analyzing the environmental monitoring results, ' + '. '.join(narr_parts) + '.'

daily_fails = [f['cat'] for f in failures if f['time'] == 'daily']
weekly_fails = [f['cat'] for f in failures if f['time'] == 'weekly']
intro_parts = []
if daily_fails:
    intro_parts.append(' and '.join(daily_fails) + ' on the date of testing')
if weekly_fails:
    intro_parts.append(' and '.join(weekly_fails) + ' from the week of testing')
fail_intro = f"However, microbial growth was observed during {' and '.join(intro_parts)}."

detail_sentences = []
for i, f in enumerate(failures):
    is_plural = bool(re.search(r'\d+', f['obs']) and int(re.search(r'\d+', f['obs']).group()) > 1)
    verb_detect = 'were' if is_plural else 'was'
    noun_id = 'organisms were' if is_plural else 'organism was'
    method_text = 'differential staining' if 'gram' in f['id'].lower() else 'microbial identification'
    note_txt = f" (Note: {f['note']})" if f.get('note') and f['note'].strip() and f['note'].strip().lower() != 'none' else ""
    base_sentence = f"{f['obs']} {verb_detect} detected during {f['cat']} and was submitted for {method_text} under sample ID {f['etx']}, where the {noun_id} identified as {f['id']}{note_txt}"
    lead = "Specifically" if i==0 else "Additionally" if i==1 else "Furthermore" if i==2 else "Also"
    detail_sentences.append(f"{lead}, {base_sentence}.")

det = f"{fail_intro} {' '.join(detail_sentences)}"
print('NARRATIVE:')
print(narr)
print('\nDETAILS:')
print(det)
