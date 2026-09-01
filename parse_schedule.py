import hashlib
import json
import re
import sys
import requests
from io import BytesIO
from openpyxl import load_workbook

URL = 'https://serp-koll.ru/images/ep/k1/rasp1.xlsx'
JSON_FILE = 'schedule.json'
HASH_FILE = 'schedule.hash'

try:
    resp = requests.get(URL, timeout=30)
    if resp.status_code != 200:
        print(f'HTTP {resp.status_code}')
        sys.exit(0)
except Exception as e:
    print(f'Download error: {e}')
    sys.exit(0)

content = resp.content
new_hash = hashlib.md5(content).hexdigest()

try:
    with open(HASH_FILE, 'r') as f:
        old_hash = f.read().strip()
except FileNotFoundError:
    old_hash = ''

if new_hash == old_hash:
    print('No changes detected.')
    sys.exit(0)

print(f'Schedule changed! Old: {old_hash[:8] if old_hash else "none"}, New: {new_hash[:8]}')

wb = load_workbook(filename=BytesIO(content), read_only=True, data_only=True)
ws = wb.active

date_cell = ws.cell(row=1, column=1).value
date_str = str(date_cell) if date_cell else ''
match = re.search(r'(\d{1,2}[.\s]\w+[.\s]\d{4}|\d{2}\.\d{2}\.\d{4})', date_str)
date = match.group(1) if match else 'Неизвестная дата'

groups = []
group_cols = {}
col = 3
while True:
    val = ws.cell(row=2, column=col).value
    if val is None:
        break
    val_str = str(val).strip()
    if re.match(r'^\d{4}[а-яА-Я]?$', val_str):
        groups.append(val_str)
        group_cols[val_str] = col
    col += 1

raw_data = {}
current_lesson = None
for row_idx in range(3, ws.max_row + 1):
    cell_a = ws.cell(row=row_idx, column=1).value
    if cell_a is not None:
        s = str(cell_a).strip()
        if s.isdigit() and 1 <= int(s) <= 7:
            current_lesson = s
    if current_lesson is None:
        continue
    if current_lesson not in raw_data:
        raw_data[current_lesson] = {}
    for gname, cidx in group_cols.items():
        cv = ws.cell(row=row_idx, column=cidx).value
        if cv is None:
            continue
        val = str(cv).strip()
        low = val.lower()
        if any(x in low for x in ['пара', 'объединён', 'разделён']):
            continue
        if len(val) < 2:
            continue
        if gname not in raw_data[current_lesson]:
            raw_data[current_lesson][gname] = set()
        for line in val.split('\n'):
            line = line.strip()
            if line and len(line) >= 2:
                raw_data[current_lesson][gname].add(line)

wb.close()

result = {'date': date, 'groups': sorted(groups), 'schedule': {}}
for gname in groups:
    lessons = []
    for num in sorted(raw_data.keys(), key=int):
        if gname not in raw_data[num]:
            continue
        lines = sorted(raw_data[num][gname])
        subject, teacher, room = '', '', ''
        for line in lines:
            rm = re.search(r'\((\d{1,3}[а-яА-Ям]?)\)', line)
            if rm:
                room = rm.group(1)
                line = re.sub(r'\s*\(\d{1,3}[а-яА-Ям]?\)\s*', '', line).strip()
            if re.match(r'^[А-Я]{1,4}\.?\s*\d', line):
                continue
            if re.match(r'^\d{1,3}[а-яА-Ям]?\.?$', line) or line == '-':
                if not room:
                    room = line
            elif re.search(r'[А-ЯЁ][а-яё]+\s+[А-ЯЁ]\.', line):
                teacher = line
            else:
                subject = line
        lessons.append({'num': num, 'subject': subject, 'teacher': teacher, 'room': room})
    result['schedule'][gname] = lessons

with open(JSON_FILE, 'w', encoding='utf-8') as f:
    json.dump(result, f, ensure_ascii=False, indent=2)

with open(HASH_FILE, 'w') as f:
    f.write(new_hash)

print(f'Done! {len(groups)} groups parsed, date: {date}')
