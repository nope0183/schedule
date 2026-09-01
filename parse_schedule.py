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

# --- Скачиваем файл ---
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
    print('No changes.')
    sys.exit(0)

print(f'Changed! {old_hash[:8] if old_hash else "none"} -> {new_hash[:8]}')

# --- Парсим Excel ---
# ВАЖНО: без read_only=True, иначе max_row/max_column могут быть None
wb = load_workbook(filename=BytesIO(content), data_only=True)
ws = wb.active

max_row = ws.max_row or 200
max_col = ws.max_column or 50

# 1. Ищем дату в первых 50 строках и 5 столбцах
date = "Неизвестная дата"
for r in range(1, min(51, max_row + 1)):
    for c in range(1, min(6, max_col + 1)):
        v = ws.cell(row=r, column=c).value
        if v:
            m = re.search(r'(\d{1,2}[.\s]\w+[.\s]\d{4}|\d{2}\.\d{2}\.\d{4})', str(v))
            if m:
                date = m.group(1)
                break
    if date != "Неизвестная дата":
        break

# 2. Ищем строку с номерами групп (сканируем ВЕСЬ файл)
groups = []
group_cols = {}
groups_row = -1

for r in range(1, max_row + 1):
    temp = []
    for c in range(3, min(max_col + 1, 50)):
        v = ws.cell(row=r, column=c).value
        if v is not None:
            s = str(v).strip()
            if re.match(r'^\d{4}[а-яА-Я]?$', s):
                temp.append((s, c))
    if len(temp) >= 20:
        groups_row = r
        for g, c in temp:
            groups.append(g)
            group_cols[g] = c
        break

if not groups:
    print("ERROR: Groups not found in any row!")
    sys.exit(1)

print(f"Found {len(groups)} groups in row {groups_row}, date: {date}")

# 3. Парсим пары
# Стратегия: сканируем ВСЕ строки файла (кроме строки с группами)
# Номер пары определяем по:
#   a) Явной цифре 1-7 в колонке A
#   b) Подсчёту блоков: строка, где >40% ячеек содержат код дисциплины = начало нового блока
raw_data = {}
current_lesson = 0

for r in range(1, max_row + 1):
    if r == groups_row:
        continue

    # Проверяем явный номер пары в колонке A
    cell_a = ws.cell(row=r, column=1).value
    if cell_a is not None:
        s = str(cell_a).strip()
        if s.isdigit() and 1 <= int(s) <= 7:
            current_lesson = int(s)
            continue

    # Проверяем, является ли строка началом нового блока кодов дисциплин
    code_count = 0
    non_empty = 0
    for gname, cidx in group_cols.items():
        cv = ws.cell(row=r, column=cidx).value
        if cv is not None:
            v = str(cv).strip()
            if v and 2 <= len(v) < 100:
                non_empty += 1
                if re.match(r'^[А-Я]{2,4}\.?\s*\d', v):
                    code_count += 1

    if non_empty > 0 and code_count > non_empty * 0.4:
        current_lesson += 1

    if current_lesson == 0:
        continue

    lk = str(current_lesson)

    # Собираем данные для каждой группы
    for gname, cidx in group_cols.items():
        cv = ws.cell(row=r, column=cidx).value
        if cv is None:
            continue
        val = str(cv).strip()
        if not val or len(val) < 2 or len(val) > 100:
            continue
        low = val.lower()
        if any(x in low for x in ['пара', 'объединён', 'разделён']):
            continue

        if lk not in raw_data:
            raw_data[lk] = {}
        if gname not in raw_data[lk]:
            raw_data[lk][gname] = set()

        for line in val.split('\n'):
            line = line.strip()
            if line and 2 <= len(line) < 100:
                raw_data[lk][gname].add(line)

wb.close()

# 4. Формируем JSON
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

print(f'Done! {len(groups)} groups, {len(raw_data)} lesson blocks, date: {date}')
