import datetime
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

# 1. Скачиваем файл
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

# 2. Парсим Excel
wb = load_workbook(filename=BytesIO(content), data_only=True)
ws = wb.active

# ВАЖНО: "разъединяем" объединённые ячейки — openpyxl отдаёт None
# для всех ячеек merged-диапазона, кроме верхней левой. Если дата
# или какая-то из пар лежит в объединённой ячейке — без этого шага
# соседние (не top-left) ячейки будут просто пустыми.
for merged_range in list(ws.merged_cells.ranges):
    min_col, min_row, max_col, max_row = merged_range.bounds
    top_left_value = ws.cell(row=min_row, column=min_col).value
    ws.unmerge_cells(str(merged_range))
    for row in range(min_row, max_row + 1):
        for col in range(min_col, max_col + 1):
            ws.cell(row=row, column=col).value = top_left_value

max_row = ws.max_row or 200
max_col = ws.max_column or 50

# Ищем дату в первых 50 строках и 5 столбцах
date = "Неизвестная дата"
MONTHS_RU = ['января', 'февраля', 'марта', 'апреля', 'мая', 'июня',
             'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря']

def format_date(d):
    return f"{d.day} {MONTHS_RU[d.month - 1]} {d.year}"

# несколько шаблонов: "01.09.2025", "1 сентября 2025", "1.09.2025 - 5.09.2025" и т.п.
DATE_PATTERNS = [
    r'\d{1,2}\s+\w+\s+\d{4}',        # 1 сентября 2025
    r'\d{2}\.\d{2}\.\d{4}',          # 01.09.2025
    r'\d{1,2}\.\d{1,2}\.\d{2,4}',    # 1.9.25 / 1.09.2025
]
for r in range(1, min(51, max_row + 1)):
    for c in range(1, min(6, max_col + 1)):
        v = ws.cell(row=r, column=c).value
        if v is None:
            continue
        # ВАЖНО: если в ячейке лежит настоящая дата (тип datetime/date,
        # а не текст), openpyxl отдаёт объект, и str(v) превращается в
        # "2026-09-01 00:00:00" — под DATE_PATTERNS такое никогда не
        # подходило, поэтому дата всегда оставалась "Неизвестная дата".
        # Обрабатываем этот тип отдельно, до преобразования в строку.
        if isinstance(v, (datetime.datetime, datetime.date)):
            date = format_date(v)
            break
        s = str(v)
        for pattern in DATE_PATTERNS:
            m = re.search(pattern, s)
            if m:
                date = m.group(0)
                break
        if date != "Неизвестная дата":
            break
    if date != "Неизвестная дата":
        break

# Ищем группы во 2-й строке (проверяем все колонки)
groups = []
group_cols = {}
for col in range(1, max_col + 1):
    val = ws.cell(row=2, column=col).value
    if val is None:
        continue
    val_str = str(val).strip()
    if re.match(r'^\d{4}[а-яА-Я]?$', val_str):
        groups.append(val_str)
        group_cols[val_str] = col

if not groups:
    print("ERROR: Groups not found in row 2!")
    sys.exit(1)

print(f"Found {len(groups)} groups, date: {date}")

# 3. Парсим пары
raw_data = {}
current_lesson = 0
for row_idx in range(3, max_row + 1):
    cell_a = ws.cell(row=row_idx, column=1).value
    # Если в первой колонке цифра 1-7, это новая пара
    if cell_a is not None:
        s = str(cell_a).strip()
        if s.isdigit() and 1 <= int(s) <= 7:
            current_lesson = int(s)
            # ВАЖНО: здесь НЕТ continue, чтобы обработать данные этой же строки!

    if current_lesson == 0:
        continue

    lk = str(current_lesson)
    if lk not in raw_data:
        raw_data[lk] = {}

    for gname, cidx in group_cols.items():
        cv = ws.cell(row=row_idx, column=cidx).value
        if cv is None:
            continue
        val = str(cv).strip()
        if not val or len(val) < 2 or len(val) > 150:
            continue
        low = val.lower()
        if any(x in low for x in ['пара', 'объединён', 'разделён']):
            continue

        if gname not in raw_data[lk]:
            raw_data[lk][gname] = set()
        for line in val.split('\n'):
            line = line.strip()
            if line and 2 <= len(line) < 150:
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

            # Пропускаем ТОЛЬКО строки, которые ЦЕЛИКОМ являются коротким
            # служебным ярлыком (например пометка подгруппы "П1").
            # Раньше re.match без "$" на конце ловил и обрезал строки
            # предмета вида "ОГСЭ.04 Физическая культура" / "ОП.03 ...",
            # где шифр дисциплины стоит в начале строки, из-за чего
            # subject терялся, хотя teacher/room оставались на месте.
            if re.match(r'^[А-Я]{1,4}\.?\s*\d+$', line):
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
