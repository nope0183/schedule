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

MONTHS_RU = ['января', 'февраля', 'марта', 'апреля', 'мая', 'июня',
             'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря']

def format_date(d):
    return f"{d.day} {MONTHS_RU[d.month - 1]} {d.year}"

date = "Неизвестная дата"

# ВАЖНО: в реальных выгрузках дата обычно вообще не лежит ни на одной
# ячейке таблицы — она зашита прямо в НАЗВАНИЕ ЛИСТА, например
# "02.09.2026 среда". Раньше парсер искал дату только по ячейкам и
# поэтому всегда получал "Неизвестная дата". Сначала проверяем title.
sheet_title = (ws.title or '').strip()
m = re.search(r'(\d{1,2})\.(\d{1,2})\.(\d{4})', sheet_title)
if m:
    day, month, year = int(m.group(1)), int(m.group(2)), int(m.group(3))
    try:
        formatted = format_date(datetime.date(year, month, day))
        weekday_word = sheet_title[m.end():].strip(' ,')
        # Убираем возможные лишние слова после дня недели (например "студ", "преп" и т.д.)
        weekday_match = re.match(r'([а-яА-ЯёЁ]+)', weekday_word)
        if weekday_match:
            weekday_word = weekday_match.group(1)
        date = f"{formatted}, {weekday_word}" if weekday_word else formatted
    except ValueError:
        pass

# Фолбэк на случай, если в какой-то выгрузке дата всё же лежит в ячейке
# (первые 50 строк, первые 5 столбцов) — текстом или настоящим типом
# datetime/date (тогда str(v) дал бы "2026-09-01 00:00:00", под старые
# текстовые шаблоны не подходящее, поэтому проверяем тип отдельно).
if date == "Неизвестная дата":
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
            if isinstance(v, (datetime.datetime, datetime.date)):
                date = format_date(v)
                break
            s = str(v)
            for pattern in DATE_PATTERNS:
                fm = re.search(pattern, s)
                if fm:
                    date = fm.group(0)
                    break
            if date != "Неизвестная дата":
                break
        if date != "Неизвестная дата":
            break

# Ищем группы во 2-й строке (проверяем колонки B-AK, то есть 2-37)
groups = []
group_cols = {}
for col in range(2, 38):  # B до AK включительно (2 до 37)
    val = ws.cell(row=2, column=col).value
    if val is None:
        continue
    val_str = str(val).strip()
    if re.match(r'^\d{4}[а-яА-Я]?$', val_str):
        groups.append(val_str)
        group_cols[val_str] = col

if not groups:
    print("ERROR: Groups not found in row 2 (columns B-AK)!")
    sys.exit(1)

print(f"Found {len(groups)} groups, date: {date}")

ROOM_RE = re.compile(r'\((\d{1,3}[а-яА-Я]?)\)')
CODE_RE = re.compile(r'^[А-Я]{1,4}\.?\s*\d+$')          # шифр дисциплины: "ООД.04"
ROOM_ONLY_RE = re.compile(r'^\d{1,3}[а-яА-Я]?\.?$')      # кабинет без скобок отдельной строкой
TEACHER_RE = re.compile(r'[А-ЯЁ][а-яё]+\s+[А-ЯЁ]\.')      # "Фамилия И."


def parse_record(cell_text):
    """Разбирает ОДНУ ячейку (один физический блок "код\\nпредмет\\nпреподаватель (каб.)")
    построчно, в естественном порядке сверху вниз — а не вперемешку с другими
    ячейками и не в алфавитном порядке, как было раньше. Это и есть источник
    прежних багов с кабинетами и предметами:
    - кабинет мог стоять отдельной строкой в скобках, например "(41)" —
      после вырезания скобок оставалась пустая строка, которая раньше течением
      кода записывалась как subject = '' и затирала уже найденный предмет;
    - alphabetical-сортировка строк вообще не гарантирует порядок
      код/предмет/преподаватель/кабинет.
    """
    subject, teacher, room = '', '', ''
    for raw_line in cell_text.split('\n'):
        line = raw_line.strip()
        if not line:
            continue

        rm = ROOM_RE.search(line)
        if rm:
            if not room:
                room = rm.group(1)
            line = ROOM_RE.sub('', line).strip()
            if not line:
                # строка была ЦЕЛИКОМ кабинетом в скобках, например "(41)" —
                # после вырезания скобок обрабатывать больше нечего, и, в
                # отличие от старого кода, мы НЕ затираем subject пустой строкой
                continue

        if CODE_RE.match(line):
            continue
        if ROOM_ONLY_RE.match(line) or line == '-':
            if not room:
                room = line
            continue
        if TEACHER_RE.search(line):
            teacher = line
            continue
        # похоже на название предмета; если предмет уже был найден в этой же
        # ячейке (например, длинное название перенесено на 2 строки) — дописываем
        subject = f'{subject} {line}'.strip() if subject else line
    return subject, teacher, room


# 3. Парсим пары — согласно КООРДИНАТНОЙ структуре таблицы:
# Пара 1: строки 14-15
# Пара 2: строки 27-28
# Пара 3: строки 40-41
# Пара 4: строки 53-54
# Пара 5: строки 66-67
# Пара 6: строки 69-70
#
# Столбцы: B-AK (каждый столбец = одна группа из row 2)
LESSON_ROWS = {
    '1': [14, 15],
    '2': [27, 28],
    '3': [40, 41],
    '4': [53, 54],
    '5': [66, 67],
    '6': [69, 70],
}

raw_records = {}   # {'1': {'1161': [ {subject,teacher,room}, ... ]}}

for lesson_num, row_range in LESSON_ROWS.items():
    raw_records[lesson_num] = {}
    
    for row_idx in row_range:
        if row_idx > max_row:
            continue
            
        for gname, cidx in group_cols.items():
            cv = ws.cell(row=row_idx, column=cidx).value
            if cv is None:
                continue
            val = str(cv).strip()
            
            if not val:
                continue

            # Парсим содержимое ячейки как запись о паре
            subject, teacher, room = parse_record(val)
            if not (subject or teacher or room):
                continue

            record = {'subject': subject, 'teacher': teacher, 'room': room}
            records = raw_records[lesson_num].setdefault(gname, [])
            if record not in records:
                records.append(record)

wb.close()

# 4. Формируем JSON
result = {'date': date, 'groups': sorted(groups), 'schedule': {}}
for gname in groups:
    lessons = []
    for num in sorted(raw_records.keys(), key=int):
        records = raw_records[num].get(gname)
        if not records:
            continue

        if len(records) == 1:
            subject, teacher, room = records[0]['subject'], records[0]['teacher'], records[0]['room']
        else:
            # Несколько РАЗНЫХ записей на одну пару у одной группы — это
            # деление на подгруппы (например, две группы английского языка
            # с разными преподавателями/кабинетами). Раньше вторая подгруппа
            # молча терялась, потому что subject/teacher/room в цикле просто
            # перезаписывались последним значением. Показываем все варианты.
            def _joined(key):
                seen = []
                for r in records:
                    if r[key] and r[key] not in seen:
                        seen.append(r[key])
                return ' / '.join(seen)
            subject, teacher, room = _joined('subject'), _joined('teacher'), _joined('room')

        lessons.append({'num': num, 'subject': subject, 'teacher': teacher, 'room': room})
    result['schedule'][gname] = lessons

with open(JSON_FILE, 'w', encoding='utf-8') as f:
    json.dump(result, f, ensure_ascii=False, indent=2)

with open(HASH_FILE, 'w') as f:
    f.write(new_hash)

print(f'Done! {len(groups)} groups, {len(raw_records)} lesson blocks, date: {date}')
