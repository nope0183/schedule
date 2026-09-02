#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import re
import json
import hashlib
import requests
import openpyxl
from datetime import datetime
import sys

# ===== Конфигурация =====
URL = "https://serp-koll.ru/images/ep/k1/rasp1.xlsx"
HASH_FILENAME = "schedule.hash"
XLSX_FILENAME = "rasp1.xlsx"
JSON_FILENAME = "schedule.json"

def get_hash(filename):
    """Вычисляет MD5 файла."""
    if not os.path.exists(filename):
        return None
    with open(filename, "rb") as f:
        return hashlib.md5(f.read()).hexdigest()

def save_hash(filename, hash_val):
    """Сохраняет хеш в файл."""
    with open(filename, "w") as f:
        f.write(hash_val)

def download_file(url, filename):
    """Скачивает файл по URL."""
    print(f"⬇️  Downloading {url}...")
    try:
        response = requests.get(url, timeout=30)
        response.raise_for_status()
        with open(filename, "wb") as f:
            f.write(response.content)
        print(f"✅ Downloaded: {filename}")
        return True
    except Exception as e:
        print(f"❌ Download error: {e}")
        return False

def clean_value(val):
    """Очищает значение ячейки."""
    if val is None:
        return ""
    return str(val).strip()

def parse_schedule_excel(xlsx_path):
    """Парсит расписание из Excel файла."""
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb.active
    
    print(f"📄 Лист: {ws.title}, Строк: {ws.max_row}, Столбцов: {ws.max_column}")
    
    # Строка 1: дата расписания
    date_str = ""
    for col in range(1, min(ws.max_column + 1, 15)):
        cell_val = clean_value(ws.cell(1, col).value)
        if cell_val:
            # Ищем дату
            match = re.search(r'(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{2,4})', cell_val)
            if match:
                day, month, year = int(match.group(1)), int(match.group(2)), int(match.group(3))
                if year < 100:
                    year += 2000
                try:
                    date_obj = datetime(year, month, day)
                    date_str = date_obj.strftime("%d.%m.%Y")
                    break
                except:
                    pass
    
    if not date_str:
        date_str = datetime.now().strftime("%d.%m.%Y")
    
    print(f"📅 Дата: {date_str}")
    
    # Строка 2: названия групп (начиная со столбца B)
    groups = []
    group_cols = {}
    for col in range(2, ws.max_column + 1):
        cell_val = clean_value(ws.cell(2, col).value)
        if cell_val and re.match(r'^[А-Яа-я0-9]{3,5}$', cell_val):
            if cell_val not in groups:
                groups.append(cell_val)
                group_cols[cell_val] = col
    
    print(f"👥 Групп найдено: {len(groups)}")
    print(f"   {', '.join(groups)}")
    
    # Столбец A: номера пар и данные
    # Каждая пара занимает 2 строки
    schedule = {group: [] for group in groups}
    
    row = 3
    while row <= ws.max_row:
        # Читаем номер пары из первой строки блока
        pair_num_cell = clean_value(ws.cell(row, 1).value)
        
        # Проверяем, это ли номер пары (1-7)
        pair_num = None
        if pair_num_cell and re.match(r'^\d+$', pair_num_cell):
            pair_num = int(pair_num_cell)
            if not (1 <= pair_num <= 7):
                pair_num = None
        
        if pair_num is not None:
            # Это начало блока пары - занимает 2 строки
            row1 = row
            row2 = row + 1
            
            # Для каждой группы берём данные из обеих строк
            for group, col in group_cols.items():
                cell1 = clean_value(ws.cell(row1, col).value)
                cell2 = clean_value(ws.cell(row2, col).value)
                
                # Объединяем содержимое обеих ячеек
                content = (cell1 + "\n" + cell2).strip()
                
                if content:
                    lesson = parse_lesson_content(pair_num, content)
                    if lesson:
                        schedule[group].append(lesson)
            
            row += 2  # Переходим на следующую пару (2 строки)
        else:
            row += 1
    
    return {
        "date": date_str,
        "groups": groups,
        "schedule": schedule
    }

def parse_lesson_content(pair_num, content):
    """Парсит содержимое пары из 2 ячеек."""
    if not content:
        return None
    
    lines = [l.strip() for l in content.split('\n') if l.strip()]
    
    subject = ""
    teacher = ""
    room = ""
    
    for line in lines:
        # Пропускаем служебные строки
        if any(x in line.lower() for x in ["разделё", "объедин", "физ", "консульт"]):
            continue
        
        # Ищем преподавателя (ФИО)
        teacher_match = re.search(r'([А-ЯЁ][а-яё]+\s+[А-ЯЁ]\.\s*[А-ЯЁ]\.?)', line)
        if teacher_match and not teacher:
            teacher = teacher_match.group(1).strip()
        
        # Ищем аудиторию (номер или в скобках)
        room_match = re.search(r'[\(]?(\d+[а-яА-ЯёЁ]?)[\)]?', line)
        if room_match and not room:
            room_candidate = room_match.group(1)
            # Проверяем, что это не время
            if not re.match(r'^\d{1,2}:\d{2}', room_candidate):
                room = room_candidate
        
        # Остаток — предмет
        if line and not teacher_match and not room_match:
            if not subject:
                subject = line
    
    # Минимальная валидация
    if not subject and not teacher and not room:
        return None
    
    return {
        "num": str(pair_num),
        "subject": subject,
        "teacher": teacher,
        "room": room,
    }

def main():
    print("=" * 60)
    print("🎓 Schedule Parser v3.0")
    print("=" * 60)
    
    # Проверяем изменения
    old_hash = get_hash(XLSX_FILENAME)
    new_hash = None
    
    if download_file(URL, XLSX_FILENAME):
        new_hash = get_hash(XLSX_FILENAME)
        
        if old_hash == new_hash:
            print("📌 Файл не изменился")
            return 0
    else:
        print("⚠️  Скачивание не удалось, используем старый файл")
        if not os.path.exists(JSON_FILENAME):
            return 1
        return 0
    
    try:
        result = parse_schedule_excel(XLSX_FILENAME)
        
        with open(JSON_FILENAME, "w", encoding="utf-8") as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        
        print(f"✅ Сохранено: {JSON_FILENAME}")
        
        if new_hash:
            save_hash(HASH_FILENAME, new_hash)
        
        # Статистика
        total = sum(len(v) for v in result["schedule"].values())
        print(f"\n📊 Статистика:")
        print(f"   Групп: {len(result['groups'])}")
        print(f"   Всего пар: {total}")
        
        return 0
    
    except Exception as e:
        print(f"❌ Ошибка: {e}")
        import traceback
        traceback.print_exc()
        return 1

if __name__ == "__main__":
    sys.exit(main())
