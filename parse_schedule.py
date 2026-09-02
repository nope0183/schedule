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

URL = "https://serp-koll.ru/images/ep/k1/rasp1.xlsx"
HASH_FILENAME = "schedule.hash"
XLSX_FILENAME = "rasp1.xlsx"
JSON_FILENAME = "schedule.json"

def get_hash(filename):
    if not os.path.exists(filename):
        return None
    with open(filename, "rb") as f:
        return hashlib.md5(f.read()).hexdigest()

def save_hash(filename, hash_val):
    with open(filename, "w") as f:
        f.write(hash_val)

def download_file(url, filename):
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

def clean(val):
    """Очищает значение ячейки."""
    if val is None:
        return ""
    s = str(val).strip()
    s = re.sub(r'\s+', ' ', s)
    return s

def extract_data(text):
    """Извлекает из текста пары: предмет, преподавателя и кабинет."""
    if not text:
        return "", "", ""
    
    lines = [l.strip() for l in text.split('\n') if l.strip()]
    
    subject = ""
    teacher = ""
    room = ""
    
    for line in lines:
        # Кабинет ТОЛЬКО в скобках в конце строки (23), (34а)
        room_match = re.search(r'\(([0-9]+[а-яА-ЯёЁ]*)\)\s*$', line)
        if room_match and not room:
            room = room_match.group(1)
            # Удаляем кабинет из строки для дальнейшего парсинга
            line = re.sub(r'\s*\([^)]*\)\s*$', '', line).strip()
        
        # Преподаватель: ФИО (Иванов И.И., Ванявина О.О. и т.д.)
        teacher_match = re.search(r'([А-ЯЁ][а-яё]+\s+[А-ЯЁ]\.\s*[А-ЯЁ]\.?)', line)
        if teacher_match and not teacher:
            teacher = teacher_match.group(1).strip()
            # Удаляем преподавателя из строки
            line = re.sub(r'[А-ЯЁ][а-яё]+\s+[А-ЯЁ]\.\s*[А-ЯЁ]\.?', '', line).strip()
        
        # Предмет: то что осталось (обычно название или код вроде ООД.04)
        if line and not subject:
            # Пропускаем служебные строки
            if not any(x in line.lower() for x in ["разделённая", "объединённая", "физкультура", "консультация"]):
                subject = line
    
    return subject.strip(), teacher.strip(), room.strip()

def parse_schedule_excel():
    """Парсит расписание из Excel."""
    wb = openpyxl.load_workbook(XLSX_FILENAME, data_only=True)
    ws = wb.active
    
    print(f"📄 Лист: {ws.title}, Строк: {ws.max_row}, Столбцов: {ws.max_column}")
    
    # ===== ДАТА =====
    date_str = datetime.now().strftime("%d.%m.%Y")
    for row in range(1, min(5, ws.max_row + 1)):
        for col in range(1, min(10, ws.max_column + 1)):
            val = clean(ws.cell(row, col).value)
            if val:
                match = re.search(r'(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{2,4})', val)
                if match:
                    d, m, y = int(match.group(1)), int(match.group(2)), int(match.group(3))
                    if y < 100:
                        y += 2000
                    try:
                        date_str = datetime(y, m, d).strftime("%d.%m.%Y")
                        print(f"📅 Дата: {date_str}")
                        break
                    except:
                        pass
    
    # ===== ГРУППЫ (строка 2) =====
    groups = []
    group_cols = {}
    for col in range(1, ws.max_column + 1):
        val = clean(ws.cell(2, col).value)
        # Группа: 4 цифры (1161, 1162 и т.д.)
        if val and re.match(r'^\d{4}$', val):
            if val not in groups:
                groups.append(val)
                group_cols[val] = col
    
    print(f"👥 Групп найдено: {len(groups)}")
    if groups:
        print(f"   Группы: {', '.join(groups[:5])}")
    
    if not groups:
        print("❌ Группы не найдены в строке 2!")
        return None
    
    # ===== ПАРЫ (ищем номера пар в столбце A) =====
    schedule = {group: [] for group in groups}
    
    pair_rows = []  # Строки где начинаются пары
    for row in range(3, ws.max_row + 1):
        val = clean(ws.cell(row, 1).value)
        # Проверяем, это ли номер пары (1-7)
        if val and re.match(r'^\d+$', val):
            pair_num = int(val)
            if 1 <= pair_num <= 7:
                pair_rows.append((row, pair_num))
    
    print(f"🔢 Пар найдено: {len(pair_rows)}")
    
    # Парсим каждую пару
    for idx, (start_row, pair_num) in enumerate(pair_rows):
        # Пара занимает 2 строки (или больше до следующей пары)
        if idx + 1 < len(pair_rows):
            end_row = pair_rows[idx + 1][0] - 1
        else:
            end_row = ws.max_row
        
        # Для каждой группы
        for group, col in group_cols.items():
            # Объединяем содержимое ячеек в блоке пары
            content_lines = []
            for row in range(start_row, end_row + 1):
                val = clean(ws.cell(row, col).value)
                if val:
                    content_lines.append(val)
            
            content = '\n'.join(content_lines)
            
            if content:
                subject, teacher, room = extract_data(content)
                
                if subject or teacher or room:
                    lesson = {
                        "num": str(pair_num),
                        "subject": subject,
                        "teacher": teacher,
                        "room": room,
                    }
                    schedule[group].append(lesson)
    
    return {
        "date": date_str,
        "groups": groups,
        "schedule": schedule
    }

def main():
    print("=" * 60)
    print("🎓 Schedule Parser v6.0")
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
        print("⚠️  Скачивание не удалось, используем локальный файл")
        if not os.path.exists(XLSX_FILENAME):
            print("❌ Файл не найден")
            return 1
    
    try:
        result = parse_schedule_excel()
        if not result:
            return 1
        
        with open(JSON_FILENAME, "w", encoding="utf-8") as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        
        print(f"\n✅ Сохранено: {JSON_FILENAME}")
        
        if new_hash:
            save_hash(HASH_FILENAME, new_hash)
        
        # Статистика
        total = sum(len(v) for v in result["schedule"].values())
        print(f"📊 Всего пар: {total}")
        for group, lessons in list(result["schedule"].items())[:5]:
            print(f"   {group}: {len(lessons)} пар")
        
        print("\n✨ Готово!")
        return 0
    
    except Exception as e:
        print(f"❌ Ошибка: {e}")
        import traceback
        traceback.print_exc()
        return 1

if __name__ == "__main__":
    sys.exit(main())
