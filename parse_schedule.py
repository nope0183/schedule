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
    # Убираем переносы строк
    s = re.sub(r'\s+', ' ', s)
    return s

def analyze_structure():
    """Анализирует структуру Excel файла."""
    if not os.path.exists(XLSX_FILENAME):
        print("❌ Файл не найден")
        return
    
    wb = openpyxl.load_workbook(XLSX_FILENAME, data_only=True)
    ws = wb.active
    
    print(f"\n📊 АНАЛИЗ СТРУКТУРЫ")
    print(f"Лист: {ws.title}")
    print(f"Строк: {ws.max_row}, Столбцов: {ws.max_column}\n")
    
    # Показываем первые 10 строк
    print("Первые 10 строк:")
    for row in range(1, min(11, ws.max_row + 1)):
        row_data = []
        for col in range(1, min(6, ws.max_column + 1)):
            val = clean(ws.cell(row, col).value)
            row_data.append(f"[{val[:15]}]" if val else "[_]")
        print(f"Строка {row}: {' '.join(row_data)}")
    
    print("\n")

def parse_excel():
    """Парсит расписание из Excel."""
    wb = openpyxl.load_workbook(XLSX_FILENAME, data_only=True)
    ws = wb.active
    
    print(f"📄 Лист: {ws.title}")
    
    # ===== ДАТА =====
    date_str = datetime.now().strftime("%d.%m.%Y")
    for col in range(1, min(10, ws.max_column + 1)):
        val = clean(ws.cell(1, col).value)
        if val:
            match = re.search(r'(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{2,4})', val)
            if match:
                d, m, y = int(match.group(1)), int(match.group(2)), int(match.group(3))
                if y < 100:
                    y += 2000
                try:
                    date_str = datetime(y, m, d).strftime("%d.%m.%Y")
                    break
                except:
                    pass
    print(f"📅 Дата: {date_str}")
    
    # ===== ГРУППЫ (строка 2) =====
    groups = []
    group_cols = {}
    for col in range(1, ws.max_column + 1):
        val = clean(ws.cell(2, col).value)
        # Группа: до 5 символов, содержит цифры
        if val and re.match(r'^[А-Яа-я0-9]{2,5}$', val) and any(c.isdigit() for c in val):
            if val not in groups:
                groups.append(val)
                group_cols[val] = col
    
    print(f"👥 Групп: {len(groups)}")
    if groups:
        print(f"   {', '.join(groups[:10])}")
    
    if not groups:
        print("❌ Группы не найдены!")
        return None
    
    # ===== ПАРЫ (начиная со строки 3) =====
    schedule = {group: [] for group in groups}
    
    row = 3
    while row <= ws.max_row:
        # Проверяем, есть ли номер пары в столбце A
        col_a = clean(ws.cell(row, 1).value)
        
        # Номер пары - это цифра от 1 до 7
        pair_num = None
        if col_a and re.match(r'^\d+$', col_a):
            pair_num = int(col_a)
            if not (1 <= pair_num <= 7):
                pair_num = None
        
        if pair_num is not None:
            # Это строка с номером пары
            # Данные пары находятся в этой же строке для каждой группы
            print(f"Пара {pair_num} на строке {row}")
            
            for group, col in group_cols.items():
                val = clean(ws.cell(row, col).value)
                if val:
                    lesson = {
                        "num": str(pair_num),
                        "subject": val,
                        "teacher": "",
                        "room": "",
                    }
                    schedule[group].append(lesson)
        
        row += 1
    
    return {
        "date": date_str,
        "groups": groups,
        "schedule": schedule
    }

def main():
    print("=" * 60)
    print("🎓 Schedule Parser v4.0")
    print("=" * 60)
    
    # Анализируем структуру
    analyze_structure()
    
    # Проверяем изменения
    old_hash = get_hash(XLSX_FILENAME)
    new_hash = None
    
    if download_file(URL, XLSX_FILENAME):
        new_hash = get_hash(XLSX_FILENAME)
        if old_hash == new_hash:
            print("📌 Файл не изменился")
            return 0
    else:
        print("⚠️  Скачивание не удалось")
        if not os.path.exists(JSON_FILENAME):
            return 1
    
    try:
        result = parse_excel()
        if not result:
            return 1
        
        with open(JSON_FILENAME, "w", encoding="utf-8") as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        
        print(f"\n✅ Сохранено: {JSON_FILENAME}")
        
        if new_hash:
            save_hash(HASH_FILENAME, new_hash)
        
        total = sum(len(v) for v in result["schedule"].values())
        print(f"📊 Пар всего: {total}")
        
        return 0
    
    except Exception as e:
        print(f"❌ Ошибка: {e}")
        import traceback
        traceback.print_exc()
        return 1

if __name__ == "__main__":
    sys.exit(main())
