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
MD5_FILENAME = "schedule.hash"
XLSX_FILENAME = "rasp1.xlsx"
JSON_FILENAME = "schedule.json"

# ===== Вспомогательные функции =====

def get_md5(filename):
    """Вычисляет MD5 файла."""
    if not os.path.exists(filename):
        return None
    with open(filename, "rb") as f:
        return hashlib.md5(f.read()).hexdigest()


def save_md5(filename, md5):
    """Сохраняет MD5 в файл."""
    with open(filename, "w") as f:
        f.write(md5)


def download_file(url, filename):
    """Скачивает файл по URL."""
    print(f"Downloading {url}...")
    try:
        response = requests.get(url, timeout=30)
        response.raise_for_status()
        with open(filename, "wb") as f:
            f.write(response.content)
        print(f"Downloaded: {filename}")
        return True
    except Exception as e:
        print(f"Download error: {e}")
        return False


# ===== Парсер Excel =====

def clean_cell_value(value):
    """Очищает значение ячейки от лишних пробелов и переносов строк."""
    if value is None:
        return ""
    if isinstance(value, (int, float)):
        return str(value).strip()
    return str(value).strip()


def get_merged_value(ws, row, col, merged_values):
    """Возвращает значение ячейки с учётом объединения."""
    for merged_range, top_left_value in merged_values.items():
        if row in range(merged_range.min_row, merged_range.max_row + 1) and \
           col in range(merged_range.min_col, merged_range.max_col + 1):
            return top_left_value
    return None


def extract_date_from_row(ws, row_idx, merged_values):
    """Извлекает дату из строки (обычно строка 1)."""
    for col_idx in range(1, min(ws.max_column + 1, 20)):
        cell = ws.cell(row=row_idx, column=col_idx)
        value = clean_cell_value(cell.value)
        
        if not value:
            merged_val = get_merged_value(ws, row_idx, col_idx, merged_values)
            if merged_val:
                value = clean_cell_value(merged_val)
        
        if value:
            # Ищем дату в формате ДД.МММ или ДД/МММ
            date_match = re.search(r'(\d{1,2})\s*[\./]\s*(\d{1,2})\s*[\./]\s*(\d{2,4})', value)
            if date_match:
                day = int(date_match.group(1))
                month = int(date_match.group(2))
                year = int(date_match.group(3))
                
                if year < 100:
                    year += 2000
                
                try:
                    date_obj = datetime(year, month, day)
                    return date_obj.strftime("%d.%m.%Y")
                except ValueError:
                    pass
    
    return datetime.now().strftime("%d.%m.%Y")


def detect_groups_from_row2(ws, merged_values):
    """Извлекает названия групп из строки 2."""
    groups = []
    group_cols = {}
    
    row_idx = 2
    for col_idx in range(2, ws.max_column + 1):  # Начиная со столбца B (группы начинаются после столба A)
        cell = ws.cell(row=row_idx, column=col_idx)
        value = clean_cell_value(cell.value)
        
        if not value:
            merged_val = get_merged_value(ws, row_idx, col_idx, merged_values)
            if merged_val:
                value = clean_cell_value(merged_val)
        
        # Проверяем на группу (код группы: 4 цифры возможно с буквой)
        if value and re.match(r'^[A-Za-zА-Яа-яЁё]?\d{3,4}[A-Za-zА-Яа-яЁё]?$', value):
            if value not in groups:
                groups.append(value)
                group_cols[value] = col_idx
    
    return groups, group_cols


def find_lesson_rows(ws):
    """Находит строки с номерами пар (в столбце A)."""
    lesson_rows = []
    
    # Ищем в столбце A (первый столбец)
    for row_idx in range(3, ws.max_row + 1):  # Начиная с строки 3 (после даты и групп)
        cell = ws.cell(row=row_idx, column=1)
        value = clean_cell_value(cell.value)
        
        if value and re.match(r'^(\d+)\s*$', value):
            num = int(re.match(r'^(\d+)', value).group(1))
            if 1 <= num <= 7:  # Пары обычно с 1 по 7
                lesson_rows.append((row_idx, num))
    
    return sorted(lesson_rows, key=lambda x: x[0])


def split_by_lines(text):
    """Разбивает текст на строки, учитывая переносы."""
    if not text:
        return []
    lines = re.split(r'[\r\n]+', text)
    return [line.strip() for line in lines if line.strip()]


def parse_teacher_name(text):
    """Парсит имя преподавателя из текста."""
    if not text:
        return None
    
    patterns = [
        r'([А-ЯЁ][а-яё]+\s+[А-ЯЁ]\.\s*[А-ЯЁ]\.?)',
        r'([А-ЯЁ][а-яё]+\s+[А-ЯЁ]\.[А-ЯЁ]\.?)',
        r'([А-ЯЁ]\.\s*[А-ЯЁ]\.?\s+[А-ЯЁ][а-яё]+)',
        r'([А-ЯЁ]\.?[А-ЯЁ]\.?\s*[А-ЯЁ][а-яё]+)',
    ]
    
    for pattern in patterns:
        match = re.search(pattern, text)
        if match:
            return match.group(1).strip()
    
    return None


def parse_room(text):
    """Парсит номер аудитории из текста."""
    if not text:
        return None
    
    patterns = [
        r'\(([^)]+)\)$',
        r'\s+([0-9]+[а-яА-ЯёЁ]?)\s*$',
        r'\s+([0-9]+[а-яА-ЯёЁ]?)\s+',
        r'^([0-9]+[а-яА-ЯёЁ]?)$',
    ]
    
    for pattern in patterns:
        match = re.search(pattern, text)
        if match:
            room = match.group(1).strip()
            if not re.match(r'^\d{1,2}[:.]\d{2}', room):
                return room
    
    return None


def parse_subject_code(text):
    """Парсит код предмета (ООД, ОП, МДК и т.д.)."""
    if not text:
        return None
    
    patterns = [
        r'^([А-ЯЁа-яё]+\.\d+[\.\d]*\s*[\*\s]*)',
        r'^([А-ЯЁа-яё]+\s*\.\s*\d+[\.\d]*)',
        r'^([А-ЯЁа-яё]+\s*\d+[\.\d]*)',
    ]
    
    for pattern in patterns:
        match = re.match(pattern, text)
        if match:
            return match.group(1).strip()
    
    return None


def parse_lesson_cell(value):
    """Парсит ячейку с данными пары."""
    if not value:
        return []
    
    value = clean_cell_value(value)
    lines = split_by_lines(value)
    
    if not lines:
        return []
    
    result = []
    
    for line in lines:
        # Пропускаем служебные строки
        if line.lower() in ["пара", "занятие", "физ-ра", "физкультура", "консультация", 
                            "разделённая пара", "объединённая пара"]:
            continue
        
        if re.match(r'^\d{1,2}:\d{2}', line):
            continue
        
        text = line
        
        # Парсим аудиторию
        room = parse_room(text)
        if room:
            for pattern in [
                r'\([^)]+\)$',
                r'\s+[0-9]+[а-яА-ЯёЁ]?\s*$',
                r'\s+[0-9]+[а-яА-ЯёЁ]?\s+',
                r'^[0-9]+[а-яА-ЯёЁ]?$'
            ]:
                text = re.sub(pattern, '', text).strip()
                if room not in text:
                    break
        
        # Парсим преподавателя
        teacher = parse_teacher_name(text)
        if teacher:
            for pattern in [
                r'[А-ЯЁ][а-яё]+\s+[А-ЯЁ]\.\s*[А-ЯЁ]\.?',
                r'[А-ЯЁ][а-яё]+\s+[А-ЯЁ]\.[А-ЯЁ]\.?',
                r'[А-ЯЁ]\.\s*[А-ЯЁ]\.?\s+[А-ЯЁ][а-яё]+',
                r'[А-ЯЁ]\.?[А-ЯЁ]\.?\s*[А-ЯЁ][а-яё]+',
            ]:
                text = re.sub(pattern, '', text).strip()
                if teacher not in text:
                    break
        
        # Остаток — предмет
        subject = text.strip()
        subject = re.sub(r'^\s*[–—-]\s*', '', subject)
        subject = re.sub(r'\s*[–—-]\s*$', '', subject)
        subject = re.sub(r'^\s*[.,;]\s*', '', subject)
        subject = re.sub(r'\s*[.,;]\s*$', '', subject)
        
        # Парсим код предмета
        code = parse_subject_code(subject)
        if code:
            subject = re.sub(r'^' + re.escape(code), '', subject).strip()
            if subject:
                subject = f"{code} {subject}"
            else:
                subject = code
        
        subject = re.sub(r'\s*\([^)]*\)\s*', ' ', subject)
        subject = ' '.join(subject.split())
        
        if subject or teacher or room:
            result.append({
                "subject": subject if subject else "",
                "teacher": teacher if teacher else "",
                "room": room if room else "",
            })
    
    return result


def merge_items(items):
    """Объединяет дублирующиеся предметы/преподавателей/аудитории."""
    unique = []
    seen = set()
    
    for item in items:
        key = (
            item.get("subject", ""),
            item.get("teacher", ""),
            item.get("room", ""),
        )
        
        if not any(key):
            continue
        
        if key in seen:
            continue
        
        seen.add(key)
        unique.append(item)
    
    return unique


def make_lesson(number, items):
    """Создаёт объект занятия из списка элементов."""
    items = merge_items(items)
    
    if not items:
        return None
    
    subjects = []
    teachers = []
    rooms = []
    
    for item in items:
        if item["subject"] and item["subject"] not in subjects:
            subjects.append(item["subject"])
        
        if item["teacher"] and item["teacher"] not in teachers:
            teachers.append(item["teacher"])
        
        if item["room"] and item["room"] not in rooms:
            rooms.append(item["room"])
    
    return {
        "num": str(number),
        "subject": " / ".join(subjects) if subjects else "",
        "teacher": " / ".join(teachers) if teachers else "",
        "room": " / ".join(rooms) if rooms else "",
        "items": items,
    }


def parse_schedule(ws, groups, group_cols, lesson_rows, merged_values):
    """Основная функция парсинга расписания."""
    schedule = {group: [] for group in groups}
    
    for index, (start_row, number) in enumerate(lesson_rows):
        # Определяем конец блока (до следующей пары или конец листа)
        if index + 1 < len(lesson_rows):
            end_row = lesson_rows[index + 1][0] - 1
        else:
            end_row = ws.max_row
        
        # Для каждой группы парсим ячейки с этой парой
        for group, col_idx in group_cols.items():
            items = []
            
            # Проходим по строкам блока
            for row_idx in range(start_row, end_row + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                value = clean_cell_value(cell.value)
                
                if not value:
                    merged_val = get_merged_value(ws, row_idx, col_idx, merged_values)
                    if merged_val:
                        value = clean_cell_value(merged_val)
                
                if value:
                    parsed_items = parse_lesson_cell(value)
                    items.extend(parsed_items)
            
            # Создаём занятие
            lesson = make_lesson(number, items)
            if lesson:
                schedule[group].append(lesson)
    
    return schedule


def main():
    """Основная функция."""
    print("=" * 60)
    print("Schedule Parser v2.0")
    print("=" * 60)
    
    # Проверяем MD5
    old_md5 = get_md5(XLSX_FILENAME)
    new_md5 = None
    
    # Скачиваем файл
    if download_file(URL, XLSX_FILENAME):
        new_md5 = get_md5(XLSX_FILENAME)
        
        if old_md5 == new_md5:
            print("File not changed, using existing schedule.json")
            return 0
    else:
        print("Download failed, using existing schedule.json")
        if not os.path.exists(JSON_FILENAME):
            print("Error: schedule.json not found")
            return 1
        return 0
    
    try:
        # Загружаем Excel
        print(f"Loading {XLSX_FILENAME}...")
        wb = openpyxl.load_workbook(XLSX_FILENAME, data_only=True)
        ws = wb.active
        
        print(f"Sheet: {ws.title}, Rows: {ws.max_row}, Cols: {ws.max_column}")
        
        # Собираем информацию об объединённых ячейках
        merged_values = {}
        for merged_range in ws.merged_cells.ranges:
            top_left = ws.cell(row=merged_range.min_row, column=merged_range.min_col)
            if top_left.value is not None:
                merged_values[merged_range] = clean_cell_value(top_left.value)
        
        # Извлекаем дату из строки 1
        date_str = extract_date_from_row(ws, 1, merged_values)
        print(f"Date: {date_str}")
        
        # Определяем группы из строки 2
        groups, group_cols = detect_groups_from_row2(ws, merged_values)
        
        if not groups:
            print("Error: Groups not found in row 2")
            return 1
        
        print(f"Found {len(groups)} groups: {', '.join(groups)}")
        
        # Находим номера пар (строки со значениями в столбце A)
        lesson_rows = find_lesson_rows(ws)
        
        if not lesson_rows:
            print("Error: Lesson rows not found in column A")
            return 1
        
        print(f"Found {len(lesson_rows)} lessons at rows: {[lr[0] for lr in lesson_rows]}")
        
        # Парсим расписание
        schedule = parse_schedule(ws, groups, group_cols, lesson_rows, merged_values)
        
        # Формируем результат
        result = {
            "date": date_str,
            "groups": groups,
            "schedule": schedule,
        }
        
        # Сохраняем JSON
        with open(JSON_FILENAME, "w", encoding="utf-8") as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        
        print(f"Saved: {JSON_FILENAME}")
        
        if new_md5:
            save_md5(MD5_FILENAME, new_md5)
        
        # Выводим статистику
        total_lessons = sum(len(lessons) for lessons in schedule.values())
        print(f"\n📊 Статистика:")
        print(f"   Групп: {len(groups)}")
        print(f"   Всего пар: {total_lessons}")
        print(f"   Средне на группу: {total_lessons // len(groups) if groups else 0}")
        
        print("\n✅ Done!")
        return 0
    
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(main())
