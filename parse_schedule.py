#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import re
import json
import hashlib
import requests
import openpyxl
from openpyxl.utils import get_column_letter
from datetime import datetime
import sys

# ===== Конфигурация =====
URL = "https://serp-koll.ru/images/ep/k1/rasp1.xlsx"
MD5_FILENAME = "rasp1.xlsx.md5"
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
    """
    Возвращает значение ячейки с учётом объединения.
    Если ячейка объединена — берёт значение из верхней левой ячейки.
    """
    for merged_range, top_left_value in merged_values.items():
        if row in range(merged_range.min_row, merged_range.max_row + 1) and \
           col in range(merged_range.min_col, merged_range.max_col + 1):
            return top_left_value
    return None


def detect_groups(ws, merged_values):
    """
    Находит все группы в строке 2 (или в первой строке с группами).
    Возвращает: список групп и словарь {группа: номер_колонки}.
    """
    groups = []
    group_cols = {}

    # Ищем группы во 2-й строке
    row_idx = 2
    for col_idx in range(1, ws.max_column + 1):
        cell = ws.cell(row=row_idx, column=col_idx)
        value = clean_cell_value(cell.value)

        # Если ячейка объединена, берём значение из объединения
        if not value:
            merged_val = get_merged_value(ws, row_idx, col_idx, merged_values)
            if merged_val:
                value = clean_cell_value(merged_val)

        # Проверяем, похоже ли значение на группу (цифры или буквы/цифры)
        if value and re.match(r'^[А-Яа-яЁё]?\d{3,4}[а-яА-Я]?$', value):
            if value not in groups:
                groups.append(value)
                group_cols[value] = col_idx

    # Если групп не найдено, пробуем строку 3
    if not groups:
        row_idx = 3
        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            value = clean_cell_value(cell.value)

            if not value:
                merged_val = get_merged_value(ws, row_idx, col_idx, merged_values)
                if merged_val:
                    value = clean_cell_value(merged_val)

            if value and re.match(r'^[А-Яа-яЁё]?\d{3,4}[а-яА-Я]?$', value):
                if value not in groups:
                    groups.append(value)
                    group_cols[value] = col_idx

    return groups, group_cols


def detect_lessons(ws):
    """
    Находит строки, в которых начинаются пары.
    Возвращает список кортежей (номер_строки, номер_пары).
    """
    lesson_rows = []

    # Ищем строки с номерами пар (1, 2, 3, 4, 5, 6, 7)
    for row_idx in range(1, ws.max_row + 1):
        cell = ws.cell(row=row_idx, column=1)
        value = clean_cell_value(cell.value)

        if not value:
            continue

        # Проверяем, является ли значение номером пары
        match = re.match(r'^(\d+)\s*$', value)
        if match:
            num = int(match.group(1))
            if 1 <= num <= 7:
                lesson_rows.append((row_idx, num))

    # Если не нашли по первой колонке, пробуем по второй
    if not lesson_rows:
        for row_idx in range(1, ws.max_row + 1):
            cell = ws.cell(row=row_idx, column=2)
            value = clean_cell_value(cell.value)

            if not value:
                continue

            match = re.match(r'^(\d+)\s*$', value)
            if match:
                num = int(match.group(1))
                if 1 <= num <= 7:
                    lesson_rows.append((row_idx, num))

    return sorted(lesson_rows, key=lambda x: x[0])


def parse_lesson_cell(value):
    """
    Парсит ячейку с данными пары.
    Возвращает: subject, teacher, room.
    """
    if not value:
        return "", "", ""

    value = clean_cell_value(value)

    # Пытаемся найти аудиторию
    room = ""
    room_patterns = [
        r'\(([^)]+)\)$',          # (305), (305а), (23)
        r'\s+([0-9]+[а-яА-Я]?)\s*$',  # 305, 305а
        r'\s+([0-9]+[а-яА-Я]?)\s+',   # 305 в середине
    ]

    for pattern in room_patterns:
        match = re.search(pattern, value)
        if match:
            room = match.group(1).strip()
            # Удаляем аудиторию из значения
            value = value[:match.start()] + value[match.end():]
            break

    # Теперь ищем преподавателя
    teacher = ""
    teacher_patterns = [
        r'([А-Я][а-я]+\s+[А-Я]\.\s*[А-Я]\.?)',  # Иванов И.И.
        r'([А-Я][а-я]+\s+[А-Я]\.[А-Я]\.?)',      # Иванов И.И. (без пробела)
        r'([А-Я]\.\s*[А-Я]\.?\s+[А-Я][а-я]+)',   # И.И. Иванов
        r'([А-Я]\.?[А-Я]\.?\s+[А-Я][а-я]+)',     # И.И.Иванов
    ]

    for pattern in teacher_patterns:
        match = re.search(pattern, value)
        if match:
            teacher = match.group(1).strip()
            # Удаляем преподавателя из значения
            value = value[:match.start()] + value[match.end():]
            break

    # Остаток — это предмет
    subject = clean_cell_value(value)

    # Убираем лишние символы
    subject = re.sub(r'^\s*[–—-]\s*', '', subject)
    subject = re.sub(r'\s*[–—-]\s*$', '', subject)
    subject = re.sub(r'^\s*[.,;]\s*', '', subject)
    subject = re.sub(r'\s*[.,;]\s*$', '', subject)

    # Если предмет начинается с кода (ОГСЭ.04, ОП.09*, МДК и т.д.)
    # оставляем код + название
    code_match = re.match(r'^([А-ЯЁа-яё]+\s*[\.\*]?\s*\d+[\.\d]*\s*[\.\*]?)\s*(.*)$', subject)
    if code_match:
        code = code_match.group(1).strip()
        name = code_match.group(2).strip()
        if name:
            subject = f"{code} {name}"
        else:
            subject = code

    # Если в предмете остались скобки — убираем
    subject = re.sub(r'\s*\([^)]*\)\s*', ' ', subject)
    subject = ' '.join(subject.split())

    return subject, teacher, room


def parse_lesson_block(ws, start_row, end_row, group_cols, merged_values):
    """
    Парсит блок строк, содержащих одну пару.
    Возвращает словарь {группа: [items]}.
    """
    items_by_group = {group: [] for group in group_cols.keys()}

    # Собираем все данные из блока строк
    for row_idx in range(start_row, end_row + 1):
        for group, col_idx in group_cols.items():
            cell = ws.cell(row=row_idx, column=col_idx)
            value = clean_cell_value(cell.value)

            # Проверяем объединённую ячейку
            if not value:
                merged_val = get_merged_value(ws, row_idx, col_idx, merged_values)
                if merged_val:
                    value = clean_cell_value(merged_val)

            if not value:
                continue

            # Проверяем, не является ли значение служебным (например, "пара")
            if value.lower() in ["пара", "занятие", "физ-ра", "физкультура", "консультация"]:
                continue

            # Проверяем, не является ли значение временем или датой
            if re.match(r'^\d{1,2}:\d{2}', value):
                continue

            # Парсим значение
            subject, teacher, room = parse_lesson_cell(value)

            if subject or teacher or room:
                items_by_group[group].append({
                    "subject": subject,
                    "teacher": teacher,
                    "room": room,
                })

    return items_by_group


def find_lesson_rows(ws):
    """
    Находит строки, с которых начинаются пары.
    """
    lesson_rows = []

    for row_idx in range(1, ws.max_row + 1):
        cell = ws.cell(row=row_idx, column=1)
        value = clean_cell_value(cell.value)

        if value and re.match(r'^(\d+)\s*$', value):
            num = int(re.match(r'^(\d+)', value).group(1))
            if 1 <= num <= 7:
                lesson_rows.append((row_idx, num))

    # Если не нашли по колонке 1, пробуем колонку 2
    if not lesson_rows:
        for row_idx in range(1, ws.max_row + 1):
            cell = ws.cell(row=row_idx, column=2)
            value = clean_cell_value(cell.value)

            if value and re.match(r'^(\d+)\s*$', value):
                num = int(re.match(r'^(\d+)', value).group(1))
                if 1 <= num <= 7:
                    lesson_rows.append((row_idx, num))

    return sorted(lesson_rows, key=lambda x: x[0])


def merge_items(items):
    """
    Объединяет одинаковые предметы/преподавателей/аудитории.
    """
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
    """
    Создаёт объект занятия из списка элементов.
    """
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
        "subject": " / ".join(subjects),
        "teacher": " / ".join(teachers),
        "room": " / ".join(rooms),
        "items": items,
    }


def parse_schedule(ws, groups, group_cols, merged_values):
    """
    Основная функция парсинга расписания.
    """
    lesson_rows = find_lesson_rows(ws)

    if not lesson_rows:
        raise ValueError("Lesson rows not found in XLSX")

    print("Lesson rows:", lesson_rows)

    schedule = {group: [] for group in groups}

    for index, (start_row, number) in enumerate(lesson_rows):
        # Определяем конец блока
        if index + 1 < len(lesson_rows):
            end_row = lesson_rows[index + 1][0] - 1
        else:
            # Последняя пара — до конца листа
            end_row = ws.max_row

        # Парсим блок
        items_by_group = parse_lesson_block(ws, start_row, end_row, group_cols, merged_values)

        # Создаём занятия для каждой группы
        for group in groups:
            items = items_by_group.get(group, [])
            lesson = make_lesson(number, items)

            if lesson:
                schedule[group].append(lesson)

    return schedule


def get_date_from_ws(ws):
    """
    Пытается найти дату в листе Excel.
    """
    # Сначала ищем в строке 1
    for col_idx in range(1, min(ws.max_column + 1, 10)):
        cell = ws.cell(row=1, column=col_idx)
        value = clean_cell_value(cell.value)

        if value:
            # Проверяем, похоже ли значение на дату
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

    # Если не нашли, пробуем другие строки
    for row_idx in range(1, 5):
        for col_idx in range(1, min(ws.max_column + 1, 5)):
            cell = ws.cell(row=row_idx, column=col_idx)
            value = clean_cell_value(cell.value)

            if value:
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

    # Если дату не нашли, используем текущую дату
    return datetime.now().strftime("%d.%m.%Y")


def main():
    """Основная функция."""
    print("=" * 60)
    print("Schedule Parser")
    print("=" * 60)

    # Проверяем MD5
    old_md5 = get_md5(XLSX_FILENAME)
    new_md5 = None

    # Скачиваем файл
    if download_file(URL, XLSX_FILENAME):
        new_md5 = get_md5(XLSX_FILENAME)

        # Проверяем, изменился ли файл
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

        # Собираем информацию об объединённых ячейках
        merged_values = {}
        for merged_range in ws.merged_cells.ranges:
            top_left = ws.cell(row=merged_range.min_row, column=merged_range.min_col)
            if top_left.value is not None:
                merged_values[merged_range] = clean_cell_value(top_left.value)

        # Определяем группы
        groups, group_cols = detect_groups(ws, merged_values)

        if not groups:
            print("Error: Groups not found")
            return 1

        print(f"Found groups: {', '.join(groups)}")

        # Определяем дату
        date_str = get_date_from_ws(ws)
        print(f"Date: {date_str}")

        # Парсим расписание
        schedule = parse_schedule(ws, groups, group_cols, merged_values)

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

        # Сохраняем MD5
        if new_md5:
            save_md5(MD5_FILENAME, new_md5)

        print("Done!")

        # Выводим статистику
        total_lessons = 0
        for group, lessons in schedule.items():
            total_lessons += len(lessons)

        print(f"Total groups: {len(groups)}")
        print(f"Total lessons: {total_lessons}")

        return 0

    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(main())
