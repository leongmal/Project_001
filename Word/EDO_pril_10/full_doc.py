import os
import json
import sys
import sqlite3
import subprocess
from docx import Document
from pathlib import Path

# --- 1. Пути ---
TEMPLATE_DIR = Path("template")
RESULTS_DIR = Path("results")
CONFIG_PATH = Path("config.json")

# Создаем папку для результатов
RESULTS_DIR.mkdir(exist_ok=True)

# --- 2.1. Путь к базе данных ---
DB_PATH = Path("database.db")

# --- 2. Загрузка данных ---
try:
    with open(CONFIG_PATH, "r", encoding="utf-8-sig") as f:
        data = json.load(f)
except FileNotFoundError:
    print(f"Ошибка: файл конфигурации {CONFIG_PATH} не найден.")
    sys.exit(1)
except json.JSONDecodeError as e:
    print(f"Ошибка в JSON файле: {e}")
    sys.exit(1)

# --- 3. Работа с базой данных SQLite ---
def init_db(conn: sqlite3.Connection):
    """Создаёт таблицу, если не существует."""
    conn.execute("""
        CREATE TABLE IF NOT EXISTS config_data (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            inn TEXT UNIQUE NOT NULL,
            ogrnip TEXT,
            ogrn TEXT,
            date TEXT,
            number_contract TEXT,
            date_contract TEXT,
            version_contract TEXT,
            form TEXT,
            abbreviated_name TEXT,
            fio TEXT,
            fioabbr TEXT,
            position TEXT,
            rpposition TEXT,
            rpfio TEXT,
            fioedo TEXT,
            emailedo TEXT,
            telephonedo TEXT,
            operatoredo TEXT,
            A TEXT,
            B TEXT
        )
    """)
    conn.commit()


def save_config_to_db(conn: sqlite3.Connection, data: dict):
    """
    Записывает данные из config.json в БД.
    Если запись с таким ИНН уже существует — пропускает.
    """
    inn = data.get("inn")
    if not inn:
        print("Ошибка: поле 'inn' отсутствует или пустое в config.json.")
        sys.exit(1)

    # Проверяем, есть ли уже запись с таким ИНН
    cursor = conn.execute(
        "SELECT id FROM config_data WHERE inn = ?", (inn,)
    )
    if cursor.fetchone() is not None:
        print(f"Запись с ИНН {inn} уже существует в базе. Пропуск.")
        return

    conn.execute("""
        INSERT INTO config_data (
            inn, ogrnip, ogrn,
            date, number_contract, date_contract,
            version_contract, form,
            abbreviated_name, fio, fioabbr, position,
            rpposition, rpfio, fioedo,
            emailedo, telephonedo, operatoredo,
            A, B
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        inn,
        data.get("ogrnip", ""),
        data.get("ogrn", ""),
        data.get("date", ""),
        data.get("number_contract", ""),
        data.get("date_contract", ""),
        data.get("version_contract", ""),
        data.get("form", ""),
        data.get("abbreviated_name", ""),
        data.get("fio", ""),
        data.get("fioabbr", ""),
        data.get("position", ""),
        data.get("rpposition", ""),
        data.get("rpfio", ""),
        data.get("fioedo", ""),
        data.get("emailedo", ""),
        data.get("telephonedo", ""),
        data.get("operatoredo", ""),
        data.get("A", ""),
        data.get("B", ""),
    ))
    conn.commit()
    print(f"Данные сохранены в базу данных: {DB_PATH}")

# --- 3. Функция замены текста (с сохранением форматирования) ---
def replace_placeholders(doc, data):
    """Ищет плейсхолдеры вида {{ключ}} и заменяет их на значения."""
    # Проходим по всем параграфам
    for para in doc.paragraphs:
        for key, value in data.items():
            placeholder = f"{{{{{key}}}}}"
            if placeholder in para.text:
                # Заменяем текст, сохраняя стили (итерируемся по ранкам)
                inline = para.runs
                full_text = "".join(run.text for run in inline)
                if placeholder in full_text:
                    new_text = full_text.replace(placeholder, str(value))
                    # Перезаписываем текст в первый ран, остальные удаляем
                    inline[0].text = new_text
                    for run in inline[1:]:
                        p = run._element
                        p.getparent().remove(p)

    # Проходим по всем таблицам
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                replace_placeholders(cell, data) # Рекурсивный вызов для ячеек

# --- 4. Основная обработка ---
template_files = list(TEMPLATE_DIR.glob("*.docx"))

if not template_files:
    print(f"Ошибка: в папке '{TEMPLATE_DIR}' не найдено файлов .docx")
    sys.exit(1)

org_name = data.get("abbreviated_name", "Организация")


for template_path in template_files:
    # Формируем новое имя файла
    new_filename = template_path.name.replace("Шаблон", org_name)
    output_path = RESULTS_DIR / new_filename
    
    print(f"Обработка: {template_path.name} -> {new_filename}")
    
    try:
        document = Document(template_path)
        replace_placeholders(document, data)
        document.save(output_path)
    except Exception as e:
        print(f"Ошибка при обработке файла {template_path.name}: {e}")

print("\nГотово! Файлы сохранены в папку 'results'.")

# --- 6. Сохранение данных в БД ---
try:
    conn = sqlite3.connect(DB_PATH)
    init_db(conn)
    save_config_to_db(conn, data)
    conn.close()
except Exception as e:
    print(f"Ошибка при работе с базой данных: {e}")

# --- 5. Открытие папки results ---
try:
    if sys.platform == "win32":
        os.startfile(RESULTS_DIR)
    elif sys.platform == "darwin":
        subprocess.run(["open", RESULTS_DIR])
    else:
        subprocess.run(["xdg-open", RESULTS_DIR])
except Exception as e:
    print(f"Не удалось открыть папку автоматически: {e}")
    print(f"Путь к папке: {RESULTS_DIR.resolve()}")