import os
import json
import sys
import subprocess
from docx import Document
from pathlib import Path

# --- 1. Пути ---
TEMPLATE_DIR = Path("template")
RESULTS_DIR = Path("results")
CONFIG_PATH = Path("config.json")

# Создаем папку для результатов
RESULTS_DIR.mkdir(exist_ok=True)

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

# --- 3. Функция замены текста (с сохранением форматирования) ---
def replace_placeholders(doc, data):
    """Ищет плейсхолдеры вида {{ключ}} и заменяет их на значения."""
    for para in doc.paragraphs:
        for key, value in data.items():
            placeholder = f"{{{{{key}}}}}"
            if placeholder in para.text:
                inline = para.runs
                full_text = "".join(run.text for run in inline)
                if placeholder in full_text:
                    new_text = full_text.replace(placeholder, str(value))
                    inline[0].text = new_text
                    for run in inline[1:]:
                        p = run._element
                        p.getparent().remove(p)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                replace_placeholders(cell, data)

# --- 4. Основная обработка ---
# Получаем краткое ФИО для подписи в имени файла
fio_short = data.get("abbreviated_name")
#fio_short = data.get("fioabbr", "ФИО")

# Определяем тип организации
org_type = None
if "ogrnip" in data:
    org_type = "ИП"
elif "ogrn" in data or "inn" in data:
    org_type = "ООО"

if not org_type:
    print("Ошибка: не удалось определить тип организации (ИП или ООО) по данным в config.json.")
    sys.exit(1)

# Словарь соответствия префиксов шаблонов
template_map = {
    "Прил_10": "Прил_10_Шаблон",
    "ДС_ЭДО": f"ДС_ЭДО_{org_type}_Шаблон" if org_type == "ИП" else "ДС_ЭДО_Шаблон"
}

processed_files = []

for prefix, template_prefix in template_map.items():
    # Ищем файл шаблона
    template_candidates = list(TEMPLATE_DIR.glob(f"{template_prefix}*.docx"))
    
    if not template_candidates:
        print(f"Ошибка: не найден шаблон '{template_prefix}' в папке 'template'.")
        continue

    template_path = template_candidates[0]
    
    # Формируем новое имя файла: Прил_10_ИП Умутбаев Д.Р.
    new_filename = f"{prefix}_{org_type} {fio_short}.docx"
    output_path = RESULTS_DIR / new_filename
    
    print(f"Обработка: {template_path.name} -> {new_filename}")
    
    try:
        document = Document(template_path)
        replace_placeholders(document, data)
        document.save(output_path)
        processed_files.append(new_filename)
    except Exception as e:
        print(f"Ошибка при обработке файла {template_path.name}: {e}")

if processed_files:
    print(f"\nГотово! Создано файлов: {len(processed_files)}.")
    for name in processed_files:
        print(f"- {name}")
else:
    print("\nОшибка: не удалось создать ни одного файла.")

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