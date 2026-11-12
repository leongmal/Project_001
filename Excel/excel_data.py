"""промпт для поиска в папках в файдах excel
дай код на python для просмотра данных. 
Иходные данные: 
искать по пути \\Esk-node2\ора\РАСПРЕДЕЛЕНИЕ ЭЭ\2026\
заходить в каждую папку и находить в ней файл .xlsx с нечальным словом в имени файла "Распределение ЭЭ"
далее в названии описано название отчетного месяца- получить его значение
отрывать лист "Потребители" и по 16 ой колонке (№ счётчика) найти "8841141132672"

в этойже стоке в 21 ой колонке ("Расход, кВтч") получить величину
итогом выдать: отчетный месяц - величина ("Расход, кВтч")"""

import os
from pathlib import Path

import openpyxl


def parse_month_from_filename(filename: str) -> str:
    """Извлекает название месяца из имени файла после слов 'Распределение ЭЭ'."""
    key = "Распределение ЭЭ"
    if key not in filename:
        return ""
    
    remainder = filename.split(key, 1)[1]
    
    # Убираем разделители в начале остатка строки
    for sep in (" ", "-", "_"):
        if remainder.startswith(sep):
            remainder = remainder[len(sep):]
            break

    months_full = [
        "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
        "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"
    ]
    
    month_lower = remainder.lower()
    for full_name in months_full:
        if full_name.lower() in month_lower:
            return full_name
            
    return "Неизвестный_месяц"


def main():
    base_path = r"\\Esk-node2\ора\РАСПРЕДЕЛЕНИЕ ЭЭ\2026"
    target_meter = "8841141132672"
    
    # Индексы колонок по условию (в Python индексация с 0, поэтому вычитаем 1)
    COL_COUNTER_INDEX = 15  # 16-я колонка (№ счётчика)
    COL_CONSUMPTION_INDEX = 20 # 21-я колонка (Расход, кВтч)
    
    results = []

    print(f"Начат поиск в папке: {base_path}")
    
    # Рекурсивно обходим все подпапки
    for root, dirs, files in os.walk(base_path):
        for file in files:
            if file.endswith(".xlsx") and file.startswith("Распределение ЭЭ"):
                file_path = Path(root) / file
                
                report_month = parse_month_from_filename(file)
                
                try:
                    wb = openpyxl.load_workbook(filename=file_path, read_only=True, data_only=True)
                    
                    if "Потребители" not in wb.sheetnames:
                        print(f"В файле {file} нет листа 'Потребители'. Пропуск.")
                        continue
                        
                    sheet = wb["Потребители"]
                    
                    found = False
                    # iter_rows(min_row=2) — предполагаем, что первая строка это заголовок.
                    # Если данные начинаются сразу с первой строки, замените на min_row=1
                    for row in sheet.iter_rows(min_row=2, values_only=True):
                        # Защита от коротких строк в Excel-файле
                        if len(row) <= max(COL_COUNTER_INDEX, COL_CONSUMPTION_INDEX):
                            continue
                            
                        counter_cell = row[COL_COUNTER_INDEX]
                        
                        # Приводим к строке без пробелов для надежного сравнения
                        if counter_cell and str(counter_cell).replace(" ", "") == target_meter.replace(" ", ""):
                            consumption_cell = row[COL_CONSUMPTION_INDEX]
                            
                            try:
                                consumption_value = float(consumption_cell)
                                results.append({
                                    "month": report_month,
                                    "consumption": consumption_value,
                                    "source_file": file
                                })
                                print(f"[OK] Файл: {file}, Месяц: {report_month}, Расход: {consumption_value}")
                            except (TypeError, ValueError):
                                print(f"[WARN] В файле {file} найден счетчик, но значение расхода ('{consumption_cell}') не является числом.")
                                
                            found = True
                            break  # Счетчик уникален, прекращаем перебор строк этого файла
                    
                    if not found:
                        print(f"[SKIP] В файле {file} счетчик {target_meter} не обнаружен.")

                except Exception as e:
                    print(f"[ERROR] Ошибка при чтении файла {file}: {e}")

    # Итоговый вывод
    print("\n--- ИТОГОВЫЕ ДАННЫЕ ---")
    if not results:
        print("Данные по счетчику не найдены.")
    else:
        # Сортируем по календарному порядку месяцев
        def sort_key(item):
            order = {
                "Январь": 1, "Февраль": 2, "Март": 3, "Апрель": 4, "Май": 5, "Июнь": 6,
                "Июль": 7, "Август": 8, "Сентябрь": 9, "Октябрь": 10, "Ноябрь": 11, "Декабрь": 12
            }
            return order.get(item['month'], 99)

        sorted_results = sorted(results, key=sort_key)
        
        for res in sorted_results:
            print(f"{res['month']} — {res['consumption']}")


if __name__ == "__main__":
    main()