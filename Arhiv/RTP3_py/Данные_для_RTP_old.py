import pandas as pd
import json

from pathlib import Path
# Параметры
#excel_file = '/Users/Leonik/AppData/Local/Programs/Python/Python312/RTP3_py/Распределение ЭЭ декабрь 2025.xlsx'
excel_file=Path(r'\\Esk-node2\ора\РАСПРЕДЕЛЕНИЕ ЭЭ\2026\01. Январь\Распределение ЭЭ январь 2026.xlsx')
column_numbers = '№ счётчика'
column_values = 'Расход, кВтч'
search_list_file = '/Users/Leonik/AppData/Local/Programs/Python/Python312/RTP3_py/serial_numbers.txt'

def read_search_list_from_file(file_path):
    try:
        with open(file_path, 'r', encoding='UTF-8') as file:
            # Предполагаем, что в файле каждое число на новой строке
            search_list = [line.strip() for line in file if line.strip()]
            # Преобразуем строки в числа
            search_list = [int(num) for num in search_list]
        return search_list
    except Exception as e:
        print(f"Ошибка при чтении файла search_list: {e}")
        return []

def create_number_dict(excel_file, column_numbers, column_values, search_list):
    # Читаем Excel-файл
    df = pd.read_excel(excel_file, sheet_name='РП')
    result_dict = {}

    # Проверяем наличие столбцов
    if column_numbers not in df.columns:
        raise ValueError(f"Столбец '{column_numbers}' не найден в файле.")
    if column_values not in df.columns:
        raise ValueError(f"Столбец '{column_values}' не найден в файле.")

    # Преобразуем search_list в строки для корректного сравнения
    search_list_str = list(map(str, search_list))

    # Итерируем по строкам DataFrame
    for index, row in df.iterrows():
        current_number = row[column_numbers]

        # Пропускаем пропущенные значения (NaN, None и т.п.)
        if pd.isna(current_number):
            continue

        # Приводим к строке и убираем пробелы по краям
        current_number = str(current_number).strip()

        # Пропускаем значения, которые превратились в 'nan' после str()
        if current_number.lower() == 'nan':
            continue

        # Проверяем, есть ли номер в списке поиска (сравнение строк)
        if current_number in search_list_str:
            value = row[column_values]

            # Если значение расхода не пропущено — добавляем в словарь
            if pd.notna(value):
                result_dict[current_number] = int(value)
            else:
                # Если расход пропущен — записываем None
                result_dict[current_number] = None

    return result_dict

# Читаем search_list из файла
search_list = read_search_list_from_file(search_list_file)

try:
    result = create_number_dict(excel_file, column_numbers, column_values, search_list)
    
    print("Результат:")
    print(result)
    
    with open('/Users/Leonik/AppData/Local/Programs/Python/Python312/RTP3_py/num_val.txt', 'w', encoding='UTF-8') as file:
        json.dump(result, file, ensure_ascii=False, indent=4)
        
except Exception as e:
    print(f"Ошибка: {e}")


