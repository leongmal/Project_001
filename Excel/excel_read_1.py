from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import re

# Функция парсинга строки вида "92100-542ЭА от 01.08.2021Б"
def parse_contract(line):
    if not line:
        return None
    match = re.match(r'(\d{5}-\w+)\s+от\s+(\d{2}\.\d{2}\.\d{4})[БВ]?', str(line).strip())
    if match:
        return {
            'num': match.group(1),
            'date': match.group(2)
        }
    # Если не пошло — возвращаем None
    return None

# ——— Шаг 1: Читаем "образец сверки" — лист "МАКЕТ" — столбцы F, G — номер+дата и организация —
wb = load_workbook('образец сверки.xlsx', data_only=True)
ws = wb['МАКЕТ']

data_list = []  # список словарей: {'num': ..., 'date': ..., 'org': ...}
for row in ws.iter_rows(min_row=2, max_row=222):  # предполагаем, что 1-я строка — заголовок
    cell_f = row[5].value if len(row) > 5 else None  # F = индекс 5 (A=0, B=1, ..., F=5)
    cell_h = row[7].value if len(row) > 6 else None  # H = индекс 7

    if cell_f:
        # Парсим номер и дату из F
        parsed = parse_contract(cell_f)
        if parsed:
            # Организация — из G
            org = str(cell_h).strip() if cell_h else ''
            data_list.append({
                'num': parsed['num'],
                'date': parsed['date'],
                'org': org
            })

# Собираем только номера для сравнения (как раньше)
result_list = [item['num'] for item in data_list]

# ——— Шаг 2: Читаем "Реестр.xlsx" — лист "Потребители КЗ" — столбец D —
wbo = load_workbook('Реестр.xlsx')
wso = wbo['Потребители КЗ']

data_registry_list = []
for row in wso['D']:
    value = row.value
    if value:
        val_str = str(value).strip()
        if val_str:
            data_registry_list.append(val_str)

# ——— Шаг 3: Окраска и подсбор недостающих —
red_fill = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')

# Окраска: тех, что в реестре, но нет в образце
for row in wso['D']:
    value = row.value
    if value:
        val_str = str(value).strip()
        if val_str and val_str not in result_list:
            row.fill = red_fill

# ——— Шаг 4: Находим недостающие (из образца), которые отсутствуют в реестре —
missing_list_with_data = [
    item for item in data_list
    if item['num'] not in data_registry_list
]

# ——— Шаг 5: Добавляем в 3 столбца D (номер), E (дата), B (организация) —
# start_row = wso.max_row + 1
start_row = 182
for i, item in enumerate(missing_list_with_data):
    wso[f'D{start_row + i}'] = item['num']
    wso[f'E{start_row + i}'] = item['date'] if item['date'] else ''
    wso[f'B{start_row + i}'] = item['org'] if item['org'] else ''

# Обновляем метрики
wso.calculate_dimension()

# ——— Шаг 6: Сохраняем в новый файл —
wbo.save('Реестр_проверка.xlsx')
print(f'✅ Добавлено {len(missing_list_with_data)} новых записей.')
print('Файл сохранён: Реестр_проверка.xlsx')

# ——— Отладка: проверка первых 5 записей —
print("\n--- Пример добавленных записей (договор | дата | организация) ---")
for item in missing_list_with_data[:5]:
    print(f"{item['num']} | {item['date']} | {item['org']}")
