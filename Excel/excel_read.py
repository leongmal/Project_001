from openpyxl import load_workbook
from openpyxl.styles import PatternFill

''''workbook - для создания , load_workbook для открытия'''

''''создание'''
wb = load_workbook('образец сверки.xlsx', data_only=True)
ws = wb['МАКЕТ']


data_list = []
for row in ws['F']:
    value = row.value
    if value:
        cleaned = str(value).strip()
        data_list.append(cleaned)
### Ограничиваем количество строк (отметаем лишние)
MAX_ROWS = 222
data_list = data_list[:MAX_ROWS]

### Сначала собираем в список
result_list = []
for i in data_list:
    result_list.append(i.split()[0])
    # print(i.split()[0])

wbo = load_workbook('Реестр.xlsx')
wso = wbo['Потребители КЗ']

data_registry_list=[]
for row in wso['D']:
    value = row.value
    if value is not None:
        val_str = str(value).strip()
        if val_str:
             data_registry_list.append(val_str)


# окрашиваем в красный цвет заливку ячеки
red_fill= PatternFill(start_color='FF0000', end_color ='Ff0000' , fill_type="solid")

for row in wso['D']:
    value =row.value
    if value is not None:
        val_str = str(value).strip()
        if val_str not in result_list:
            row.fill = red_fill


missing_list=[]
for num in result_list:
    if num not in data_registry_list:
        if '92100' in num:
            missing_list.append(num)

# добавляем в конец,  182 это пустая строка в excel
start_row = 182
for i, num in enumerate(missing_list):
    wso[f'D{start_row + i}'] = num

wso.calculate_dimension()

wbo.save('Реестр_проверка.xlsx')
print(f'Добавлено {len(missing_list)} новых записей. Файл сохранён!')
# # print(data_registry_list)
# print(f"result_list (first 5): {result_list[:5]}")
# print(f"data_registry_list: {data_registry_list[:5]}")
# print(f"missing_list: {missing_list[:10]}")
# print("\n--- Последние 5 строк в листе после добавления ---")
# for row in wso.iter_rows(min_row=start_row + len(missing_list) - 5, max_row=start_row + len(missing_list) - 1):
#     for cell in row:
#         print(f"{cell.coordinate} = {cell.value}")
