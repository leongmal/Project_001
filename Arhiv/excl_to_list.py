import pandas as pd

"""Выбираем значения из первого столбца файла Excel и формируем список """ 

# Путь к файлу
ex_file = 'example.xlsx'


# Читаем файл
df = pd.read_excel(ex_file)

# Проверяем, что DataFrame не пустой и есть столбцы
if df.empty:
    print("Файл пуст!")
elif len(df.columns) == 0:
    print("В файле нет столбцов!")
else:
    # Получаем первый столбец и преобразуем в список
    first_column_list = df.iloc[:, 0].tolist()
    print(first_column_list)
