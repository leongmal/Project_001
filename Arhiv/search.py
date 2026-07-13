# import pandas as pd
""" поиск в IDLE по фрагменту из столбца Контрагент, вызов фнкции со значением.."""
# def search(stroka):
    # df = pd.read_excel(r'\\Esk-node2\ора\РАСПРЕДЕЛЕНИЕ ЭЭ\2026\04. Апрель\Распределение ЭЭ апрель 2026.xlsx', sheet_name = 'РП')
    # pred_result = df['№ счётчика'].astype(str).str.strip().str.contains(str(stroka).strip(), case=False, na=False)
    # result = df.loc[pred_result,['Точка учётиа','№ счётчика','Расход, кВтч']] 
 
 
    
# import pandas as pd
# """ поиск значения Квч*ч  по номеру счетчика """
# def search(stroka):
    # file_path = r'\\Esk-node2\ора\РАСПРЕДЕЛЕНИЕ ЭЭ\2026\04. Апрель\Распределение ЭЭ апрель 2026.xlsx'
    # try:
        # df = pd.read_excel(file_path, sheet_name='РП')        
        # # Ваш исходный поиск
        # pred_result = df['№ счётчика'].astype(str).str.contains(str(stroka).strip(), case=False, na=False)
        # first_col_name = df.columns[0]
        # columns_to_show = [first_col_name, 'Точка учётиа', '№ счётчика', 'Расход, кВтч']
        # result = df.loc[pred_result, columns_to_show]
        
        # if result.empty:
            # print("\nСовпадений не найдено.")
        # else:
            # print(result)           
    # except Exception as e:
        # print(f"Ошибка при чтении файла: {e}")

# Вызов функции
# search('50937766')




# import pandas as pd
# """ поиск в IDLE по фрагменту из столбца Контрагент, вызов фнкции со значением.."""

# def search():
    # stroka = input('ищем :  ').strip()
    # df = pd.read_excel(r'C:\Users\Leonik\Desktop\Конверт\Потр_email.xlsx', sheet_name = 'МАКЕТ')
    # pred_result = df['Контрагент'].astype(str).str.strip().str.contains(stroka, case=False, na=False)
    # result = df.loc[pred_result,['Контрагент','ЭЛ.ПОЧТА']] 
    
    
# search()

import pandas as pd

def search():
    # 1. Запрашиваем у пользователя строку для поиска
    stroka = input('Ищем: ').strip()
    
    if not stroka:
        print("Вы ничего не ввели. Выход.")
        return # Выходим из функции, если поиск пустой

    try:
        # 2. Читаем файл Excel
        # Обратите внимание: путь к файлу должен существовать.
        df = pd.read_excel(r'C:\Users\Leonik\Desktop\Конверт\Потр_email.xlsx', sheet_name='МАКЕТ')
        
        # 3. Подготовка данных и поиск
        # .str.strip() убирает лишние пробелы в начале и конце каждой ячейки.
        # .str.contains() ищет подстроку (без учета регистра).
        mask = df['Контрагент'].astype(str).str.strip().str.contains(stroka, case=False, na=False)
        
        # 4. Фильтрация и вывод результата
        result = df.loc[mask, ['Контрагент', 'ЭЛ.ПОЧТА']]
        
        print("\n--- Результаты поиска ---")
        if result.empty:
            print("Ничего не найдено.")
        else:
            # Красиво выводим результат без индекса
            print(result.to_string(index=False))
            
    except FileNotFoundError:
        print(f"Ошибка: Файл по пути {r'C:\Users\Leonik\Desktop\Конверт\Потр_email.xlsx'} не найден.")
    except Exception as e:
        print(f"Произошла непредвиденная ошибка: {e}")

# Вызываем функцию
search()