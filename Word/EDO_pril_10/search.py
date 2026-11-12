import pandas as pd

def search():
    stroka = input('Ищем: ').strip()
    if not stroka:
        print("Вы ничего не ввели. Выход.")
        return

    try:
        df = pd.read_excel(
            r'\\Esk-node2\ора\РАСПРЕДЕЛЕНИЕ ЭЭ\2026\09. Сентябрь\Распределение ЭЭ сентябрь 2026.xlsx', 
            sheet_name='МАКЕТ',
            header=None # Указываем, что заголовков в файле нет
        )
        
        # Столбец H имеет индекс 7 (A=0, B=1, ... H=7)
        # .astype(str) превращаем в строку, .str.strip() убираем пробелы
        mask = df[7].astype(str).str.strip().str.contains(stroka, case=False, na=False)
        
        # Для вывода берем столбцы 7 (Контрагент), 8 (ИНН), 9 (Договор) 
        # (замените индексы 8 и 9 на реальные номера ваших столбцов ИНН и Договора)
        result = df.loc[mask, [7, 5, 6]]
        
        print("\n--- Результаты поиска ---")
        if result.empty:
            print("Ничего не найдено.")
        else:
            # Чтобы вывод был красивым, временно назначим столбцам имена
            result.columns = ['Контрагент', 'ИНН', 'Договор']
            print(result.to_string(index=False))
            
    except FileNotFoundError:
        print("Ошибка: файл не найден. Проверьте сетевой путь.")
    except Exception as e:
        print(f"Произошла ошибка: {e}")

search()