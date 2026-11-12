import os
import time

# Базовый сетевой путь (не меняется)
base_path = r"\\Esk-node2\ора\Договора Стачек\Договора арендаторы"

def find_and_open_folder(path, fragment):
    if not os.path.exists(path):
        print(f"Ошибка: Сетевой путь {path} недоступен. Проверьте правописание и права доступа.")
        input("Нажмите Enter, чтобы выйти...")
        return

    print(f"Поиск организаций по фрагменту: '{fragment}'...")
    
    # 1. Ищем папки организаций (рекурсивно)
    org_matches = []
    for root, dirs, files in os.walk(path):
        for d in dirs:
            if fragment.lower() in d.lower():
                org_matches.append(os.path.join(root, d))

    if not org_matches:
        print(f"\nПапки, содержащие '{fragment}', не найдены.")
        input("Нажмите Enter, чтобы выйти...")
        return

    # 2. Берем последнюю измененную организацию
    latest_org = max(org_matches, key=os.path.getmtime)
    print(f"Найдена целевая организация: {latest_org}")

    # 3. Ищем самую свежую папку внутри неё
    try:
        items = os.listdir(latest_org)
        subfolders = [os.path.join(latest_org, item) for item in items if os.path.isdir(os.path.join(latest_org, item))]

        if not subfolders:
            print("Внутри папки организации нет вложенных папок. Открываю папку самой организации.")
            # Если внутри нет папок, откроем саму папку организации
            target_path = latest_org
        else:
            # Выбираем самую свежую папку
            latest_subfolder = max(subfolders, key=os.path.getmtime)
            target_path = latest_subfolder
        
        # Получаем время изменения для вывода
        mod_time = os.path.getmtime(target_path)
        
        print(f"\nНайдена последняя папка: {target_path}")
        print(f"Дата изменения: {time.strftime('%d.%m.%Y %H:%M:%S', time.localtime(mod_time))}")
        
        # 4. Открываем папку в проводнике Windows
        print("Открываю папку...")
        os.startfile(target_path)
        
    except PermissionError:
        print(f"\nОшибка: Нет прав на чтение содержимого папки {latest_org}")
        input("Нажмите Enter, чтобы выйти...")
    except Exception as e:
        print(f"\nПроизошла непредвиденная ошибка: {e}")
        input("Нажмите Enter, чтобы выйти...")

# --- Основная часть программы ---
if __name__ == "__main__":
    # Очищаем экран для удобства (работает в Windows)
    os.system('cls' if os.name == 'nt' else 'clear')
    
    # Спрашиваем название у пользователя
    org_name = input("Введите фрагмент названия организации для поиска: ")
    
    if org_name.strip():
        find_and_open_folder(base_path, org_name.strip())
    else:
        print("Имя не введено. Выход.")