import os
import fnmatch

def search_in_folders(root_folder, search_query):
    """
    Рекурсивный поиск по именам папок и файлов
    :param root_folder: корневая папка для поиска
    :param search_query: искомый фрагмент текста
    :return: список найденных путей
    """
    results = []

    for root, dirs, files in os.walk(root_folder):
        # Поиск в именах папок
        for dir_name in dirs:
            if search_query.lower() in dir_name.lower():
                full_path = os.path.join(root, dir_name)
                results.append(f"ПАПКА: {full_path}")

        # Поиск в именах файлов
        for file_name in files:
            if search_query.lower() in file_name.lower():
                full_path = os.path.join(root, file_name)
                results.append(f"ФАЙЛ: {full_path}")

    return results

# Использование
root_folder = r"\\Esk-node2\ора\Договора Стачек\Договора арендаторы"  # Укажите путь к вашей папке X
search_query = input("Введите фрагмент для поиска: ")
found_items = search_in_folders(root_folder, search_query)

if found_items:
    print("\nНайденные элементы:")
    for item in found_items:
        print(item)
else:
    print("Ничего не найдено.")




########################## выводим содержимое папки
##import os
##
##path='\\\\Esk-node2\\ора\\Договора Стачек\\Договора арендаторы\\Лебедев'
##
##for folderName, subfolders, filenames  in os.walk(path):
##    folderName1= os.path.basename(folderName)
##    print('Текущая папка - ' + folderName)
##
##    for subfolder in subfolders:
##        print('Подпапка ПАПКИ' + folderName + ': ' + subfolder)
##
##    for filename in filenames:
##        print('Файл в ПАПКЕ '+ folderName1 + ': ' + filename )
##
##    print(' ')