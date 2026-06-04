import win32com.client
import re

outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

def get_all_folders():
    """Рекурсивно получает все папки из всех хранилищ"""
    all_folders = []
    
    def traverse_folders(folder, path=""):
        """Рекурсивный обход папок"""
        current_path = f"{path}/{folder.Name}" if path else folder.Name
        all_folders.append({
            'name': folder.Name,
            'path': current_path,
            'folder': folder,
            'parent_path': path
        })
        
        # Рекурсивно обходим вложенные папки
        for subfolder in folder.Folders:
            traverse_folders(subfolder, current_path)
    
    # Обходим все хранилища (почтовые ящики)
    for store in namespace.Stores:
        try:
            root_folder = store.GetRootFolder()
            traverse_folders(root_folder)
        except Exception as e:
            print(f"Ошибка при доступе к хранилищу {store.DisplayName}: {e}")
    
    return all_folders

def show_all_folders_tree():
    """Показывает дерево всех папок"""
    print("\n" + "="*80)
    print("ДЕРЕВО ВСЕХ ПАПОК")
    print("="*80)
    
    folders = get_all_folders()
    
    # Группируем по путям для красивого вывода
    for folder in folders:
        indent = "  " * (folder['path'].count('/'))
        print(f"{indent}📁 {folder['name']}")
    
    return folders

def select_folder_from_list(folders):
    """Позволяет пользователю выбрать папку из списка"""
    print("\n" + "="*80)
    print("ВЫБОР ПАПКИ ДЛЯ ПОИСКА")
    print("="*80)
    
    # Показываем все папки с номерами
    for i, folder in enumerate(folders, 1):
        indent = "  " * (folder['path'].count('/'))
        print(f"{i:3}. {indent}📁 {folder['name']} ({folder['path']})")
    
    print(f"\n{len(folders) + 1}. Поиск во ВСЕХ папках")
    
    while True:
        try:
            choice = input(f"\nВыберите номер папки (1-{len(folders) + 1}): ")
            num = int(choice)
            
            if 1 <= num <= len(folders):
                selected = folders[num-1]
                print(f"\n✅ Выбрана папка: {selected['path']}")
                return selected['folder']
            elif num == len(folders) + 1:
                print("\n✅ Выбран поиск во всех папках")
                return None  # None означает поиск во всех папках
            else:
                print(f"❌ Введите число от 1 до {len(folders) + 1}")
        except ValueError:
            print("❌ Введите корректное число")

def search_in_all_folders(folders, search_query):
    """Ищет письма во всех папках"""
    results = []
    search_filter = f"@SQL=(\"urn:schemas:httpmail:textdescription\" LIKE '%{search_query}%') OR (\"urn:schemas:httpmail:subject\" LIKE '%{search_query}%')"
    
    print(f"\n🔍 Поиск во всех папках...")
    
    for i, folder_info in enumerate(folders, 1):
        folder = folder_info['folder']
        print(f"   {i}/{len(folders)} Проверка: {folder_info['path']}")
        
        try:
            filtered = folder.Items.Restrict(search_filter)
            if filtered.Count > 0:
                for item in filtered:
                    if item.Class == 43:
                        results.append({
                            'folder_path': folder_info['path'],
                            'folder_name': folder_info['name'],
                            'item': item
                        })
        except Exception as e:
            # Некоторые папки могут быть недоступны для поиска
            pass
    
    return results

def search_and_display():
    """Основная функция поиска и отображения писем"""
    
    # Получаем все папки
    print("Загрузка списка папок...")
    all_folders = get_all_folders()
    
    # Показываем дерево папок
    show_all_folders_tree()
    
    # Выбираем папку
    search_folder = select_folder_from_list(all_folders)
    
    # Поисковый запрос
    search_query = input("\n🔍 Введите текст для поиска: ")
    
    if search_folder is None:
        # Поиск во всех папках
        results = search_in_all_folders(all_folders, search_query)
        print(f"\n📊 Всего найдено писем: {len(results)}")
        emails_list = results
    else:
        # Поиск в выбранной папке
        search_filter = f"@SQL=(\"urn:schemas:httpmail:textdescription\" LIKE '%{search_query}%') OR (\"urn:schemas:httpmail:subject\" LIKE '%{search_query}%')"
        
        print(f"\n🔍 Поиск в папке: {search_folder.Name}")
        print("⏳ Подождите...")
        
        try:
            filtered_items = search_folder.Items.Restrict(search_filter)
            filtered_items.Sort("[ReceivedTime]", True)
            
            print(f"\n📊 Найдено писем: {filtered_items.Count}")
            
            emails_list = []
            for msg in filtered_items:
                if msg.Class == 43:
                    emails_list.append({
                        'folder_path': search_folder.Name,
                        'folder_name': search_folder.Name,
                        'item': msg
                    })
        except Exception as e:
            print(f"❌ Ошибка при поиске: {e}")
            return
    
    if len(emails_list) == 0:
        print("❌ Письма не найдены.")
        return
    
    # Выводим список писем
    print("\n" + "="*80)
    print("СПИСОК НАЙДЕННЫХ ПИСЕМ:")
    print("="*80)
    
    for i, email_data in enumerate(emails_list, 1):
        msg = email_data['item']
        received_time = msg.ReceivedTime.strftime("%d.%m.%Y %H:%M:%S")
        print(f"\n[{i}] 📅 {received_time}")
        print(f"    📌 Тема: {msg.Subject[:80]}")
        print(f"    👤 От: {msg.SenderName}")
        print(f"    📁 Папка: {email_data['folder_path']}")
    
    # Запрашиваем выбор письма
    print("\n" + "="*80)
    while True:
        try:
            choice = input(f"\n📋 Выберите номер письма (1-{len(emails_list)}) или 'q' для выхода: ")
            
            if choice.lower() == 'q':
                print("👋 Выход...")
                break
            
            num = int(choice)
            if 1 <= num <= len(emails_list):
                email_data = emails_list[num-1]
                selected_email = email_data['item']
                
                # Выводим полное содержимое письма
                print("\n" + "="*80)
                print(f"ПОЛНОЕ СОДЕРЖИМОЕ ПИСЬМА #{num}")
                print(f"📁 Папка: {email_data['folder_path']}")
                print("="*80)
                
                print(f"\n📅 Дата и время: {selected_email.ReceivedTime}")
                print(f"📌 Тема: {selected_email.Subject}")
                print(f"👤 Отправитель: {selected_email.SenderName}")
                print(f"📧 Email: {selected_email.SenderEmailAddress}")
                print(f"👥 Кому: {selected_email.To}")
                
                # Выводим тело письма
                print("\n" + "─"*80)
                print("ТЕКСТ ПИСЬМА:")
                print("─"*80)
                print(selected_email.Body)
                print("─"*80)
                
                # Поиск контактных данных
                phones = re.findall(r'\+?\d[\d\s\-\(\)]{7,}\d', selected_email.Body)
                emails = re.findall(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', selected_email.Body)
                
                if phones:
                    print("\n📞 НАЙДЕННЫЕ ТЕЛЕФОНЫ:")
                    for phone in set(phones):
                        print(f"   • {phone}")
                
                if emails:
                    other_emails = [e for e in set(emails) if e != selected_email.SenderEmailAddress]
                    if other_emails:
                        print("\n📧 НАЙДЕННЫЕ EMAIL:")
                        for email in other_emails:
                            print(f"   • {email}")
                
                print("\n" + "="*80)
                
                # Сохранение в файл
                save = input("\n💾 Сохранить письмо в файл? (y/n): ").lower()
                if save == 'y':
                    safe_subject = re.sub(r'[\\/*?:"<>|]', "_", selected_email.Subject[:50])
                    filename = f"email_{num}_{selected_email.ReceivedTime.strftime('%Y%m%d_%H%M%S')}.txt"
                    with open(filename, 'w', encoding='utf-8') as f:
                        f.write(f"Папка: {email_data['folder_path']}\n")
                        f.write(f"Тема: {selected_email.Subject}\n")
                        f.write(f"От: {selected_email.SenderName} <{selected_email.SenderEmailAddress}>\n")
                        f.write(f"Дата: {selected_email.ReceivedTime}\n")
                        f.write(f"Кому: {selected_email.To}\n")
                        f.write("\n" + "="*80 + "\n\n")
                        f.write(selected_email.Body)
                    print(f"✅ Письмо сохранено: {filename}")
                
                # Продолжить или нет
                cont = input("\n📖 Посмотреть другое письмо? (y/n): ").lower()
                if cont != 'y':
                    break
            else:
                print(f"❌ Введите число от 1 до {len(emails_list)}")
        except ValueError:
            print("❌ Введите корректное число")

# Запуск программы
if __name__ == "__main__":
    print("\n" + "="*80)
    print("🔍 ПРОГРАММА ПОИСКА В OUTLOOK")
    print("="*80)
    
    while True:
        search_and_display()
        
        again = input("\n🔄 Выполнить новый поиск? (y/n): ").lower()
        if again != 'y':
            print("\n👋 До свидания!")
            break