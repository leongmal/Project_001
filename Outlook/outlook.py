# """отправка письма из Outtlook"""
# import win32com.client


# def send_email():
    # outlook = win32com.client.Dispatch("Outlook.Application")
    # mail = outlook.CreateItem(0) # 0 = olMailItem

    # mail.To = "le24mur@yandex.ru"
    # mail.Subject = "Тема письма"
    # mail.Body = "Текст письма"
    
    # # Для HTML-форматирования используйте:
    # # mail.HTMLBody = "<h1>Заголовок</h1><p>Текст</p>"
    # # mail.BodyFormat = 2 # 1=Plain, 2=HTML

    # mail.Send() # Или .Display(), чтобы показать окно перед отправкой

# send_email()

# """поиск в  письмах Outtlook"""

import win32com.client
import re

outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

def show_all_folders():
    """Показывает информацию о всех стандартных папках"""
    print("\n" + "="*80)
    print("ИНФОРМАЦИЯ О ПАПКАХ OUTLOOK")
    print("="*80)
    
    folder_names = {
        3: "Отправленные",
        4: "Удаленные", 
        5: "Черновики",
        6: "Входящие",
        10: "Спам",
        15: "Архив"
    }
    
    for idx in [3, 4, 5, 6, 10, 15]:
        try:
            folder = namespace.GetDefaultFolder(idx)
            print(f"\n📁 Папка {idx}: {folder.Name}")
            print(f"   📊 Количество писем: {folder.Items.Count}")
            print(f"   📂 Путь: {folder.FolderPath}")
        except Exception as e:
            print(f"\n❌ Папка {idx}: ошибка - {e}")
    
    print("\n" + "="*80)

def select_search_folder():
    """Позволяет пользователю выбрать папку для поиска"""
    print("\nВЫБОР ПАПКИ ДЛЯ ПОИСКА")
    print("-" * 40)
    print("3 - Отправленные")
    print("4 - Удаленные")
    print("5 - Черновики")
    print("6 - Входящие")
    print("10 - Спам")
    print("15 - Архив")
    print("-" * 40)
    
    while True:
        try:
            choice = input("Выберите номер папки (3,4,5,6,10,15): ")
            folder_num = int(choice)
            if folder_num in [3, 4, 5, 6, 10, 15]:
                folder = namespace.GetDefaultFolder(folder_num)
                print(f"\n✅ Выбрана папка: {folder.Name}")
                return folder
            else:
                print("❌ Неверный номер. Выберите из списка.")
        except ValueError:
            print("❌ Введите число.")

def search_and_display():
    """Основная функция поиска и отображения писем"""
    # Показываем информацию о папках
    show_all_folders()
    
    # Выбираем папку для поиска
    search_folder = select_search_folder()
    
    # Поисковый запрос
    search_query = input("\n🔍 Введите текст для поиска: ")
    
    # DASL фильтр для поиска
    search_filter = f"@SQL=(\"urn:schemas:httpmail:textdescription\" LIKE '%{search_query}%') OR (\"urn:schemas:httpmail:subject\" LIKE '%{search_query}%')"
    
    print(f"\n🔍 Поиск писем, содержащих: '{search_query}'")
    print(f"📁 В папке: {search_folder.Name}")
    print("⏳ Подождите, выполняется поиск...")
    
    try:
        filtered_items = search_folder.Items.Restrict(search_filter)
        filtered_items.Sort("[ReceivedTime]", True)  # Сортировка по дате (новые сверху)
        
        print(f"\n📊 Найдено писем: {filtered_items.Count}")
        
        if filtered_items.Count == 0:
            print("❌ Письма не найдены.")
            return
        
        # Сохраняем все письма в список
        emails_list = []
        
        # Выводим список писем
        print("\n" + "="*80)
        print("СПИСОК НАЙДЕННЫХ ПИСЕМ:")
        print("="*80)
        
        for i, msg in enumerate(filtered_items, 1):
            if msg.Class == 43:
                received_time = msg.ReceivedTime.strftime("%d.%m.%Y %H:%M:%S")
                print(f"\n[{i}] 📅 {received_time}")
                print(f"    📌 Тема: {msg.Subject[:80]}")
                print(f"    👤 От: {msg.SenderName} <{msg.SenderEmailAddress}>")
                print(f"    📁 Папка: {search_folder.Name}")
                
                # Сохраняем письмо в список
                emails_list.append(msg)
        
        # Запрашиваем выбор письма
        print("\n" + "="*80)
        while True:
            try:
                choice = input(f"\n📋 Выберите номер письма для просмотра (1-{len(emails_list)}) или 'q' для выхода: ")
                
                if choice.lower() == 'q':
                    print("👋 Выход...")
                    break
                
                num = int(choice)
                if 1 <= num <= len(emails_list):
                    selected_email = emails_list[num-1]
                    
                    # Выводим полное содержимое письма
                    print("\n" + "="*80)
                    print(f"ПОЛНОЕ СОДЕРЖИМОЕ ПИСЬМА #{num}")
                    print("="*80)
                    
                    print(f"\n📅 Дата и время: {selected_email.ReceivedTime}")
                    print(f"📌 Тема: {selected_email.Subject}")
                    print(f"👤 Отправитель: {selected_email.SenderName}")
                    print(f"📧 Email отправителя: {selected_email.SenderEmailAddress}")
                    print(f"👥 Кому: {selected_email.To}")
                    
                    # Важность
                    importance_map = {0: "Низкая", 1: "Обычная", 2: "Высокая"}
                    print(f"⚠️ Важность: {importance_map.get(selected_email.Importance, 'Неизвестно')}")
                    
                    # Проверяем наличие вложений
                    if selected_email.Attachments.Count > 0:
                        print(f"\n📎 Вложения ({selected_email.Attachments.Count}):")
                        for att in selected_email.Attachments:
                            try:
                                size = att.Size / 1024  # в КБ
                                print(f"   • {att.FileName} ({size:.1f} KB)")
                            except:
                                print(f"   • {att.FileName}")
                    
                    # Выводим тело письма полностью
                    print("\n" + "─"*80)
                    print("ТЕКСТ ПИСЬМА:")
                    print("─"*80)
                    print(selected_email.Body)
                    print("─"*80)
                    
                    # Поиск контактных данных в письме
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
                    
                    # Спрашиваем, хочет ли пользователь сохранить письмо в файл
                    save = input("\n💾 Сохранить письмо в файл? (y/n): ").lower()
                    if save == 'y':
                        safe_subject = re.sub(r'[\\/*?:"<>|]', "_", selected_email.Subject[:50])
                        filename = f"email_{num}_{selected_email.ReceivedTime.strftime('%Y%m%d_%H%M%S')}_{safe_subject}.txt"
                        with open(filename, 'w', encoding='utf-8') as f:
                            f.write(f"Тема: {selected_email.Subject}\n")
                            f.write(f"От: {selected_email.SenderName} <{selected_email.SenderEmailAddress}>\n")
                            f.write(f"Дата: {selected_email.ReceivedTime}\n")
                            f.write(f"Кому: {selected_email.To}\n")
                            f.write(f"Папка: {search_folder.Name}\n")
                            f.write("\n" + "="*80 + "\n\n")
                            f.write(selected_email.Body)
                            
                            if phones:
                                f.write("\n\n" + "="*80 + "\n")
                                f.write("НАЙДЕННЫЕ ТЕЛЕФОНЫ:\n")
                                for phone in set(phones):
                                    f.write(f"• {phone}\n")
                            
                            if emails:
                                other_emails = [e for e in set(emails) if e != selected_email.SenderEmailAddress]
                                if other_emails:
                                    f.write("\n" + "="*80 + "\n")
                                    f.write("НАЙДЕННЫЕ EMAIL:\n")
                                    for email in other_emails:
                                        f.write(f"• {email}\n")
                        
                        print(f"✅ Письмо сохранено в файл: {filename}")
                    
                    # Спрашиваем, хочет ли пользователь продолжить
                    cont = input("\n📖 Посмотреть другое письмо? (y/n): ").lower()
                    if cont != 'y':
                        break
                else:
                    print(f"❌ Пожалуйста, введите число от 1 до {len(emails_list)}")
            except ValueError:
                print("❌ Пожалуйста, введите корректное число")
            except Exception as e:
                print(f"❌ Ошибка: {e}")
    
    except Exception as e:
        print(f"❌ Ошибка при поиске: {e}")

# Запуск программы
if __name__ == "__main__":
    print("\n" + "="*80)
    print("🔍 ПРОГРАММА ПОИСКА В OUTLOOK")
    print("="*80)
    
    while True:
        search_and_display()
        
        # Спрашиваем, хочет ли пользователь выполнить новый поиск
        again = input("\n🔄 Выполнить новый поиск? (y/n): ").lower()
        if again != 'y':
            print("\n👋 До свидания!")
            break
            
            
            
