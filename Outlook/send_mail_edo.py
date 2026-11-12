# """отправка письма из Outtlook"""
#import win32com.client


#def send_email():
    #outlook = win32com.client.Dispatch("Outlook.Application")
    #mail = outlook.CreateItem(0) # 0 = olMailItem

    #mail.To = "le24mur@yandex.ru"
    #mail.Subject = "Тема письма"
    #mail.Body = "Текст письма "
    
    # Для HTML-форматирования используйте:
    # mail.HTMLBody = "<h1>Заголовок</h1><p>Текст</p>"
    # mail.BodyFormat = 2 # 1=Plain, 2=HTML

    #mail.Send() # Или .Display(), чтобы показать окно перед отправкой

#send_email()

"""отправка писем без вложений по шаблону """
import win32com.client
import time

def send_personalized_emails():
    outlook = win32com.client.Dispatch("Outlook.Application")
    
    clients = [
        {"company_name": 'ООО "КОМПАНИЯ "АНГСТРЕМ ТРЕЙДИНГ""', "email": "angstrem@angstrem.lek.ru"},
        # {"company_name": 'ПАО Банк "Санкт-Петербург"', "email": "cc@bspb.ru"},
        # {"company_name": "ИП Белынский", "email": "kira1607@yahoo.com"},
        # {"company_name": "ООО «ИнвестСтрой»", "email": "info@inv-str.ru"},
        # {"company_name": "ООО «РЦПП КВАЛИТЕТ»", "email": "rcppkvalitet@yandex.ru"},
        # {"company_name": "ООО Фирма «Парсек»", "email": "info@parsek-spb.ru"},
        # {"company_name": "ООО «ПОСТ»", "email": "info@postomsk.ru"},
        # {"company_name": 'ООО "МЗ "СЕВЕРНЫЙ МЕТАЛЛОЦЕНТР""', "email": "info@smcspb.ru"},
        # {"company_name": 'АО "ТУРБОТЕКТ САНКТ-ПЕТЕРБУРГ"', "email": "info@turbotectspb.com"},
        # {"company_name": "ИП Храмцов", "email": "ofic-info@list.ru"}
    ]

    subject_base = "Приглашение для работы в ЭДО"
    
    # Шаблон письма с правильным размещением переменной {company_name}
    body_template = """          Добрый день!

Для повышения эффективности взаимодействия и оптимизации рабочих процессов между АО «ЭСК» и {company_name} предлагаем рассмотреть возможность перехода на электронный документооборот (ЭДО).

Для подготовки Соглашения ЭДО необходима следующая информация:
- ФИО лица, ответственного за ЭДО;
- номер телефона лица, ответственного за ЭДО;
- эл. почта лица, ответственного за ЭДО;
- ваш доверенный Оператор ЭДО.

Заранее благодарим за оперативный ответ.

С уважением,
АО «ЭСК»"""

    for client in clients:
        try:
            print(f"Подготовка письма для {client['company_name']} ({client['email']})...")
            
            mail = outlook.CreateItem(0)
            mail.To = client["email"]
            # Подставляем название компании в тему
            mail.Subject = f"{subject_base} — {client['company_name']}"
            # Корректная подстановка названия компании в тело письма
            mail.Body = body_template.format(company_name=client["company_name"])
            
            # Отправка письма
            #mail.Send()
            mail.Display()
            print(f"Письмо успешно отправлено на {client['email']}")
            
            # Задержка перед следующим письмом (рекомендуется 5-10 секунд)
            time.sleep(15) 
            
        except Exception as e:
            print(f"Ошибка при работе с {client['email']}: {e}")
            # Продолжаем цикл даже после ошибки
            continue

#send_personalized_emails()