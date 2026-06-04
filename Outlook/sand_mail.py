# """отправка письма из Outtlook"""
import win32com.client


def send_email():
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0) # 0 = olMailItem

    mail.To = "le24mur@yandex.ru"
    mail.Subject = "Тема письма"
    mail.Body = "Текст письма"
    
    # Для HTML-форматирования используйте:
    # mail.HTMLBody = "<h1>Заголовок</h1><p>Текст</p>"
    # mail.BodyFormat = 2 # 1=Plain, 2=HTML

    mail.Send() # Или .Display(), чтобы показать окно перед отправкой

send_email()
