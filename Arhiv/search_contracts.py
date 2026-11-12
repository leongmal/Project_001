import os
import re
import sys
import traceback

# --- НАСТРОЙКИ ---
path = r"\\Esk-node2\ора\Договора Стачек\Договора дочерних потребителей"
phrase = "в целях исполнения им Государственного контракта"
# ------------------

# Преобразуем фразу в регулярное выражение \W+ (любые символы между словами)
words = phrase.split()
regex_pattern = r"\b" + r"\W+".join(words) + r"\b"
compiled_regex = re.compile(regex_pattern, flags=re.IGNORECASE | re.DOTALL)

print(f"Поиск фразы: '{phrase}'")
print(f"Путь: {path}")
print("-" * 50)

found_anything = False

def extract_text_from_docx(file_path):
    """Извлекает текст из .docx"""
    try:
        import docx2txt
        return docx2txt.process(file_path)
    except Exception:
        return None

def extract_text_from_pdf(file_path):
    """Извлекает текст из .pdf"""
    try:
        import pdfplumber
        text = ""
        with pdfplumber.open(file_path) as pdf:
            for page in pdf.pages:
                # Если страница пуста для OCR-сканов, метод .extract_text() вернет None
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"
        return text.strip() if text else None
    except Exception:
        return None

def extract_text_generic(file_path):
    """Попытка прочитать файл как текст (для .txt, .rtf или кривых .doc)"""
    try:
        # Пробуем разные популярные кодировки
        encodings = ['utf-8', 'cp1251', 'utf-16']
        for enc in encodings:
            try:
                with open(file_path, 'r', encoding=enc) as f:
                    return f.read()
            except UnicodeDecodeError:
                continue
    except Exception:
        pass
    
    # Крайний случай: открываем как latin-1 (отобразит кракозябры, но хотя бы прочитает байты)
    try:
        with open(file_path, 'r', encoding='latin-1') as f:
            return f.read()
    except Exception:
        return None


for root, dirs, files in os.walk(path):
    for filename in files:
        full_path = os.path.join(root, filename)
        
        file_text = None
        
        # Выбираем способ чтения в зависимости от расширения
        if filename.lower().endswith('.docx'):
            file_text = extract_text_from_docx(full_path)
        elif filename.lower().endswith('.pdf'):
            file_text = extract_text_from_pdf(full_path)
        elif filename.lower().endswith('.doc'):
            # Старые .doc очень сложны, пробуем сначала как общий текст
            file_text = extract_text_generic(full_path)
        else:
            # Для всех остальных типов тоже попробуем прочитать как текст
            file_text = extract_text_generic(full_path)

        # Если текст удалось получить — ищем фразу
        if file_text and compiled_regex.search(file_text):
            found_anything = True
            print(f"НАЙДЕНО: {full_path}")
            
            # Ищем конкретное место для контекста
            lines = file_text.splitlines()
            for i, line in enumerate(lines):
                if compiled_regex.search(line):
                    context_start = max(0, i - 1)
                    context_end = min(len(lines), i + 2)
                    
                    print("Контекст:")
                    for j in range(context_start, context_end):
                        marker = ">>> " if j == i else "    "
                        print(f"{marker}{lines[j].strip()}")
                    break
            print("-" * 50)

if not found_anything:
    print("\nФраза не найдена ни в одном доступном для чтения файле.")
    print("Возможные причины:")
    print("1. Файлы являются сканами (картинками). Требуется распознавание текста (OCR).")
    print("2. Файлы защищены паролем.")
    print("3. У компьютера нет доступа к сетевой папке \\Esk-node2 под текущей учетной записью.")