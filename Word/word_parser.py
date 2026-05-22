from docx import Document
import re

def extract_consumer_info(docx_file, output_file):
    doc = Document(docx_file)
    consumer_info = []

    # Ищем таблицу с реквизитами
    for table in doc.tables:
        for row in table.rows:
            cells = row.cells
            for cell in cells:
                text = cell.text
                # Проверяем, содержит ли ячейка информацию о Потребителе
                # if 'Потребитель:' in text:
                if 'Сторона 1' in text:
                    # Очищаем текст от лишних символов
                    cleaned_text = text.replace('|', '').strip()
                    # Разбиваем текст на строки
                    lines = cleaned_text.split('\n')
                    for line in lines:
                        line = line.strip()
                        if line:  # пропускаем пустые строки
                            consumer_info.append(line)

    # Форматируем текст с правильными переносами
    formatted_text = '\n'.join(consumer_info)

    # Сохраняем в файл
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(formatted_text)

    print(f"Данные успешно сохранены в файл {output_file}")

# Запуск скрипта
input_docx = 'ДС.docx'
output_txt = 'Реквизиты_Потребителя.txt'
extract_consumer_info(input_docx, output_txt)
