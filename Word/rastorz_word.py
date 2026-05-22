## внесение изменений в файл Согл_расторж2

from docx import Document

def read_replacement_dict_from_txt(file_path):
    replacement_dict={}
    with open(file_path, encoding='utf-8') as f:
        for line in f:
            parts = line.strip().split(':')
            if len(parts) != 2:
                continue 
            old_word, new_word = parts
            replacement_dict[old_word]=new_word
#    print(replacement_dict)
    return replacement_dict

def replace_text_in_docx(file_path, word_replacements):
    document = Document(file_path)
    for paragraph in document.paragraphs:
        for old_word, new_word in word_replacements.items():
            if old_word in paragraph.text:
                paragraph.text = paragraph.text.replace(old_word, new_word)
    document.save(file_path)
    
    
if __name__ == '__main__':
    word_file = r'C:\Users\Leonik\Desktop\Конверт\data_word.txt'
    word_replacements = read_replacement_dict_from_txt(word_file)
    
    docx_file=r'\\Esk-node2\ора\Договора Стачек\Договора арендаторы\Гидромеханика\расторжение\Согл_расторж2.docx'
    
    replace_text_in_docx(docx_file, word_replacements)
    print(f"Текст успешно обновлён в {docx_file}.")
