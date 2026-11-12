# game_langvige
import random
import keyboard

def load_vocabulary(filename):
    """Загружает слова из файла в словарь."""
    vocabulary = {}
    try:
        with open(filename, 'r', encoding='utf-8') as file:
            for line in file:
                clean_line = line.strip()
                if not clean_line or '-' not in clean_line:
                    continue
                
                parts = clean_line.split(' - ', 1)
                if len(parts) == 2:
                    eng_word = parts[0].strip().lower()
                    rus_word = parts[1].strip()
                    vocabulary[eng_word] = rus_word
    except FileNotFoundError:
        print(f"Ошибка: Файл '{filename}' не найден.")
        exit(1)
    except PermissionError:
        print(f"Ошибка: Нет прав на чтение файла '{filename}'.")
        exit(1)
    
    return vocabulary

def main():
    # Используем абсолютный путь к вашему файлу
    filename = r"C:\Users\Leonik\AppData\Local\Programs\Python\Python312\test_py\dict.txt"
    
    words_dict = load_vocabulary(filename)
    
    if not words_dict:
        print("Файл пустой или содержит ошибки формата.")
        return

    word_list = list(words_dict.items())
    current_index = None
    showing_translation = False

    print("Программа запущена. Нажмите Пробел, чтобы показать слово.")
    print("Нажмите Esc для выхода.\n")

    def on_space_press(e):
        nonlocal current_index, showing_translation
        
        if current_index is None or showing_translation:
            current_index = random.randint(0, len(word_list) - 1)
            eng_word, _ = word_list[current_index]
            print(f"[{current_index + 1}/{len(word_list)}] {eng_word}")
            showing_translation = False
        else:
            _, rus_word = word_list[current_index]
            print(f"-> {rus_word}\n")
            showing_translation = True
            current_index = None

    def on_esc_press(e):
        print("\nПрограмма завершена.")
        keyboard.unhook_all()
        exit(0)

    keyboard.on_press_key("space", on_space_press)
    keyboard.on_press_key("esc", on_esc_press)

    keyboard.wait()

if __name__ == "__main__":
    main()