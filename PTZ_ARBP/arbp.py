import re, os

folder_PTZ = r'\\Esk-node2\ора\Договора Стачек\Договора дочерних потребителей\Петербургский тракторный завод'

valid_extension = ('.pdf')

for line in os.listdir(folder_PTZ):
    if line.endswith(valid_extension):
        name_list = os.path.splitext(line)
        if 'Приложение' in name_list or 'АРБП' in name_list:
            print(line)





