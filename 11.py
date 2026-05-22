##from openpyxl import load_workbook
##import time
##
##fn =  'example.xlsx'
##wb = load_workbook(fn)
##ws = wb['data']
##
##ws['A5'] = 'Проба'
##ws['B5'] = '654'
##ws['C6'] = time.strftime("%H:%M:%S")
##wb.save(fn)
##wb.close()
##print('Ok!')

##import os
##import re
##
##pattern = r' status='1'>'
##replacement = ' '
##
##with open('D:/12.txt','r') as file:
##    content = file.read()
##
##content_new = re.sub(pattern, replacement, content)
##########
##b = []
##with open('dua.txt','r') as f:
##    data =''.join(f.read())
##    for i in data:
##        if i != '\n':
##            b.append(i)
##        else:
##            b.append('-')
##                
##print(''.join(b) )
##################
##f = open('dua.txt','r')
##b = []
##while True:
##    line = f.readline()
##    b.append(line,end ='')
##    if not line:
##        break
##print(b)    
##############################  отрытие, множество, запись  
##f = open('dua.txt','r')        
##lin = []
##lines = f.readlines()
##for line in lines:
##    lin += line.split()
###    print(line.split())
##print(len(lin), lin)
##d =str(set(lin))
##ff = open('dua1.txt','w')
##ff.write(d)
##f.close()
##ff.close()
#####################################

##import datetime
##### Создание объекта datetime
##date = datetime.datetime(2024,8,8)
##
##print(date.weekday())
##############################
st ='195277, Г.САНКТ-ПЕТЕРБУРГ, ВН.ТЕР.Г. МУНИЦИПАЛЬНЫЙ ОКРУГ САМПСОНИЕВСКОЕ, ПР-КТ ФИНЛЯНДСКИЙ, Д. 4, ЛИТЕРА А, ПОМЕЩ. 14-Н-878, 879, ОФИС 331' 
print(st.title())
print('***************************************************************************')
##print(st.lower())
##################
##import random
##
##def email():
##    n=random.randint(300,599)
##    print(f'MihLih{n}@ya.ru')
##email()    
#####################
##tableData = [['apples', 'oranges', 'cherries', 'banana'],
##             ['Alice', 'Bob', 'Carol', 'David'],
##             ['dogs', 'cats', 'moose', 'goose']]
##a=[]
##
##for i in tableData:
##    a+=i
##long_a=max(a, key=len)
##
##def printTable(data):
##    # Получаем количество строк и столбцов
##    num_rows = len(data)      # 3 строки (фрукты, имена, животные)
##    num_cols = len(data[0])   # 4 столбца (каждый элемент в строке) — исправлено!
##
##    # Проходим по каждому столбцу (это будут строки вывода)
##    for col in range(num_cols):
##        # Формируем строку для вывода, беря элемент из каждой строки по текущему столбцу
##        row_output = ' '.join([data[row][col] for row in range(num_rows)])
##        print(row_output)
##
##printTable(tableData)
##  
##print(len(tableData))

#######################
##import re
##
##a= 'это мой телефонный номер 8-921-30-80-720'
##phoneRegex=re.compile(r'\d-\d{3}-\d{2}-\d{2}-\d{3}')
##mo=phoneRegex.search(a)
##print(mo.group())
########################## выводим содержимое папок
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
############################# просмотр параграфа документа
##from docx import Document
##
##document =Document(r'C:\Users\Leonik\Desktop\Образцы\исходники_не для работы\Соглашение ЭДО - ШАБЛОН.docx')
##paragraph = document.paragraphs[56].text
##paragraph=document.paragraphs[4].text
##print(paragraph)
##print(len(document.paragraphs))

##################промсмотр таблиц в документе
##from docx import Document
##
## # Загрузка документа
##doc = Document(r'C:\Users\Leonik\Desktop\Образцы\исходники\Соглашение ЭДО - ШАБЛОН.docx')
##
## # Перебор всех таблиц в документе
##for t_idx, table in enumerate(doc.tables):
##    print(f'Таблица {t_idx + 1}:')
##    # Перебор строк в таблице
##    for r_idx, row in enumerate(table.rows):
##        # Получение текста из каждой ячейки строки
##        row_texts = [cell.text for cell in row.cells]
##        print(f'  Строка {r_idx + 1}: ' + ' | '.join(row_texts))
##    print('')
######################################
#dict_abonents={}

with open('nabor.txt', 'r', encoding ='utf-8') as file:
    content= file.read()
print(content)
    