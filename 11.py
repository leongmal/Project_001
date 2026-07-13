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
print('*'*100)
##print(st.lower())
##################
##import random
##
##def email():
##    n=random.randint(300,599)
##    print(f'MihLih{n}@ya.ru')
##email()    
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

""""3.5"""
# import random

# def show(A):
#     for a in A:
#         for s in a:
#             print(s, end='')
#         print()


# m=int(input('какое число строк в матрице?  :'))
# n=int(input('какое число столбцов в матрице?  :'))
# def rands(m,n):
#     res=[[random.randint(0,9) for i in range(n)] for i in range(m)]
#     return res

# matrix = rands(m,n)
# print('исходная матрица : ') 
# show(matrix)

# a= int(input(' какую строку удалить? :' ))-1
# # b =int(input('какой столбец удалить ?'))
# def del_str( matrix,a):
#     if 0 <= a <len(matrix):
#         matrix.pop(a)
#         for i in range(len(matrix)):
#             matrix
#     else:
#         print('ошибка, значение вне диапазона')
#     return matrix

# matrix= del_str(matrix, a)
# print('матрица после удаления строки ....')
# show(matrix)

# b =int(input('какой столбец удалить ?'))-1

# def del_row (matrix, b):
#     if not b >= len(matrix[0]):
#         print('ошибка, значение вне диапазона')
#     else:
#         return matrix
#     for row in matrix:
#         row.pop(b)
#     return(matrix)
# matrix=del_row(matrix, b)
# print('матрица после удаления столбца ..........')
# show(matrix)

"""3.3"""
# def symbs(m,n):
#     val='A'
#     res= [['' for i in range(n)] for j in range(m)]
#     for i in range(m):
#         for j in range(n):
#             res[i][j] =val
#             val = chr(ord(val)+1)
#     return res
# print(symbs(4,5))
# show(symbs(4,5))
"""3.7"""
# a = [16,5,2,8,14,3,4,2]
# max_a=max(a)
# ind_a= a.index(max_a)+1

# print(f' максимальное значение в списке {max_a}, а позиция в списке {ind_a}','\n')
"""3.8"""
# import random
# list_rand_num = [random.randint(21,83) for n in range(20)]
# # ind_even =[]
# # ind_odd =[]
# # for index, i in enumerate(list_rand_num):
# #     if index % 2 == 0:
# #         ind_odd.append(i)
# #     else:
# #         ind_even.append(i)
# """ подсмотрено но очень здорово! """
# ind_even= sorted(list_rand_num[::2])
# ind_odd= sorted(list_rand_num[1::2], reverse=True)
# new=[]
# for i in range(len(ind_even)):
#     new.append(ind_odd[i])
#     new.append(ind_even[i])

# print(list_rand_num,'\n')
# print('числа с четными индексами' , ind_even)
# print('числа с НЕчетными индексами' , ind_odd, '\n')
# print('новый список..............')
# print(new)
"""3.9"""
# import random

# rand_list =[random.randint(0,9) for i in range(10)]
# # rand_list =[7, 7, 1, 6, 8, 1] 
# new=[]

# for i in range(0, len(rand_list), 2):
#     s_num=rand_list[i]+rand_list[i+1]
#     new.append(rand_list[i])
#     new.append(s_num)
#     new.append(rand_list[i+1])
  
# print(rand_list,'\n')
# print(new)

 


# days=["Пн","Вт","Ср","Чт","Пт","Сб", "Вс"]

# week={days[s]:s for s in range(len(days))}
# add_week ={"Ух":7,"Му":8}
# sum_week = week.update(add_week)
# print(week,'\n')
# print(sum_week)

# symbs = input('введитетекстовое значение:  ')
# symbs = 'ASDDDASA'
# res={}
# key_symb =set(symbs)
# for i in key_symb:
#     res[i]=symbs.replace(i,'',1)
# print(symbs)    
# print(key_symb)
# print(res)

