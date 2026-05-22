### Проверка исключения ( на ввод значений - не числа)
def collatz(number):
    if number%2 == 0:
        return number // 2
    else:
        return 3*number+1

def input_number():
    print('Веди число')
    while True:
        try:
            a= int(input())
            break   
        except ValueError:
            print('Введи целое число')
    
    while int(a) > 1:
        print(a)
        a = collatz(a)
    print(a)
    
    
input_number()   
