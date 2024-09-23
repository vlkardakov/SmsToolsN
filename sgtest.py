# подключаем библиотеки
import FreeSimpleGUI as sg
import random
import os


# что будет внутри окна
# первым описываем кнопку и сразу указываем размер шрифта
layout = [[sg.Button('Новое число',enable_events=True, key='-FUNCTION-', font='Helvetica 16')],
        # затем делаем текст
        [sg.Text('Результат:', size=(25, 1), key='-text-', font='Helvetica 16')]]
# рисуем окно
window = sg.Window('Генератор случайных чисел', layout, size=(350,100))
# запускаем основной бесконечный цикл
while True:
    # получаем события, произошедшие в окне
    event, values = window.read()
    # если нажали на крестик
    # если нажали на кнопку
    if event == '-FUNCTION-':
        # запускаем связанную функцию
        os.system("start https://youtu.be/VPRjCeoBqrI?si=E1ovF1nebfEpmOug")
    if event in (sg.WIN_CLOSED, 'Exit'):
        # выходим из цикла
        break
# закрываем окно и освобождаем используемые ресурсы
window.close()