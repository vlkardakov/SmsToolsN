from SmsToolsN import *
import FreeSimpleGUI as sg

# Сначала устанавливаем тему
sg.theme('DarkAmber') 

def menu_add_contacts():
    # Затем определяем интерфейс
    layout = [
        [sg.Text('Имя:'), sg.InputText(key='name')],
        [sg.Text('Телефон:'), sg.InputText(key='phone')],
        [sg.Button('Добавить контакт'), sg.Button('Очистить'), sg.Button('Выход')],
        [sg.Text('Список контактов:')],
        [sg.Multiline(size=(40, 10), key='contacts', disabled=True)]
    ]

    # Создание окна
    window = sg.Window('Менеджер контактов', layout)

    # Список для хранения контактов
    contacts_list = []

    # Цикл событий
    while True:
        event, values = window.read()

        if event in (sg.WINDOW_CLOSED, 'Выход'):
            break

        if event == 'Добавить контакт':
            if values['name'] and values['phone']:
                contact = f"Имя: {values['name']}, Телефон: {values['phone']}"
                contacts_list.append(contact)
                print(f"Добавлен новый контакт: {contact}")
                # Обновляем поле со списком контактов
                window['contacts'].update('\n'.join(contacts_list))
                # Очищаем поля ввода
                window['name'].update('')
                window['phone'].update('')

        if event == 'Очистить':
            window['name'].update('')
            window['phone'].update('')
    window.close()

def err_msg(text):
    # Затем определяем интерфейс
    layout = [
        [sg.Text(text), sg.Button('OK')]
    ]

    # Создание окна
    window = sg.Window('Главное меню', layout)

    # Цикл событий
    while True:
        event, values = window.read()

        if event in (sg.WINDOW_CLOSED, 'OK'):
            break
    window.close()

def menu_main():
    # Затем определяем интерфейс
    layout = [
        [sg.Button('Отправить смс :('), sg.Button('Получить смс :(')],
        [sg.Button('Меню добавления контактов'), sg.Button('Выход')],

    ]

    # Создание окна
    window = sg.Window('Главное меню', layout)

    # Цикл событий
    while True:
        event, values = window.read()

        if event in (sg.WINDOW_CLOSED, 'Выход'):
            break
        if event == 'Отправить смс :(' or event == "Получить смс :(":
            err_msg("Эта функция пока  не поддерживается в графическом интерфейсе :(")
        
        if event == 'Меню добавления контактов':
            menu_add_contacts()
    window.close()

if __name__ == "__main__":
    menu_main()
