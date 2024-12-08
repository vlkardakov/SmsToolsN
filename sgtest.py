from SmsToolsN import *
import FreeSimpleGUI as sg

# Сначала устанавливаем тему
sg.theme('DarkAmber') 

def do_continue(text):
    # Затем определяем интерфейс
    layout = [
        [sg.Text(text)],
        [sg.Button('НЕТ'), sg.Button('ДА')]
    ]

    # Создание окна
    window = sg.Window('Уведомление', layout)

    # Цикл событий
    while True:
        event, values = window.read()

        if event in (sg.WINDOW_CLOSED, "НЕТ"):
            window.close()
            return False

        if event == "ДА":
            window.close()
            return True


def menu_contacts():
    # Затем определяем интерфейс
    existing = load_contacts('Files/contacts.xlsx')
    print(f"Контакты = {existing}")

    layout = [
        [sg.Text('Имя:'), sg.InputText(key='name',size=(25,10))],
        [sg.Text('Телефон:'), sg.InputText(key='phone',size=(30,10))],
        [sg.Button('Добавить контакт'), sg.Button('Добавить контакт') , sg.Button('Выход')],
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
        [sg.Text(text), sg.Button('Смириться')]
    ]

    # Создание окна
    window = sg.Window('Уведомление', layout)

    # Цикл событий
    while True:
        event, values = window.read()

        if event in (sg.WINDOW_CLOSED, 'Смириться'):
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
            menu_contacts()
    window.close()

def menu_choose_contacts():
    # Затем определяем интерфейс
    layout = [
        [sg.Text('Аргументы: '), sg.InputText(key='search', enable_events=True)],
        [sg.Button('Поиск', size=5, bind_return_key=True), sg.Button('Отмена', size=5), sg.Button('Очистить', size=5)],
        [sg.Text('Список контактов:')],
        [sg.Multiline(size=(56, 10), key='contacts', disabled=True)],
        [sg.Button('Применить')]
    ]

    # Создание окна
    window = sg.Window('Поиск контактов', layout)

    # Список для хранения контактов
    contacts_list = []

    # Цикл событий
    while True:
        event, values = window.read()

        if event in (sg.WINDOW_CLOSED, 'Отмена'):
            break

        if event == "Выбрать все существующие" or event == "Поиск":
            search = values['search']
            numbers = []
            # Обновляем поле со списком контактов
            searched = search_contacts("Files/contacts.xlsx", search)
            contacts = searched[0]
            for el in searched[1]:
                numbers.append(str(el["number"]))

            window['contacts'].update('')
            complete = ""
            for contact in contacts:
                complete+=f"{contact}\n"
            window['contacts'].update(complete)

        if event == 'Очистить':
            window['contacts'].update('')
            window['search'].update('')
        if event == 'Применить':
            try:
                window.close()
                if numbers:
                    return numbers

            except Exception as e:
                err_msg(f"Неизвестная ошибка {e}")
            break


    window.close()

if __name__ == "__main__":
    print(menu_choose_contacts())
    menu_main()
