from SmsToolsN import *
import FreeSimpleGUI as sg

# Сначала устанавливаем тему
#sg.theme('DarkAmber')
sg.theme('LightGreen3')

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
    # Загружаем существующие контакты
    existing = search_contacts("Files/contacts.xlsx", "")[1]

    print(f"Контакты = {existing}")
    
    # Создаем заголовки для таблицы
    headings = ['Имя', 'Телефон']



    # Преобразуем существующие контакты в формат для таблицы
    contacts_data = []
    if existing:
        for el in existing:
            contacts_data.append([el["name"], el["number"]])

    layout = [
        [sg.Text('Имя:'), sg.InputText(key='name',size=(34,10)),sg.Button('Перезагрузить даные')],
        [sg.Text('Телефон:'), sg.InputText(key='phone',size=(30,10))],
        [sg.Button('Добавить контакт'), sg.Button('Очистить'), sg.Button('Удалить выбранные'), sg.Button('Написать выбранным')],
        [sg.Text('Список контактов:')],
        [sg.Table(values=contacts_data,
                 headings=headings,
                 max_col_width=35,
                 auto_size_columns=True,
                 alternating_row_color="",
                 justification='left',
                 num_rows=10,
                 key='table',
                 enable_events=True,
                 size=(60, 20),
                 select_mode=sg.TABLE_SELECT_MODE_EXTENDED)],
        [sg.Button('Сохранить'), sg.Button('Выход')]
    ]

    # Создание окна
    window = sg.Window('Менеджер контактов', layout)

    # Цикл событий
    while True:
        event, values = window.read()
        print(event)
        print(values)
        
        if event == 'table':  # когда кликаем по таблице
            selected_rows = values['table']
            for row_index in selected_rows:
                selected_contact = contacts_data[row_index]
                print(f"Выбран контакт: Имя = {selected_contact[0]}, Телефон = {selected_contact[1]}")
                window['name'].update(selected_contact[0])
                window['phone'].update(selected_contact[1])


        if event in (sg.WINDOW_CLOSED, 'Выход'):
            break

        if event == 'Добавить контакт':
            if values['name'] and values['phone']:
                new_contact = [values['name'], values['phone']]
                contacts_data.append(new_contact)
                window['table'].update(values=contacts_data)
                print(f"Добавлен новый контакт: {new_contact}")
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

def get_messages():
    pass

def menu_choose_contacts():
    # Затем определяем интерфейс
    layout = [
        [sg.Text('Аргументы: '), sg.InputText(key='search', enable_events=True)],
        [sg.Button('Поиск', bind_return_key=True), sg.Button('Отмена'), sg.Button('Очистить')],
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
