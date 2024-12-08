
from openpyxl.styles.builtins import total
from pyexpat.errors import messages
import FreeSimpleGUI as sg
import time

with open("Files/color.txt", "r") as f:
    COLOR = f.read()


# Сначала устанавливаем тему
#sg.theme('DarkAmber')
sg.theme(COLOR)
contacts_file = "Files/contacts.xlsx"  # Путь к файлу с контактами
output_file = "Files/sms_log.xlsx"


from SmsToolsN import *



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

def sets():
    all_themes = sg.theme_list()
    # Затем определяем интерфейс
    layout = [
        [sg.Text("Настройка темы: "), sg.Combo(all_themes, default_value=sg.theme(), key='theme', enable_events=True)]

    ]

    # Создание окна
    window = sg.Window('Настройки', layout)

    # Цикл событий
    while True:
        event, values = window.read()

        if event in (sg.WINDOW_CLOSED, 'theme'):  # когда меняется тема
            new_theme = values['theme']
            if new_theme:
                sg.theme(new_theme)
                with open("Files/color.txt", "w") as f:
                    f.write(new_theme)
            # Закрываем текущее окно и создаем новое с новой темой
            window.close()
            #menu_main()
            return

def menu_analysing():
    if do_continue("Анализировать данные? 🤨"):
        analysis()
        err_msg("Успешно 👌")

def sending(text, nums):
    pass


def menu_contacts():

    def reload_data():
        # Загружаем существующие контакты
        existing = search_contacts("Files/contacts.xlsx", values["args"])[1]

        print(f"Контакты = {existing}")

        # Преобразуем существующие контакты в формат для таблицы
        contacts_data = []
        if existing:
            for el in existing:
                contacts_data.append([el["name"], el["number"]])
        window["table"].update(values=contacts_data)



    selected_numbers = []
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
    total_console = ""

    layout = [
        [sg.Text('Имя:'), sg.InputText(key='name',size=(34,10)),sg.Button('Перезагрузить данные', bind_return_key=True)],
        [sg.Text('Телефон:'), sg.InputText(key='phone',size=(30,10)), sg.Button('Анализировать данные')],
        [sg.Button('Добавить контакт'), sg.Button('Очистить'), sg.Button('Удалить выбранные'), sg.Button('Написать выбранным')],
        [sg.Text('Список контактов:'), sg.Text('Аргументы для поиска: '), sg.InputText(key='args',size=(34,10))],
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
            selected_numbers = []
            for row_index in selected_rows:
                selected_contact = contacts_data[row_index]
                print(f"Выбран контакт: Имя = {selected_contact[0]}, Телефон = {selected_contact[1]}")
                selected_numbers.append(selected_contact[1])
                window['name'].update(selected_contact[0])
                window['phone'].update(selected_contact[1])


        if event in (sg.WINDOW_CLOSED, 'Выход'):
            break

        if event == "Анализировать данные":
            reload_data()
            menu_analysing()

        if event == 'Добавить контакт':
            if values['name'] and values['phone']:
                add_contacts("Files/contacts.xlsx", [[values["phone"].replace("+7", ""),values["name"]]])
                new_contact = [values['name'], values['phone']]
                contacts_data.append(new_contact)
                window['table'].update(values=contacts_data)

                reload_data()

                print(f"Добавлен новый контакт: {new_contact}")
                # Очищаем поля ввода
                window['name'].update('')
                window['phone'].update('')

        if event == 'Перезагрузить данные':
            reload_data()

        if event == "Удалить выбранные":
            if selected_numbers and do_continue(f"Удалить {len(selected_numbers)} контакта?" if len(selected_numbers)%10 < 5 and len(selected_numbers)%10 > 1 else f"Удалить {len(selected_numbers)} контактов?"):
                delete_contact(selected_numbers)
                reload_data()
                err_msg("Успешно.")
            else:
                err_msg("Сначала выберите контакты!")

        if event == "Написать выбранным":
            if selected_numbers:
                print(selected_numbers)
            else:
                err_msg("Сначала выберите контакты!")

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
        [sg.Button('Настройки'), sg.Button('Получить смс :(')],
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
            get_messages()
        if event == 'Меню добавления контактов':
            menu_contacts()
        if event == 'Настройки':
            sets()
    window.close()

def get_messages():
    # Затем определяем интерфейс
    layout = [
        [sg.Checkbox('Получать постоянно', key='continuous_receive', enable_events=True)],
        [sg.Button('Получить'), sg.Button('Сохранить'), sg.Button('Очистить'), sg.Button('Выход')],
        [sg.Text('Входящие сообщения:')],
        [sg.Multiline(size=(60, 20), key='messages', autoscroll=True, reroute_stdout=True,
                     reroute_stderr=False, write_only=True, disabled=True)],
    ]

    # Создание окна
    window = sg.Window('Получение сообщений', layout)

    total_messages = ""

    # Флаг для контроля постоянного получения
    continuous = False

    # Цикл событий
    while True:
        event, values = window.read(timeout=1000 if continuous else None)  # таймаут 1 секунда при постоянном получении

        if event in (sg.WINDOW_CLOSED, 'Выход'):
            break

        if event == 'continuous_receive':
            continuous = values['continuous_receive']
            
        if event == 'Получить':# or (continuous and event == sg.TIMEOUT_KEY)
            print("Получаем смс...")
            log = read_sms_and_save(modem_port, contacts_file, output_file)
            if log != '':
                total_messages = f"{total_messages}{log}"
            window["messages"].update(total_messages)



        if event == 'Очистить':
            window['messages'].update('')
            
        if event == 'Сохранить':
            # Здесь будет код сохранения сообщений
            print("Сообщения сохранены")

    window.close()

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

can_modem = False

if __name__ == "__main__":
    delete_contact(["+79875325498"])
    if modem_port != "COM":
        setup_modem(modem_port)
        can_modem = True
    print(menu_choose_contacts())
    menu_main()
