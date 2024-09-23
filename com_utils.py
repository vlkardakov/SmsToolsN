import serial
import serial.tools.list_ports as list_ports
import os



def find_available_ports():
    """Функция для поиска всех доступных COM портов."""
    ports = list(list_ports.comports())
    return [port.device for port in ports]

def read_settings(settings_file):
    """Функция для чтения настроек из файла."""
    if not os.path.exists(settings_file):
        print(f"Файл {settings_file} не существует.")
        return {}

    settings = {}
    with open(settings_file, 'r') as file:
        for line in file:
            # Пропускаем пустые строки и строки без знака '='
            if '=' not in line:
                continue
            name, value = line.strip().split('=', 1)  # split('=', 1) позволяет избежать ошибки
            settings[name.strip()] = value.strip()
    return settings

# Проверяем настройки отладки из файла settings.txt
settings_file = "Files/settings.txt"
debug_mode = False
settings = read_settings(settings_file)
if settings.get('debug') == '1':
    debug_mode = True


def send_at_command(port, debug=False):
    """
    Функция для отправки команды AT на указанный COM порт и получения ответа.
    Возвращает ответ на команду AT или None, если ответа нет.
    """
    try:
        ser = serial.Serial(port, timeout=2)
        ser.write(b'AT\r\n')
        response = ser.read(100).decode('utf-8').strip()
        ser.close()
        return response
    except serial.SerialException:
        if debug:
            print(f"Не удалось открыть порт {port}.              - debug")
        return None

# Находим все доступные COM порты
available_ports = find_available_ports()

if not available_ports:
    modem_port = 'COM'
else:
    num_ports = len(available_ports)
    if debug_mode:
        if num_ports == 1:
            print(f"Найден 1 доступный порт, попытка подключения...")
        else:
            print(f"Найдено {num_ports} возможных порта, попытка подключения...")

    # Проверяем настройки отладки из файла settings.txt
    settings_file = "Files/settings.txt"
    debug_mode = False
    settings = read_settings(settings_file)
    if settings.get('debug') == '1':
        debug_mode = True

    # Проходим по каждому доступному порту
    modem_port = None
    for port in available_ports:
        if debug_mode:
            print(f"Отправка AT команды на порт {port}...        - debug")
        response = send_at_command(port, debug_mode)
        if response:
            if debug_mode:
                print(f"Ответ от порта {port}: {response}                    - debug")
            # Сохраняем первый найденный порт и завершаем выполнение
            modem_port = port
            break
    if debug_mode:
        if modem_port is None:
            print("Не удалось подключить модем!")
        else:
            print("Модем подключен!                           - debug")
            if debug_mode:
                print(f"Модем найден на порту {modem_port}!                - debug")

# Теперь вы можете использовать переменную modem_port
if debug_mode:
    print("Модем порт: ", modem_port, '                         - debug')


