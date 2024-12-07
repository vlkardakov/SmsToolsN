import serial
import binascii


def ucs_to_ansi(ucs_string):
    result = ''
    for i in range(0, len(ucs_string), 4):
        ucs_char = ucs_string[i:i + 4]
        j = int(ucs_char, 16)
        if 1040 <= j <= 1103:
            j -= 848
        elif j == 1105:
            j = 184
        result += chr(j)
    return result


def ansi_to_ucs(ansi_string):
    result = ''
    for char in ansi_string:
        j = ord(char)
        if 192 <= j <= 255:
            j += 848
        elif j == 184:
            j = 1105
        result += f'{j:04X}'
    return result


def send_sms_message(com_port, message, phone_number):
    # Открываем порт
    with serial.Serial(f'COM{com_port}', 9600, timeout=1) as ser:
        # Подготовка номера телефона
        if len(phone_number) % 2 == 1:
            phone_number += 'F'

        formatted_number = ''.join(phone_number[i:i + 2][::-1] for i in range(0, len(phone_number), 2))
        text_message = ansi_to_ucs(message)

        # Формируем сообщение
        sms_length = f'{len(text_message) // 2:02X}'
        sms_message = (
                '00'  # Длина и номер SMS центра
                + '11'  # SMS-SUBMIT
                + '00'  # Длина и номер отправителя
                + f'{len(phone_number):02X}'  # Длина номера получателя
                + '91'  # Тип-адреса
                + formatted_number  # Телефонный номер получателя
                + '00'  # Идентификатор протокола
                + '08'  # Кодировка (0 - латиница, 8 - кириллица)
                + 'C1'  # Срок доставки сообщения
                + sms_length  # Длина текста сообщения
                + text_message  # Текст сообщения
        )

        # Отправляем SMS
        ser.write(f'AT+CMGF=0\r\n'.encode())
        ser.write(f'AT+CMGS={len(sms_message) // 2}\r\n'.encode())
        ser.write((sms_message + '\r\n\x1A').encode())

if __name__ == '__main__':
    COM_PORT = 3  # Замените на номер вашего COM порта
    MESSAGE = 'Привет, как дела?'  # Сообщение для отправки
    PHONE_NUMBER = '+79875324724'  # Номер телефона в международном формате

    send_sms_message(COM_PORT, MESSAGE, PHONE_NUMBER)