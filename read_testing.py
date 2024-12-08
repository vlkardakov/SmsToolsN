import serial
import time
from smspdudecoder.codecs import GSM
from smspdudecoder.easy import read_incoming_sms

"""
Шпаргалка по командам:
Перезагрузить модем: AT+CFUN=1,1
Задать память на память модема: AT+CPMS="ME","ME","ME"
задать текстовый режим/режим upd: AT+CMGF=1/ или 2
Прочитать все сообщения: AT+CMGL="ALL" //не забыть сначала текстовый режим поставить

"""


def send_at_command(ser, command, response_timeout=1):
    ser.write((command + '\r\n').encode())
    time.sleep(response_timeout)
    response = ser.read_all().decode()
    return response

def read_sms(port, baudrate=9600):
    try:
        # Open the serial port
        with serial.Serial(port, baudrate, timeout=1) as ser:
            # Check if the modem is responsive
            response = send_at_command(ser, 'AT')

            # Set SMS text mode


            # Read SMS messages
            response = send_at_command(ser, 'AT+CMGL="ALL"')
            print(response)

            # Parse the response to extract SMS details
            sms_list = response.split('\r\n')
            for sms in sms_list:
                if '+CMGL:' in sms:
                    parts = sms.split(',')
                    index = parts[0].split(':')[1].strip()
                    status = parts[1].strip().strip('"')
                    sender = parts[2].strip().strip('"')
                    timestamp = parts[4].strip().strip('"')
                    message = sms_list[sms_list.index(sms) + 1]
                    print(f"Index: {index}")
                    print(f"Status: {status}")
                    print(f"Sender: {sender}")
                    print(f"Timestamp: {timestamp}")
                    print(f"Message: {message}")
                    print("-" * 40)

    except serial.SerialException as e:
        print(f"Serial port error: {e}")


baudrate = 9600
port = "COM29"


def parse_sms_response(response):
    messages = []
    lines = response.splitlines()
    i = 0
    '''
    пример сообщения: 
    +CMGL: 10,"REC READ","+79875324724",,"24/11/15,14:35:51+12"
    Hello!
    Или:

    '''
    while i < len(lines):
        if "+CMGL: " in lines[i]:
            parts = lines[i].split(",")
            index = parts[0].split(": ")[1].strip()
            sender_number = parts[2].strip('"')
            date_and_time = lines[i].split(",,")[1].replace('"','').split(',')
            print(f"{date_and_time=}")
            date_dates = date_and_time[0].split("/")
            date_date = f"{date_dates[2]}.{date_dates[1]}.{date_dates[0]}"
            #print(f"ДАТА = {date_date}")

            date_time = date_and_time[1].split("+")[0].split("-")[0]
            #print(f"ВРЕМЯ = {date_time}")


            message_lines = []

            # Проверяем, есть ли следующая строка
            if i + 1 < len(lines):
                # Если следующая строка не начинается с "+CMGL: ", то добавляем ее к сообщению
                j = i + 1
                while j < len(lines) and "+CMGL: " not in lines[j]:
                    if "OK" in lines[j]:
                        break
                    message_lines.append(lines[j].strip())
                    j += 1
                i = j - 1  # Установим индекс на последнюю строку сообщения

            # Декодируем строки сообщения
            decoded_lines = []
            for line in message_lines:
                try:
                    # Преобразуем сообщение в формат UCS2
                    decoded_line = bytes.fromhex(line).decode('utf-16be')
                    decoded_lines.append(decoded_line)
                except (ValueError, UnicodeDecodeError):
                    # Если декодирование не удалось, оставляем все строки сообщения в первоначальном виде
                    decoded_lines = message_lines
                    break

            message = '\n'.join(decoded_lines)



            # Преобразуем дату в формат DD/MM/YYYY
            #formatted_date = format_date(date.strip())

            messages.append({
                "index": index,
                "sender_number": sender_number,
                "date": date_date,
                "time": date_time,
                "message": message.strip()
            })
        i += 1
    return messages

# Open the serial port
with serial.Serial(port, baudrate, timeout=1) as ser:
    if True:
        print("modem responsed")
        send_at_command(ser, 'AT+CMGF=1')
        send_at_command(ser, 'AT+CPMS="ME","ME","ME"')
        #send_at_command(ser, 'AT+CMGL="ALL"')



        while True:
            command = input()
            #sms = read_incoming_sms()
            ret = send_at_command(ser, command)
            print(ret)
            print(parse_sms_response(ret))

