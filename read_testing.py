import serial
import time
"""
Шпаргалка по командам:
Перезагрузить модем: AT+CFUN=1,1
Задать память на память модема: AT+CPMS="ME","ME","ME"
задать текстовый режим/режим upd: AT+CMGF=1/
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
            if 'OK' not in response:
                print("Modem not responding. Check connection.")
                return

            # Set SMS text mode
            send_at_command(ser, 'AT+CMGF=1')

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
port = "COM3"
# Open the serial port
with serial.Serial(port, baudrate, timeout=1) as ser:
    # Check if the modem is responsive
    #response = send_at_command(ser, 'AT')
    #if 'OK' not in response:
        #print("Modem not responding. Check connection.")
        #exit()
    #else:
        print("modem responsed")

        while True:
            command = input()
            print(send_at_command(ser, command))



if __name__ == "__main__":
    # Replace 'COM3' with the appropriate port for your system
    modem_port = 'COM14'


    exit()
    read_sms(modem_port)
