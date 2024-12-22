**Использованные библиотеки Python:**
> PySimpleGUI - создает графические пользовательские интерфейсы (GUI)

> Pyserial - Этот модуль инкапсулирует доступ к последовательному порту.

> Pandas, Openpyxl - Чтение и запись таблиц.

>
Поступил заказ: есть 180 датчиков отслеживания машин для защиты от угона. Все эти датчики нужно проверять через СМС-сообщения, и проверять уровень заряда. Это очень рутинная работа, которую нужно оптимизировать.
До создания этой программы **человеку** приходилось вручную отправлять СМС-команды всем маячкам через телефон. Теперь же достаточно сделать всего несколько кликов для отправки команд на все маячки. 

С помощью этих маячков заказчик проверяет местоположение и уровень заряда.




Массовая проверка маячков автомобилей на языке Python.

Этот провект помогает оптимизировать бизнес-процесс связи со скрытыми Маячками отслеживания автомобилей.

**Как использовать (Релиз)**: 
> Скачать последний релиз

> Разархивировать.

> Запустить SmsTools.exe

**Как использовать в Python**
> `git clone https://github.com/vlkardakov/SmsToolsN.git`
> Или скачать **исходный код.**

> Открыть командную строку в директории с sgtest.py и прописать `pip install altgraph==0.17.4 colorama==0.4.6 et_xmlfile==2.0.0 FreeSimpleGUI==5.1.1 numpy==2.2.0 openpyxl==3.1.5 packaging==24.2 pandas==2.2.3 pefile==2023.2.7 psutil==6.1.0 pyinstaller==6.11.1 pyinstaller-hooks-contrib==2024.10 pyserial==3.5 python-dateutil==2.9.0.post0 python-gsmmodem-new==0.13.0 pywin32-ctypes==0.2.3 setuptools==75.6.0 six==1.17.0 tzdata==2024.2 six==1.17.0`

> Запустить программу `python sgtest.py`

**И радоваться**
![Menu](https://github.com/user-attachments/assets/61ab5261-1053-484b-b495-e625a994e821)

![image](https://github.com/user-attachments/assets/61c3d6e8-6757-4399-8ee9-ff309cf7fc9c)

**Инструкция**
> Перед запуском программы вставить модем, установить **`Connect Manager`**и добиться, чтобы на модеме был **синий / голубой** индикатор.

> Запустить программу.

> Подождать.

> Меню должно загрузиться.

> **Если** какая-то функция **не работает**, нажмите кнопку **⟳** и перезагрузите модем. 

**Получение сообщений:**
> Нажмите кнопку "**Получить сообщения**" в меню.

> **Готово**. Сообщения появятся в окне и сохранятся в таблицу. Можно сделать **анализ данных**.

**Отправка сообщений:**
> Во встроенной таблице с контактами выделите нужные контакты, используя ***CTRL*** И ***SHIFT*** кнопки.

> Нажмите "**Отправить**!", подтвердите. Программа **зависнет** на некоторое количество времени в зависимости от количества получателей.

**Управление контактами:**
> Поиск контактов. В поле Аргументов для поиска можно вводить части имен / номеров контактов через пробел. Если нет веденных аргументов, то показываются все корнтакты. Можно делать отрицательные аргументы. Аргументы выполняются последовательно, напрмиер: "8 7 -6" (и Enter) сначала добавит все контакты содержащие цифру 8, потом 7, а потом уберет из них все, в которых есть цифра 6. Найденные контакты отобразятся во встроенной таблице.

> Чтобы **добавить** контакт - введите *имя* и *номер* (Обязательно с **+7**) в соответствующие поля и нажмите "**Добавить контакт**".

> Чтобы **удалить** контакты, выделите их и нажмите "**Удалить**".



**Другое:**
> чтобы открыть **папку с файлами** - кнопка **ⓘ**

> Чтобы сменить (что?)

> vpn: vpn://AAAHWHjanVXfc6IwEH7vX-Ewfaun_AiCzvTBqj2rd4pS29rScSLEmhMDB9GqHf_3SwJFnOMeesDD8n3fLrubJfm4KLFLcgNCISYoiqVG6UVg_PrILKGC72-MPgcF0VUYLilyTa4rJjCBVC7QqEJjKrJRN2uyUajRuAaYNd1UmbJQArhEUzRDB0a9XiTpuSJKIbWGO07qcjGLSVJIEWuLInW9kBPF6YVF-TCmM9bgBebdkz4cwmGHNc1h706ubY5Uzkg1JbN-5UktIbNG5TmQcFmHclzPTf3yGOtIguryGYzJZ3o52E5z1vU8mOaq53N0fYwIvfMS6lfLvMfRfuW_WWQhu21rFz4EQGnLUB_C6jSkj6t40QRgv4yv_woyw-FnKhWzolTOMkolYYS3sxXaJ8L94_vTW9SZA92t99AzUZv1XndrWFhbDeryZjKVNUxJeDOcdwo-F27mp1D_m7lY7yTEyx2hKFpAF706Dml6XoTiuHRdysqpaioj2gObgZfW-O5nczydsddy6dLutIaDdvrORBarE1LUR3um_VKZDum5zAdwgy2u-L6w4Y7ZOrdthVs6t1RuGczqKkKZjShDOJebS4ZoPHA2jAwADDhNoEPY82IhFPEGWJu5j92kAmNrkJHaetrief0puO_3x-rt4mbro9_D77Wr3QM6HIDW33i70fRaVI_iJYyQl3jbo9ry1lV-uLc9-tyh4y6djFxzfvXrQEB1rd4D42FyJc8nNg24d9P3g3fk3Vm8-XJF3FW5XGo0qrz6DvHCABPKa6lVNLkC2NpockOTDY3XYLHdEceUTUgfoRD6eIuYVOXtyi38MojpAK5R-qflAuVEa7pJJ1ozajk8DCLKCfHJExqvTvP4pZpPkWMUbVF0Ptpfar7kkKNDivY3nrTYk3nSRQIaQRJzFftLAxpw7cYLpTPh8dzvdCBxNVwTdMDwGz9_MtlRWK-Jn-ShBdz4tPUvt0wWuxEOKQ7ERm-LrpSUjCZxcpJVxJ2Dk8NLzEwGfy61OG1yCy1dHC_-AAy947A

> https://drive.google.com/file/d/1h5YqfqMMGxAmI59vGxU1hPCyhnM3qJa_/view?usp=sharing

> https://drive.google.com/file/d/1qOO1Fqd4IB4WnpBOlY5jghbnCCLAQz9N/view?usp=drive_link
