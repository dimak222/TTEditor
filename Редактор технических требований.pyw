#-------------------------------------------------------------------------------
# Author:      dimak222
#
# Created:     18.01.2022
# Copyright:   (c) dimak222 2022
# Licence:     No
#-------------------------------------------------------------------------------

title = "TTEditor" # наименование приложения
ver = "v1.5.3.0b1" # версия файла

#------------------------------Настройки!---------------------------------------

use_txt_file = True # использовать txt файл (True - да; False - нет)

dict_settings = {
"check_update" : [True, '# проверять обновление программы ("True" - да; "False" или "" - нет)'],
"beta" : [False, '# скачивать бета версии программы ("True" - да; "False" или "" - нет)'],
} # словарь с настройками программы

#------------------------------Импорт модулей-----------------------------------

import psutil # модуль вывода запущеных процессов
import os # работа с файовой системой

from threading import Thread # библиотека потоков
import tkinter as tk # модуль окон
import tkinter.messagebox as mb # окно с сообщением

#-------------------------------------------------------------------------------

def DoubleExe(): # проверка на уже запущеное приложение

    global program_directory # значение делаем глобальным

    list = [] # список найденых названий программы

    filename = psutil.Process().name() # имя запущеного файла
    filename2 = title + ".exe" # имя запущеного файла

    if filename == "python.exe" or filename == "pythonw.exe": # если программа запущена в IDE/консоли
        pass # пропустить

    else: # запущено не в IDE/консоли

        for process in psutil.process_iter(): # перебор всех процессов

            try: # попытаться узнать имя процесса
                proc_name = process.name() # имя процесса

            except psutil.NoSuchProcess: # в случае ошибки
                pass # пропускаем

            else: # если есть имя
                if proc_name == filename or proc_name == filename2: # сравниваем имя
                    list.append(process.cwd()) # добавляем в список название программы
                    if len(list) > 2: # если найдено больше двух названий программы (два процесса)
                        Message("Приложение уже запущено!") # сообщение, поверх всех окон и с автоматическим закрытием
                        exit() # выходим из программы

    if list == []: # если нет найденых названий программы
        program_directory = os.path.dirname(os.path.abspath(__file__)) # директория рядом с программой

    else: # если путь найден
        program_directory = os.path.dirname(psutil.Process().exe()) # директория программы

    program_directory = program_directory.replace("\\", "//", 1) # замена на слеш первого символа (при "\\192.168....")

def KOMPASexe():# проверка на запущеный КОМПАС-3D, с отключённым консольным окном "CREATE_NO_WINDOW"

    import subprocess # модуль вывода запущеных процессов
    from sys import exit # для выхода из приложения без ошибки

    CREATE_NO_WINDOW = 0x08000000 # отключённое консольное окно
    processes = subprocess.Popen('tasklist', stdin=subprocess.PIPE, stderr=subprocess.PIPE, stdout=subprocess.PIPE, creationflags=CREATE_NO_WINDOW).communicate()[0] # список всех процессов
    processes = processes.decode('cp866') # декодировка списка

    processes = processes.casefold() # приводим список процессов к единому регистру (до v21 KOMPAS.Exe, в v21 KOMPAS.exe)

    if processes.find("kompas.exe") == -1: # если КОМПАС не запущен выдать сообщение
        Message("Откройте КОМПАС-3D!") # сообщение, поверх всех окон с автоматическим закрытием
        exit() # выходим из програмы

def KompasAPI(): # подключение API компаса

    from win32com.client import Dispatch, gencache # библиотека API Windows
    from sys import exit # для выхода из приложения без ошибки

    try: # попытаться подключиться к КОМПАСу

        global KompasAPI7 # значение делаем глобальным
        global iApplication # значение делаем глобальным
        global iKompasObject # значение делаем глобальным

        KompasConst3D = gencache.EnsureModule("{2CAF168C-7961-4B90-9DA2-701419BEEFE3}", 0, 1, 0).constants # константа 3D документов
        KompasConst2D = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants # константа 2D документов
        KompasConst = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants # константа для скрытия вопросов перестроения

        KompasAPI5 = gencache.EnsureModule('{0422828C-F174-495E-AC5D-D31014DBBE87}', 0, 1, 0) # API5 КОМПАСа
        iKompasObject = Dispatch('Kompas.Application.5', None, KompasAPI5.KompasObject.CLSID) # интерфейс API КОМПАС

        KompasAPI7 = gencache.EnsureModule('{69AC2981-37C0-4379-84FD-5DD2F3C0A520}', 0, 1, 0) # API7 КОМПАСа
        iApplication = Dispatch('Kompas.Application.7') # интерфейс приложения КОМПАС-3D.

        iKompasDocument = iApplication.ActiveDocument # делаем активный открытый документ

        iDocuments = iApplication.Documents # интерфейс для открытия документов

        if iApplication.Visible == False: # если компас невидимый
            iApplication.Visible = True # сделать КОМПАС-3D видемым

    except: # если не получилось выдать сообщение

        Message("КОМПАС-3D не найден!\nУстановите или переустановите КОМПАС-3D!") # сообщение, поверх всех окон с автоматическим закрытием
        exit() # выходим из програмы

def Resource_path(relative_path): # для сохранения картинки внутри exe файла

    try: # попытаться определить путь к папке
        base_path = sys._MEIPASS # путь к временной папки PyInstaller

    except Exception: # если ошибка
        base_path = os.path.abspath(".") # абсолютный путь

    return os.path.join(base_path, relative_path) # объеденяем и возващаем полный путь

def Message(text = "Ошибка!", counter = 4): # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия)

    def Message_Thread(text, counter): # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия)

        if counter == 0: # время до закрытия окна (если 0)
            counter = 1 # закрытие через 1 сек
        window_msg = tk.Tk() # создание окна
        try: # попытаться использовать значок
            window_msg.iconbitmap(default = Resource_path("cat.ico")) # значок программы
        except: # если ошибка
            pass # пропустить
        window_msg.attributes("-topmost", True) # окно поверх всех окон
        window_msg.withdraw() # скрываем окно "невидимое"
        time = counter * 1000 # время в милисекундах
        window_msg.after(time, window_msg.destroy) # закрытие окна через n милисекунд
        if mb.showinfo(title, text, parent = window_msg) == "": # информационное окно закрытое по времени
            pass # пропустить
        else: # если не закрыто по времени
            window_msg.destroy() # окно закрыто по кнопке
        window_msg.mainloop() # отображение окна

    msg_th = Thread(target = Message_Thread, args = (text, counter)) # запуск окна в отдельном потоке
    msg_th.start() # запуск потока

    msg_th.join() # ждать завершения процесса, иначе может закрыться следующие окно

def Resource_path(relative_path): # для сохранения картинки внутри exe файла

    import os # работа с файовой системой

    try: # попытаться определить путь к папке
        base_path = sys._MEIPASS # путь к временной папки PyInstaller

    except Exception: # если ошибка
        base_path = os.path.abspath(".") # абсолютный путь

    return os.path.join(base_path, relative_path) # объеденяем и возващаем полный путь

def Kompas_message(text): # сообщение в окне компаса если он открыт

    if iApplication.Visible == True: # если компас видимый
        iApplication.MessageBoxEx(text, 'Message:', 64) # сообщение

    else: # если компас невидимый
        Message(text) # сообщение, поверх всех окон с автоматическим закрытием

def Txt_file(): # считываем значения настроек из txt файла

    import os # работа с файовой системой

    name_txt_file = os.path.join(local_path, title + ".txt") # название текстового файла

    if os.path.exists(name_txt_file): # если есть txt файл использовать его

        txt_file = open(name_txt_file, encoding = "utf-8") # открываем файл записи (w+), для чтения (r), (невидимый режим)
        lines = txt_file.readlines()  # прочитать все строки
        txt_file.close() # закрываем файл

        text_processing(lines, name_txt_file) # обработка текста (строки текста, путь к файлу)

        Parameters() # присвоене значений прочитаных параметров

        print(parameters)

    else: # если нет файла
        to_create_txt_file(name_txt_file) # создать txt файл с записью в него значений (путь к txt файлу)

def text_processing(lines, Path): # обработка текста (строки текста, путь к файлу)

    import re # модуль регулярных выражений
    from sys import exit # для выхода из приложения без ошибки

    global parameters # делаем глобальным список с параметрами

    if lines == []: # если в файле нет записи, вписываем в файл опции для редактирования и инструкцию по использованию
        to_create_txt_file(Path) # создать txt файл с записью в него значений

    else: # если есть текст обратобать его

        try: # попытаться обработать значения в txt файле

            lines_clean = clearing_the_list(lines) # очистка списка строк от "#"

            for line in lines_clean: # для каждой строки производим обработку
                parameter = line.split("=") # делим по "="
                parameter[0] = parameter[0].strip() # убираем пробелы по бокам
                parameter[1] = parameter[1].strip().strip('"') # убираем пробелы и "..." по бокам

                if parameter[1].find("True") != -1: # если есть параметр со словом True, обрабатываем его
                    parameter[1] = True # присвоем значение "True"

                elif parameter[1].find("False") != -1 or parameter[1].strip() == "": # если есть параметр со словом False или "", обрабатываем его
                    parameter[1] = False # присвоем значение "False"

                elif parameter[1].find(";") != -1: # если есть параметр с ";", обрабатываем его
                    parameter[1] = parameter[1].split(";") # разделяем параметр по ";", создаёться список

                try: # пытаемся добавить параметры в словарь
                    parameters[parameter[0]] = parameter[1] # добавляем в словарь параметры

                except NameError: # если нет словаря создаём его и добавляем параметры
                    parameters = {} # создаём словарь с параметрами
                    parameters[parameter[0]] = parameter[1] # добавляем в словарь параметры

        except:
            Message("Проверте правильность записи файла: \n\"" + Path + "\"\nИли удалите его, новый будет создан автоматически.") # сообщение, поверх всех окон с автоматическим закрытием
            exit() # выходим из програмы

def clearing_the_list(lines): # очистка списка строк от "#"

    import re # модуль регулярных выражений

    lines_clean = [] # список строк с чистым текстом (без текста после "#")

    for line in lines: # для каждой строки производим обработку

        if line.isspace(): # если пустая строка пропустить
            continue

        ignore_grid = re.findall("\".[^\"]+?#.+?\"", line) # проверка строки на содержание текста с решоткой в ковычках "***#***"

        if ignore_grid != []: # если строка содержит текст с решоткой в ковычках "***#***"

            line = line.replace(ignore_grid[0], "Текст с решоткой в ковычках!=)", 1) # заменяем текст с решоткой в ковычках "***#***" на "|"

            line_clean = line.split("#", 1)[0] # если в строке есть "#" не записывать всё что после неё

            line_clean = line_clean.replace("Текст с решоткой в ковычках!=)", ignore_grid[0], 1) # заменяем "|" на текст с решоткой в ковычках "***#***"

        else: # не содержит текст с решоткой в ковычках "***#***"
            line_clean = line.split("#", 1)[0] # если в строке есть "#" не записывать всё что после неё

        if line_clean.strip() == "": # если нет записи до #, пропустиь строку
            continue

        lines_clean.append(line_clean) # список строк с чистым текстом (без текста после "#")

    return lines_clean # возврящаем чистые строки

def to_create_txt_file(name_txt_file): # создать txt файл с записью в него значений

    import os # работа с файовой системой
    from sys import exit # для выхода из приложения без ошибки

    txt_file = open(name_txt_file, "w+", encoding = "utf-8") # открываем файл записи (w+), для чтения (r), (невидимый режим)

    txt = """# txt файл настроек программы и файлы ТТ создаются в папке "C:\Program Files\ASCON\"
server_path = "" # путь к папке с txt файлами на сервере (для копирования файла рядом спрограммой), "False" или "" - использует файлы только из папки ASCON
#-------------------------------------------------------------------------------
more_options = "False" # добавление доп. файлов через ";", пример: "ТТ МЧ : ТТ МЧ.txt; ТТ Л3 : ТТ Л3.txt" (имя кнопки : Название txt файла). "False" или "" - не добавлять
on_top = "True" # запуск программы поверх всех окон, "False" или "" - выключить
msg_server = "True" # выдавать сообщения если на сервере нет txt файлов, "False" или "" - не выдавать
#last_choice = "True" # запоминать изменения, "False" или "" - не запоминать (Настройка пока не реализована!)
last_choice_file = "ТТ.txt" # файл открываемый при запуске програмы
window_size = "800; 700; 200; 200" # размер и положение окна (ширина; высота; положение окна по X; положение окна по Y)
""" # текст записываемый в .txt файл

    txt_file.write(txt) # записываем текст в файл
    txt_file.close() # закрываем файл

    os.startfile(name_txt_file) # открываем файл в системе
    Message("Введите необходимые значения! \nИ запустите приложение повторно.") # сообщение с названием файла
    exit() # выходим

def Parameters(): # присвоене значений прочитаных параметров

    global server_path # значение делаем глобальным
    global on_top # запуск программы поверх всех окон
    global more_options # добавление доп. файлов
    global msg_server # значение делаем глобальным
    global last_choice # значение делаем глобальным
    global last_choice_file # значение делаем глобальным
    global window_size # значение делаем глобальным

    server_path = parameters.setdefault("server_path", False) # путь к файлам на сервере
    on_top = parameters.setdefault("on_top", True) # запуск программы поверх всех окон
    more_options = parameters.setdefault("more_options", False) # добавление доп. файлов
    msg_server = parameters.setdefault("msg_server", True) # запоминать изменения
    last_choice = parameters.setdefault("last_choice", True) # выдавать сообщения если на сервере нет .txt файлов
    last_choice_file = parameters.setdefault("last_choice_file", "ТТ.txt") # если файл вписан он откроется при запуске программы
    window_size = parameters.setdefault("window_size", [800, 700, 200, 200]) # чтение парамотров размера и положения окна

def comparing_and_copying_files(server_path, msg_server): # сравнение txt файлов с сервера и рядом с программой и их копирование (путь к файлам на сервере, выдавать сообщения если на сервере нет .txt файлов)

    import os # работа с файовой системой
    import filecmp # модуль сравнения файлов
    from send2trash import send2trash # модуль для удаления файлов в корзину
    import shutil # библиотека для копирования/перемещений/переименований

    if server_path: # если путь к файлам на сервере указан

        for file_name in options.values(): # для каждого значения из списка взять значение

            file_server_path = os.path.join(server_path, file_name) # полный путь к файлу на сервере
            file_full_name = os.path.join(local_path, file_name) # полный путь к файлу в папке ASCON

            if os.path.exists(file_server_path): # если есть файл на сервере

                if os.path.exists(file_full_name): # если есть файл в папке ASCON

                    if filecmp.cmp(file_server_path, file_full_name, shallow = True) == False: # если файлы разные, обработать (сравнивает только метаданные файлов)

                        send2trash(file_full_name) # старый файл удаляем в корзину
                        shutil.copy2(file_server_path, file_full_name) # копируем файл с сервера с сохранением методанных
                        Message("Файл \"" + file_name + "\" обновлён!") # сообщение, поверх всех окон с автоматическим закрытием

                else: # нет файла в папке ASCON
                    shutil.copy2(file_server_path, file_full_name) # копируем файл с сервера с сохранением методанных
                    Message("Файл \"" + file_name + "\" скопирован с сервера!", 2) # сообщение, поверх всех окон с автоматическим закрытием

            else: # нет файла на сервере
                if msg_server: # выдавать сообщения если на сервере нет .txt файлов
                    Message("Файл \"" + file_name + "\" на сервере не найден!") # сообщение, поверх всех окон с автоматическим закрытием

def More_options(): # добавление доп. файлов

    if more_options: # если опция есть

        if type(more_options) == list: # если значение параметра список

            for option in more_options: # обрабатываем каждую опцию
                option = option.split(":") # разделяем опцию по ":"
                options[option[0].strip()] = option[1].strip() # добавляем в словарь опцию и убираем пробелы

        else: # не список
            option = more_options.split(":") # разделяем опцию по ":"
            options[option[0].strip()] = option[1].strip() # добавляем в словарь опцию и убираем пробелы

def Last_choice_file(): # файл открываемый при запуске програмы

    import os # работа с файовой системой

    global last_choice_file # значение делаем глобальным

    if last_choice_file == False: # если параметр выключен

        last_choice_file = "ТТ.txt" # файл открывающийся при запуске программы

    if os.path.exists(os.path.join(local_path, last_choice_file)) == False: # если файла нет использовать стандардный

        Message("Файл \"" + last_choice_file + "\" не найден! \nИспользуется ТТ.txt") # сообщение, поверх всех окон с автоматическим закрытием

        last_choice_file = "ТТ.txt" # файл открывающийся при запуске программы

    return last_choice_file # возврящаем название txt файла

def TT_file(txt): # считываем ТТ с txt файла

    import os # работа с файовой системой
    from sys import exit # для выхода из приложения без ошибки

    name_txt_file = os.path.join(local_path, txt) # название текстового файла

    if os.path.exists(name_txt_file) == False: # если нет txt файла

        to_create_TT_file(name_txt_file) # создать txt файл с записью в него значений
        Message("Создан пустой файл " + txt + " , заполните его!", 8) # сообщение, поверх всех окон с автоматическим закрытием

    print("Используется: " + name_txt_file)
    txt_file = open(name_txt_file, encoding = "utf-8") # открываем файл с кодировкой
    data = txt_file.read() # считываем весь файл
    txt_list = data.split('\n') # делим на строки
    txt_file.close() # закрываем файл

    return txt_list # возвращаем считаный список

def to_create_TT_file(name_txt_file): # создать txt файл с записью в него значений

    import os # работа с файовой системой
    from sys import exit # для выхода из приложения без ошибки

    txt_file = open(name_txt_file, "w+", encoding = "utf-8") # открываем файл записи (w+), для чтения (r), (невидимый режим)

    txt = """# заполните файл для отображения строк ТТ в программе. Значения после "#" работают как комментарии (не передаются в ТТ чертежа).
# для перевода текста на новую строку без нумерации, испольуйте "\\n". Пример: "Редактор\\nтехнических\\nтребований"
------------------------------------------------------------------------------------------- # для визуального отделения ТТ в программе можно использовать "----"
""" # текст записываемый в .txt файл

    txt_file.write(txt) # записываем текст в файл
    txt_file.close() # закрываем файл

    os.startfile(name_txt_file) # открываем файл в системе

def Window(): # формирование окна программы

    import tkinter as tk # модуль окон

    global window # делаем глобальным

    window = tk.Tk() # создание окна
    window.iconbitmap(default = Resource_path("cat.ico")) # значок программы
    window.title(title) # заголовок окна
    window.geometry("%dx%d+%d+%d" % (int(window_size[0]), int(window_size[1]), int(window_size[2]), int(window_size[3]))) # размер окна и его положение
    window.resizable(width = True, height = True) # возможность менять размер окна
    window.attributes("-topmost", on_top) # окно поверх всех окон

    f_top = tk.Frame(window) # блок окна (вверх)
    f_senter = tk.LabelFrame() # блоки окна с рамкой (центр)
    f_bot = tk.Frame(window) # блок окна (низ)

    listbox_function(f_top) # отображения списока в программе

    checkbox(f_senter) # радиокнопка с выбором файла для считывания (его роложение)

    button_function(f_bot, "Удалить последний пункт", clear_line, "del1.png", 20) # кнопка (положение(frame), текст, действие, путь к картинке, маштаб картинки)

    button_function(f_bot, "Очистить ТТ", clear, "del.png", 20) # кнопка (положение(frame), текст, действие, путь к картинке, маштаб картинки)

    button_function(f_bot, "Редактировать файл", Edit_file, "pen.png", 20) # кнопка (положение(frame), текст, действие, путь к картинке, маштаб картинки)

    button_function(f_bot, "Настройки", Settings, "settings.png", 20) # кнопка (положение(frame), текст, действие, путь к картинке, маштаб картинки)

    button_function(f_bot, "", On_top, "stud button.png", 19) # кнопка (положение(frame), текст, действие, путь к картинке, маштаб картинки)

    f_top.pack(expand = True, fill = "both") # размещение блока (с возможностью расширяться и заполненем окна во всех направлениях)
    f_senter.pack() # размещение блока
    f_bot.pack() # размещение блока

    window.mainloop() # формирование окна

def listbox_function(frame): # отображения списока в программе (окно)

    import tkinter as tk # модуль окон

    global listbox # делаем глобальным что бы можно было перечитать через радиокнопку

    scrollbarY = tk.Scrollbar(frame, orient = "vert") # скроллбар
    scrollbarY.pack(side = "right", fill = tk.Y, padx = 1, pady = 1,) # положение скроллбара (с отступом по x и y)

    scrollbarX = tk.Scrollbar(frame, orient = "hor") # скроллбар
    scrollbarX.pack(side = "bottom", fill = tk.X, padx = 1, pady = 1) # положение скроллбара (с отступом по x и y)

    listbox = tk.Listbox(frame, yscrollcommand = scrollbarY.set, xscrollcommand = scrollbarX.set, width = 27, height = 10, font = ('arial', 13)) # шрифт и размер тектса в окне программы
    listbox.bind('<<ListboxSelect>>', select_item) # действие при выбранной строчке
    listbox.pack(expand = True, fill = "both") # размер списка в окне (с возможностью расширяться и заполненем окна во всех направлениях)

    for item in txt_list: # добавление текста в список построчно
        listbox.insert(tk.END, item) # добавляем строку в конец списка

    scrollbarY.configure(command = listbox.yview) # для перемещения тектса скролбаром
    scrollbarX.configure(command = listbox.xview) # для перемещения тектса скролбаром

def checkbox(frame): # радиокнопка с выбором файла для считывания (его роложение)

    import tkinter as tk # модуль окон

    global var # делаем глобальным что бы можно было перечитать через радиокнопку

    var = tk.StringVar(frame) # переменная для чтения txt файла
    var.set(txt) # выбор точки на радиокопке

    for (NameRadiobutton, val) in options.items(): # для каждого значения из списка создать радиокнопку
        Radiobutton = tk.Radiobutton(frame, text = NameRadiobutton, variable = var, value = val) # создание радиокнопки
        Radiobutton.configure(command = updating_list) # команда выполняемая при нажатии кнопки
        Radiobutton.pack(side = "left", ipady = 0) # положение радиокнопки

def updating_list(): # обновление списка

    import tkinter as tk # модуль окон

    comparing_and_copying_files(server_path, msg_server = False) # сравнение txt файлов с сервера и рядом с программой и их копирование

    txt_list = TT_file(var.get()) # считывание с txt файла

    listbox.delete(0, listbox.size()) # удаляем весь список
    for item in txt_list: # добавление текста в список построчно
        listbox.insert(tk.END, item) # добавляем строку в конец списка

def select_item(event): # характеристики текста ТТ

    line = listbox_text_processing() # обработка текста списка (не вводим лишнее в ТТ)

    if type(line) == list: # если значение параметра список, обработать каждое значение
        namber_new_line = len(line) - 1 # количество новых строк (n - 1)
        line = "\n".join(line) # строки разделённые знаком переноса

    else:
        namber_new_line = False # нет количества новых строк

    if line: # если есть строка использеум

        iKompasDocument = iApplication.ActiveDocument # делаем активный открытый документ

        if iKompasDocument == None or iKompasDocument.DocumentType != 1: # если нет открытого документа выдать сообщение
            Kompas_message("Откройте чертёж!") # сообщение в окне компаса если он открыт

        else:
            iDrawingDocument = KompasAPI7.IDrawingDocument(iKompasDocument) # интерфейс чертежа
            iTechnicalDemand = iDrawingDocument.TechnicalDemand # интерфейс технических требований
            iText = iTechnicalDemand.Text # интерфейс текста ТТ
            #iText.Clear() # очистить ТТ

            iTextLine = iText.Add() # указатель на интерфейс ITextLine текста (добавляет новую строку в ТТ)
            iTextLine.Align = -1 # выравнивание (0 - слева,  1 - по центру, 2 - справа, 3 - во всю ширину, 4 - по десятич­ной точке, -1 - по умолчанию (из стиля)
##            iTextLine.IndentedLine = 5 # cмещение красной строки
##            iTextLine.LeftEdge = 0 # отступ текста слева
            iTextLine.Level = 0 # уровень вложенности нумерации
            iTextLine.Numbering = 1 # тип нумерации абзаца (-1 - Тип не определенный; 0 - Cрока без нумерации; 1 - Cтрока с нумерацией уровня level; 2 - На строке начинается новая нумерация пунктов; 3 - Cтрока не должна нумероваться никогда)
##            iTextLine.RightEdge = 0 # отступ текста справа
##            iTextLine.Step = 10 # шаг строк
##            iTextLine.StepAfterParagraph = 0 # дополнительный шаг после абзаца
##            iTextLine.StepBeforeParagraph = 0 # дополнительный шаг перед абзаца
##            iTextLine.Style = 2 # системный стиль текста (2 - Текст для технических требований)
##            iTextLine.Delete() # удалить строку текста (строку с текстом)

            iTextItem = iTextLine.Add() # указатель на интерфейс ITextItem (добавить компонент строки в конец строки)
            iTextItem.ItemType = 0 # тип компонента текста (0 - cтрока, остальное см. ksTextItemEnum)
##            iTextItem.NewLine = True # признак начала строки
            iTextItem.Str = line # текстовое значение компоненты текста (сам текст из списка)
##            iTextItem.SymbolFontName = '' # имя шрифта для символа
##            iTextItem.Delete() # удалить компоненту строки (только текст)
##
##            iTextFont = KompasAPI7.ITextFont(iTextItem) # интерфейс параметров шрифта
##            iTextFont.Bold = False # жирный шрифт
##            iTextFont.Color = 0 # цвет
##            iTextFont.FontName = 'GOST type A' # имя шрифта
##            iTextFont.Height = 5 # высота текта
##            iTextFont.Italic = True # курсив
##            iTextFont.Underline = False # подчёркивание
##            iTextFont.WidthFactor = 1 # коэффициент сужения
            iTextItem.Update() # обновить данные компонента
            iTechnicalDemand.Update() # обновить данные ТТ

    if namber_new_line: # если есть новые строки

        Count = iText.Count # количество строк

        for n in range(1, namber_new_line + 1): # для каждой строки

            iTextLine = iText.TextLine(Count - n) # указываем на последнюю строку (отсчёт с 0)
            iTextLine.Numbering = 3 # тип нумерации абзаца (-1 - Тип не определенный; 0 - Cрока без нумерации; 1 - Cтрока с нумерацией уровня level; 2 - На строке начинается новая нумерация пунктов; 3 - Cтрока не должна нумероваться никогда)

        iTechnicalDemand.Update() # обновить данные ТТ

def listbox_text_processing(): # обработка текста списка (не вводим лишнее в ТТ)

    import re # модуль регулярных выражений

    line = (listbox.get(listbox.curselection())) # строка из списка

    if re.findall("--.+--", line) == []: # если строка не "-----", то используем её

        line = line.split("#", 1)[0] # если есть "#" не записывать всё что после неё

        line = line.strip() # убираем пробелы по бокам

        if line.find("\\n") != -1: # если есть параметр переноса строки, обработать его
            line = line.split("\\n") # создаём список строк

        return line # выводим строку

    else:
        return False # ничего не вывадим

def button_function(frame, text, action, image = None, m = 1): # кнопка (положение(frame), текст, действие, путь к картинке, маштаб картинки)

    import tkinter as tk # модуль окон

    if image != None: # если путь к картинке указан
        image = tk.PhotoImage(file = Resource_path(image)) # создание объекта изображения
        image = image.subsample(m, m) # мастаб картинки

    button = tk.Button(frame, # действие кнопки, её размеры, шрифт
                 text = text, # текст на кнопке
                 width = 0, height = 0, # размеры кнопки (0 р-р - по тектсу)
                 image = image, # картинка в кнопке
                 compound = "left", # положение картинки
                 font = ('arial', 11)) # шрифт текста
    button.bind("<Button-1>", action) # действие кнопки
    button.image = image # для отображения картинки (из-за PhotoImage которая находиться в функции она не отображаеться)
    button.pack(side = "left", pady = 2) # размер отступа вокруг кнопки

def clear_line(event): # очистка послетней строки ТТ

    from threading import Thread # модуль потоков

    iKompasDocument = iApplication.ActiveDocument # делаем активный открытый документ

    if iKompasDocument == None: # пока нет активного 2D документа выдавать сообщение
        Kompas_message("Откройте чертёж!") # сообщение в окне компаса если он открыт

    else:
        iDrawingDocument = KompasAPI7.IDrawingDocument(iKompasDocument) # интерфейс чертежа
        iTechnicalDemand = iDrawingDocument.TechnicalDemand # интерфейс технических требований
        iText = iTechnicalDemand.Text # интерфейс текста ТТ
        Count = iText.Count # количество строк
        iTextLine = iText.TextLine(Count - 1) # указываем на последнюю строку (отсчёт с 0)

        iTextItem = iTextLine.TextItem(0) # указываем на последнюю строку (отсчёт с 0)
        NewLine = iTextItem.NewLine # признак начала строки

        iTextLine.Delete() # удалить строку текста (всю строку с текстом)
        iTechnicalDemand.Update() # обновить данные ТТ

        while NewLine == False: # если строка не начало, удаляем её

            Count = iText.Count # количество строк
            iTextLine = iText.TextLine(Count - 1) # указываем на последнюю строку (отсчёт с 0)

            iTextItem = iTextLine.TextItem(0) # указываем на последнюю строку (отсчёт с 0)
            NewLine = iTextItem.NewLine # признак начала строки
            iTextLine.Delete() # удалить строку текста (всю строку с текстом)
            iTechnicalDemand.Update() # обновить данные ТТ

def clear(event): # очистка ТТ

    iKompasDocument = iApplication.ActiveDocument # делаем активный открытый документ

    if iKompasDocument == None: # пока нет активного 2D документа выдавать сообщение
        Kompas_message("Откройте чертёж!") # сообщение в окне компаса если он открыт

    else:
        iDrawingDocument = KompasAPI7.IDrawingDocument(iKompasDocument) # интерфейс чертежа
        iTechnicalDemand = iDrawingDocument.TechnicalDemand # интерфейс технических требований
        iText = iTechnicalDemand.Text # интерфейс текста ТТ
        iText.Clear() # очистить ТТ
        iTechnicalDemand.Update() # обновить данные ТТ

def Edit_file(event): # открыть .txt файл для редактирования

    import os # работа с файовой системой

    if server_path: # если путь к файлам на сервере указан

        file_server_path = os.path.join(server_path, var.get()) # полный путь к файлу на сервере

        if os.path.exists(file_server_path): # если есть файл
            os.startfile(file_server_path) # открываем файл в системе

        else:
            Message("Открыт локальный файл!", 8) # сообщение, поверх всех окон с автоматическим закрытием

    else: # путь к файлам на сервере не указан
        os.startfile(os.path.join(local_path, var.get())) # открываем файл в системе

def Settings(event): # открыть файл с настройками программы

    import os # работа с файовой системой

    print(os.path.join(local_path, title + ".txt"))
    os.startfile(os.path.join(local_path, title + ".txt")) # открываем файл в системе

    Message("Введите необходимые значения! \nИ запустите приложение повторно.") # сообщение с названием файла

def On_top(event): # поверх всех окон

    global on_top # значение делаем глобальным

    if on_top: # если включено

        window.attributes("-topmost", False) # выключаем окно поверх всех окон
        on_top = False # значение выключено

    else: # если выключено
        window.attributes("-topmost", True) # включаем окно поверх всех окон
        on_top = True # значение включено

#-------------------------------------------------------------------------------

local_path = "C:\Program Files\ASCON\\" # путь к локальной папке хранения txt файлов

options = {"Дет." : "ТТ.txt", # название радиокнопок и список txt файлов к ним
		   "СБ" : "ТТ СБ.txt",
		   "УЧ" : "ТТ УЧ.txt",
           "Опт." : "ТТ Опт..txt"}

DoubleExe() # проверка на уже запущеное приложение, с отключённым консольным окном "CREATE_NO_WINDOW"

##KOMPASexe() # проверка на запущеный КОМПАС-3D
KompasAPI() # подключение API компаса

if use_txt_file: # использовать txt файл
    Txt_file() # считываем значения настроек из txt файла

More_options() # добавление доп. файлов

comparing_and_copying_files(server_path, msg_server) # сравнение txt файлов с сервера и рядом с программой и их копирование (путь к файлам на сервере, выдавать сообщения если на сервере нет .txt файлов)

txt = Last_choice_file() # файл открываемый при запуске програмы
txt_list = TT_file(txt) # считываем ТТ с txt файла

Window() # формирование окна программы