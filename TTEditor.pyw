#-------------------------------------------------------------------------------
# Author:      dimak222
#
# Created:     18.01.2022
# Copyright:   (c) dimak222 2022
# Licence:     No
#-------------------------------------------------------------------------------

title = "TTEditor" # наименование приложения
ver = "v2.0.0.0" # версия файла

#------------------------------Настройки!---------------------------------------

parameters = {
'check_update' : 'True # проверять обновление программы (True - да; False или "" - нет)',
'beta' : 'False # скачивать бета версии программы (True - да; False или "" - нет)\n',

'server_path' : '"" # путь к папке с txt файлами на сервере (для копирования файла рядом спрограммой), False или "" - использует файлы только рядом с программой\n',

'more_options' : 'False # добавление доп. файлов через ";", пример: "ТТ МЧ : ТТ МЧ.txt; ТТ Л3 : ТТ Л3.txt" (Имя кнопки : Название txt файла). False или "" - не добавлять\n',

'on_top' : 'True # запуск программы поверх всех окон, False или "" - выключить',
'msg_server' : 'True # выдавать сообщения если на сервере нет txt файлов, False или "" - не выдавать\n',

'last_choice' : 'True # запоминать изменения, False или "" - не запоминать (Настройка пока не реализована!)',
'last_choice_file' : 'ТТ.txt # файл открываемый при запуске програмы',
'window_size' : '900; 600; 150; 150 # размер и положение окна (ширина; высота; положение окна по X; положение окна по Y)',
}

#------------------------------Импорт модулей-----------------------------------

import psutil # модуль вывода запущеных процессов
import os # работа с файовой системой

import configparser # модуль для работы с INI-файлами
import ast # модуль преобразования переменных

from sys import exit # для выхода из приложения без ошибки

##import pythoncom # модуль для запуска без IDE
from win32com.client import Dispatch, gencache # библиотека API Windows
from pythoncom import connect # подключаемся к запущенному экземпляру КОМПАСа

import filecmp # модуль сравнения файлов
from send2trash import send2trash # модуль для удаления файлов в корзину
import shutil # библиотека для копирования/перемещений/переименований

from PyQt6.QtWidgets import QApplication, QMainWindow, QMessageBox, QButtonGroup, QRadioButton
from PyQt6.QtGui import QIcon
from PyQt6.QtCore import Qt, QTimer, pyqtSignal, QEventLoop

from main_window import Ui_MainWindow # импорт GUI окна
import resources_rc  # импорт ресурсов (иконок)

import re # модуль регулярных выражений

import sys # для определения строки ошибки

import traceback # для вывода ошибок

#-------------------------------------------------------------------------------

def main(): # запуск интерфейса

    app = QApplication(sys.argv)  # новый экземпляр QApplication (sys.argv)

    window = MyApp()  # создаём объект класса ExampleApp
    window.setWindowIcon(QIcon(resource_path("cat.ico"))) # значок программы
    window.setWindowTitle(f"{title} {ver}") # заголовок окна

    window.show()  # показываем окно

    app.exec() # запускаем приложение

class MyApp(QMainWindow, Ui_MainWindow): # основное окно

    def __init__(self): # нужно для доступа к переменным, методам и т.д. в файле main_window.pyw

        super().__init__()

        self.setupUi(self)  # инициализация дизайна

        self.doubleExe() # проверка на уже запущеное приложение

        self.load_settings() # загружает настройки из INI-файла

        self.check_update() # проверить обновление приложение

        self.kompasAPI() # подключение API компаса

        self.comparing_and_copying_files(self.settings_val("msg_server")) # сравнение txt файлов с сервера и рядом с программой и их копирование (путь к файлам на сервере, выдавать сообщения если на сервере нет .txt файлов)

        self.window_settings() # применяем настройки окна

        self.setup_radio_buttons() # настройка радиокнопок

        self.det_rbtn.hide() # спрятать кнопоку из дизайна
        self.SB_rbtn.hide() # спрятать кнопоку из дизайна
        self.UCH_rbtn.hide() # спрятать кнопоку из дизайна
        self.CE_rbtn.hide() # спрятать кнопоку из дизайна
        self.optics_rbtn.hide() # спрятать кнопоку из дизайна

        self.listWidget.itemClicked.connect(self.select_item) # для работы нажатия списка

        self.clear_line_btn.clicked.connect(self.clear_line) # для работы кнопок
        self.clear_btn.clicked.connect(self.clear) # для работы кнопок
        self.edit_file_btn.clicked.connect(self.edit_file) # для работы кнопок
        self.settings_btn.clicked.connect(self.open_settings) # для работы кнопок
        self.on_top_btn.clicked.connect(self.on_top) # для работы кнопок

    def doubleExe(self): # проверка на уже запущеное приложение

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
                            self.message("Приложение уже запущено!") # сообщение, поверх всех окон и с автоматическим закрытием
                            exit() # выходим из программы

        if list == []: # если нет найденых названий программы
            program_directory = os.path.dirname(os.path.abspath(__file__)) # директория рядом с программой

        else: # если путь найден
            program_directory = os.path.dirname(psutil.Process().exe()) # директория программы

        program_directory = program_directory.replace("\\", "//", 1) # замена на слеш первого символа (при "\\192.168....")

    def load_settings(self): # загружает настройки из INI-файла

        global settings # делаем глобальным

        def load_ini_config(ini_path): # загружает параметры из INI-файла, автоматически удаляя комментарии (путь к INI-файлу)

            def create_default_ini(ini_path): # cоздаёт INI-файл с параметрами по умолчанию и комментариями (путь к INI-файлу)

                config = configparser.ConfigParser()
                config['Settings'] = parameters

                with open(ini_path, 'w', encoding='utf-8') as configfile:
                    config.write(configfile)

            config = configparser.ConfigParser(inline_comment_prefixes=("#",))

            if os.path.exists(ini_path): # если файл существует
                config.read(ini_path, encoding='utf-8')
                return config['Settings']
            else:
                # Если файл не существует, создаём его с параметрами по умолчанию
                create_default_ini(ini_path)
                return load_ini_config(ini_path)

        def convert_value(value): # функция для автоматического преобразования

            try:
                value = value.strip('"').strip() # удаляем '"' и пробелов по бокам если они есть
                return ast.literal_eval(value) # преобразуем значения

            except (ValueError, SyntaxError):
                return value  # Если преобразование невозможно, оставляем как есть

        ini_path = os.path.join(program_directory, title + ".ini") # путь к INI-файлу
        settings = load_ini_config(ini_path) # загружает параметры из INI-файла, автоматически удаляя комментарии (путь к INI-файлу)

        # Преобразование словаря
        settings = {key: convert_value(value) for key, value in settings.items()}

        # Вывод параметров для отладки
        print("Загружены параметры из INI-файла:")
        for key, value in settings.items():
            print(f"{key} = {value}")

    def check_update(self): # проверить обновление приложение

        global url # значение делаем глобальным

        if self.settings_val("check_update"): # если проверка обновлений включена

            try: # попытаться импортировать модуль обновления

                from Updater import Updater # импортируем модуль обновления

                if "url" not in globals(): # если нет ссылки на программу
                    url = "" # нет ссылки

                Updater.Update(title, ver, self.settings_val("beta"), url, program_directory, resource_path("cat.ico")) # проверяем обновление (имя программы, версия программы, скачивать бета версию, ссылка на программу, директория программы, путь к иконке)

            except SystemExit: # если закончили обновление (exit в Update)
                exit() # выходим из программы

            except: # не удалось
                msg = "Ошибка обновления!"
                print(msg)
                self.message(msg) # сообщение, поверх всех окон с автоматическим закрытием
                pass # пропустить

        else:
            print("Проверка обновлений выключена")

    def kompasAPI(self): # подключение API компаса

        try: # попытаться подключиться к КОМПАСу

            global KompasAPI7 # значение делаем глобальным
            global iApplication # значение делаем глобальным
            global iKompasObject # значение делаем глобальным

            KompasConst3D = gencache.EnsureModule("{2CAF168C-7961-4B90-9DA2-701419BEEFE3}", 0, 1, 0).constants # константа 3D документов
            KompasConst = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants # константа для скрытия вопросов перестроения и 2D документов

            KompasAPI5 = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0) # API5 КОМПАСа
            iKompasObject = KompasAPI5.KompasObject(connect('Kompas.Application.5')) # подключение к запущенному экземпляру КОМПАСа
            ##iKompasObject = Dispatch("Kompas.Application.5", None, KompasAPI5.KompasObject.CLSID) # интерфейс API КОМПАС

            KompasAPI7 = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0) # API7 КОМПАСа
            iApplication = KompasAPI7.IApplication(connect('Kompas.Application.7')) # подключение к запущенному экземпляру КОМПАСа
            ##iApplication = Dispatch("Kompas.Application.7") # интерфейс приложения КОМПАС-3D.

            if iApplication.Visible == False: # если компас невидимый
                iApplication.Visible = True # сделать КОМПАС-3D видемым

        except: # если не получилось подключиться к КОМПАСу
            msg = "Не удалось подключиться к КОМПАС-3D!\nЗапустите или перезапустите КОМПАС-3D!"
            print(msg)
            self.message(msg, 8) # сообщение, поверх всех окон с автоматическим закрытием
            exit() # выходим из программы

    def comparing_and_copying_files(self, msg_server): # сравнение txt файлов с сервера и рядом с программой и их копирование (путь к файлам на сервере, выдавать сообщения если на сервере нет .txt файлов)

        server_path = self.settings_val("server_path")

        if server_path: # если путь к файлам на сервере указан

            for file_name in options.values(): # для каждого значения из списка взять значение

                file_server_path = os.path.join(server_path, file_name) # полный путь к файлу на сервере
                file_full_name = os.path.join(program_directory, file_name) # полный путь к файлу в папке с программой

                if os.path.exists(file_server_path): # если есть файл на сервере

                    if os.path.exists(file_full_name): # если есть файл в папке с программой

                        if filecmp.cmp(file_server_path, file_full_name, shallow = True) == False: # если файлы разные, обработать (сравнивает только метаданные файлов)

                            try: # попытаться удалить в корзину
                                send2trash(file_full_name) # старый файл удаляем в корзину

                            except: # в случае ошибки
##                                self.message("Файл \"" + file_name + "\" не может быть удалён в корзину!") # сообщение, поверх всех окон с автоматическим закрытием
                                print("Файл \"" + file_name + "\" не может быть удалён в корзину!")

                            shutil.copy2(file_server_path, file_full_name) # копируем файл с сервера с сохранением методанных
                            self.message("Файл \"" + file_name + "\" обновлён!") # сообщение, поверх всех окон с автоматическим закрытием

                    else: # нет файла в папке с программой
                        shutil.copy2(file_server_path, file_full_name) # копируем файл с сервера с сохранением методанных
                        self.message("Файл \"" + file_name + "\" обновлён с сервера!", 2) # сообщение, поверх всех окон с автоматическим закрытием

                else: # нет файла на сервере
                    if msg_server: # выдавать сообщения если на сервере нет .txt файлов
                        self.message("Файл \"" + file_name + "\" на сервере не найден!") # сообщение, поверх всех окон с автоматическим закрытием

    def settings_val(self, key): # получить значение из словаря

        key = key.lower() # преобразуем ключ в нижний регистр для поиска

        value = settings.get(key) # берем переменную из словаря

        return value # возвращаем значение

    def window_settings(self): # применяет настройки размера и положения окна из INI-файла

        self.set_on_top() # установка состояния "поверх всех окон"

        window_size_str = self.settings_val("window_size") # размеры и положения окна

        if not window_size_str: # если настройка отсутствует, используем размеры по умолчанию
            return

        try:
            size_parts = [part.strip() for part in window_size_str.split(';')] # разбираем строку с настройками

            if len(size_parts) >= 4: # если прописаны все значения

                width, height, pos_x, pos_y = map(int, size_parts[:4]) # извлекаем и преобразуем значения

                self.resize(width, height) # устанавливаем размер окна

                self.move(pos_x, pos_y) # устанавливаем положение окна

                self.ensure_on_screen() # проверяем, чтобы окно не вышло за пределы экрана

        except (ValueError, IndexError) as e:
            msg = f"Некорректные настройки размера окна\nОшибка: {e}"
            print(msg)
            self.message(msg, 8) # сообщение, поверх всех окон с автоматическим закрытием

    def set_on_top(self): # установка состояния "поверх всех окон"

        if self.settings_val("on_top"): # если включено
            self.setWindowFlags(self.windowFlags() | Qt.WindowType.WindowStaysOnTopHint) # установливаем "поверх всех окон"
            # self.on_top_btn.setIcon(QIcon(":/TTEditor/stud_button_active.png")) # меняем иконку на активную

        self.show() # необходимо вызвать show() после изменения флагов

    def ensure_on_screen(self): # проверяем, чтобы окно не вышло за пределы экрана

        screen_geometry = QApplication.primaryScreen().availableGeometry()
        window_geometry = self.frameGeometry()

        if not screen_geometry.intersects(window_geometry): # если окно полностью вне экрана, перемещаем его в центр
            center_point = screen_geometry.center()
            self.move(center_point.x() - window_geometry.width() / 2,
                     center_point.y() - window_geometry.height() / 2)
        else:
            # Корректируем положение, если окно частично вне экрана
            x = max(screen_geometry.left(), min(self.x(), screen_geometry.right() - window_geometry.width()))
            y = max(screen_geometry.top(), min(self.y(), screen_geometry.bottom() - window_geometry.height()))
            self.move(x, y)

    def closeEvent(self, event): # обработчик события закрытия окна

        self.save_window_settings() # сохраняем текущие размер и положение окна
        event.accept() # подтверждаем закрытие окна

    def save_window_settings(self): # cохраняет текущие размер и положение окна

        geometry = self.geometry() # получаем текущую геометрию окна

        window_size_str = f"{geometry.width()}; {geometry.height()}; {geometry.x()}; {geometry.y()}" # формируем строку для сохранения

        self.save_settings_val("window_size", window_size_str) # сохраняем в настройки

    def setup_radio_buttons(self): # настраивает радиокнопки и связывает их с файлами

        self.radio_group = QButtonGroup(self) # создаем группу для радиокнопок

        self.create_additional_radio_buttons() # создаем кнопки

        self.radio_group.buttonClicked.connect(self.on_radio_button_clicked) # подключаем сигнал изменения выбора кнопки

    def create_additional_radio_buttons(self): # создаем кнопки

        self.more_options() # добавление доп. кнопок если они есть

        self.last_choice() # файл открываемый при запуске программы

        last_file = self.settings_val("last_choice_file") # опция последней нажатой кнопки

        for name, file_name in options.items(): # для каждого значения из списка создать радиокнопку (имя кнопки, файл)

            radio_btn = QRadioButton(name, self.groupBox) # создаем радиокнопку в готовой группе

            self.horizontalLayout_3.addWidget(radio_btn) # добавляем кнопку в layout группы

            radio_btn.setProperty("file_name", file_name) # присваеваем свойство кнопке, как имя файла

            self.radio_group.addButton(radio_btn) # добавляем кнопку в группу

            if file_name == last_file: # сравниваем по имени файла
                self.update_list(last_file) # заполняем список
                radio_btn.setChecked(True) # устанавливаем кнопку как выбранную по умолчанию

    def more_options(self): # добавление доп. файлов

        if self.settings_val("more_options"): # если опция есть

            if type(self.settings_val("more_options")) == list: # если значение параметра список

                for option in self.settings_val("more_options"): # обрабатываем каждую опцию
                    option = option.split(":") # разделяем опцию по ":"
                    options[option[0].strip()] = option[1].strip() # добавляем в словарь опцию и убираем пробелы

            else: # не список
                option = self.settings_val("more_options").split(":") # разделяем опцию по ":"
                options[option[0].strip()] = option[1].strip() # добавляем в словарь опцию и убираем пробелы

    def last_choice(self): # файл открываемый при запуске программы

        last_choice_file = self.settings_val("last_choice_file") # параметр последнего файла

        if last_choice_file == False: # если параметр выключен

            last_choice_file = "ТТ.txt" # файл открывающийся при запуске программы

        if os.path.exists(os.path.join(program_directory, last_choice_file)) == False: # если файла нет использовать стандардный

            self.message("Файл \"" + last_choice_file + "\" не найден!\nИспользуется: ТТ.txt") # сообщение, поверх всех окон с автоматическим закрытием

            last_choice_file = "ТТ.txt" # файл открывающийся при запуске программы

        self.save_settings_val("last_choice_file", last_choice_file) # сохранить значения в словарь

    def save_settings_val(self, key, val): # сохранить значения в словарь

        key = key.lower() # преобразуем ключ в нижний регистр для поиска

        settings[key] = val # присваеваем новое значение

    def on_radio_button_clicked(self, radio_btn): # обрабатываем клик по кнопке

        file_name = radio_btn.property("file_name") # получаем имя файла из свойства кнопки

        if file_name: # если есть значение
            self.update_list(file_name) # обновляем список

    def update_list(self, file_name): # обновляем список

        self.comparing_and_copying_files(msg_server = False) # сравнение txt файлов с сервера и рядом с программой и их копирование

        txt_list = self.read_TT_file(file_name) # считывание с txt файла

        self.listWidget.clear() # удаляем весь список
        self.listWidget.addItems(txt_list) # добавляем все строки из списка

    def read_TT_file(self, txt_filename): # считываем ТТ с txt файла

        file_path = os.path.join(program_directory, txt_filename) # название текстового файла

        if not os.path.exists(file_path): # если нет txt файла
            txt_list = self.to_create_TT_file(file_path) # создать txt файл с записью в него значений
            self.message(f"Создан пустой файл {txt_filename}, заполните его!", 8) # сообщение, поверх всех окон с автоматическим закрытием

        with open(file_path, 'r', encoding='utf-8') as txt_file: # используем менеджер контекста для автоматического закрытия файла
            txt_list = txt_file.read().splitlines() # читаем и сразу разбиваем на строки

        return txt_list # возвращаем считаный список

    def to_create_TT_file(self, file_path): # создать txt файл с записью в него значений

        txt = """# заполните файл для отображения строк ТТ в программе. Значения после "#" работают как комментарии (не передаются в ТТ чертежа).
    # для перевода текста ТТ на новую строку без нумерации, испольуйте "\\n". Пример: "Редактор\\nтехнических\\nтребований"
    ------------------------------------------------------------------------------------------- # для визуального отделения ТТ в программе можно использовать "----"
    """ # текст записываемый в .txt файл

        with open(file_path, "w", encoding="utf-8") as txt_file: # используем менеджер контекста для автоматического закрытия файла
            txt_file.write(txt)

        os.startfile(file_path) # открываем файл в системе

        return txt # возвращаем значение текста

    def select_item(self, item): # выбираем и прописываем в строку ТТ

        line = self.listbox_text_processing(item.text()) # обработка текста списка (не вводим лишнее в ТТ)

        if line: # если есть строка используем

            if isinstance(line, list): # если значение параметра список, обработать каждое значение
                line_count = len(line) # количество строк
                line = "\n".join(line) # строки разделённые знаком переноса

            else:
                line_count = 0 # нет строк

            iKompasDocument = iApplication.ActiveDocument # делаем активный открытый документ

            if iKompasDocument is None or iKompasDocument.DocumentType != 1: # если нет открытого документа выдать сообщение
                self.kompas_message("Откройте чертёж!") # сообщение в окне компаса если он открыт

            else:
                iDrawingDocument = KompasAPI7.IDrawingDocument(iKompasDocument) # интерфейс чертежа
                iTechnicalDemand = iDrawingDocument.TechnicalDemand # интерфейс технических требований
                iText = iTechnicalDemand.Text # интерфейс текста ТТ
                #iText.Clear() # очистить ТТ

                iTextLine = iText.Add() # указатель на интерфейс ITextLine текста (добавляет новую строку в ТТ)
                iTextLine.Align = -1 # выравнивание (0 - слева,  1 - по центру, 2 - справа, 3 - во всю ширину, 4 - по десятич­ной точке, -1 - по умолчанию (из стиля)
##                iTextLine.IndentedLine = 5 # cмещение красной строки
##                iTextLine.LeftEdge = 0 # отступ текста слева
                iTextLine.Level = 0 # уровень вложенности нумерации
                iTextLine.Numbering = 1 # тип нумерации абзаца (-1 - Тип не определенный; 0 - Cрока без нумерации; 1 - Cтрока с нумерацией уровня level; 2 - На строке начинается новая нумерация пунктов; 3 - Cтрока не должна нумероваться никогда)
##                iTextLine.RightEdge = 0 # отступ текста справа
##                iTextLine.Step = 10 # шаг строк
##                iTextLine.StepAfterParagraph = 0 # дополнительный шаг после абзаца
##                iTextLine.StepBeforeParagraph = 0 # дополнительный шаг перед абзаца
##                iTextLine.Style = 2 # системный стиль текста (2 - Текст для технических требований)
##                iTextLine.Delete() # удалить строку текста (строку с текстом)

                iTextItem = iTextLine.Add() # указатель на интерфейс ITextItem (добавить компонент строки в конец строки)
                iTextItem.ItemType = 0 # тип компонента текста (0 - cтрока, остальное см. ksTextItemEnum)
##                iTextItem.NewLine = True # признак начала строки
                iTextItem.Str = line # текстовое значение компоненты текста (сам текст из списка)
##                iTextItem.SymbolFontName = '' # имя шрифта для символа
##                iTextItem.Delete() # удалить компоненту строки (только текст)
##
##                iTextFont = KompasAPI7.ITextFont(iTextItem) # интерфейс параметров шрифта
##                iTextFont.Bold = False # жирный шрифт
##                iTextFont.Color = 0 # цвет
##                iTextFont.FontName = 'GOST type A' # имя шрифта
##                iTextFont.Height = 5 # высота текта
##                iTextFont.Italic = True # курсив
##                iTextFont.Underline = False # подчёркивание
##                iTextFont.WidthFactor = 1 # коэффициент сужения
                iTextItem.Update() # обновить данные компонента

                if line_count > 1:
                    iCount = iText.Count # количество строк

                    for i in range(1, line_count): # для каждой строки
                        line_index = iCount - i

                        if line_index >= 0:  # Проверяем, что индекс существует
                            iTextLine = iText.TextLine(line_index) # указываем на последнюю строку (отсчёт с 0)
                            iTextLine.Numbering = 3 # тип нумерации абзаца (-1 - Тип не определенный; 0 - Cрока без нумерации; 1 - Cтрока с нумерацией уровня level; 2 - На строке начинается новая нумерация пунктов; 3 - Cтрока не должна нумероваться никогда)

                iTechnicalDemand.Update() # обновить данные ТТ

    def kompas_message(self, text): # сообщение в окне КОМПАСа если он открыт

        if iApplication.Visible == True: # если компас видимый
            iApplication.MessageBoxEx(text, 'Message:', 64) # сообщение в КОМПАСе

    def listbox_text_processing(self, text): # обработка текста списка (не вводим лишнее в ТТ)

        if re.search(r"--.+--", text): # Проверяем, является ли строка разделителем (содержит -- с любым содержимым между ними)
            return False

        # Удаляем комментарии с помощью регулярного выражения
        line = re.sub(r"#.*", "", text).strip()

        if not line: # Если после обработки строка пустая, возвращаем False
            return False

        if "\\n" in line: # Обрабатываем переносы строк
            return line.split("\\n")

        return line # возвращаем строку

    def message(self, text="Ошибка!", counter=4): # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия) После запуска основного окна!!!

        message(text, counter, self) # используем универсальную функцию с указанием текущего окна как родителя

    def clear_line(self): # очистка послетней строки ТТ

        iKompasDocument = iApplication.ActiveDocument # делаем активный открытый документ

        if iKompasDocument == None: # пока нет активного 2D документа выдавать сообщение
            self.kompas_message("Откройте чертёж!") # сообщение в окне компаса если он открыт

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

    def clear(self): # очистка ТТ

        iKompasDocument = iApplication.ActiveDocument # делаем активный открытый документ

        if iKompasDocument == None: # пока нет активного 2D документа выдавать сообщение
            self.kompas_message("Откройте чертёж!") # сообщение в окне компаса если он открыт

        else:
            iDrawingDocument = KompasAPI7.IDrawingDocument(iKompasDocument) # интерфейс чертежа
            iTechnicalDemand = iDrawingDocument.TechnicalDemand # интерфейс технических требований
            iText = iTechnicalDemand.Text # интерфейс текста ТТ
            iText.Clear() # очистить ТТ
            iTechnicalDemand.Update() # обновить данные ТТ

    def edit_file(self): # открыть .txt файл для редактирования

        server_path = self.settings_val("server_path") # путь к папке на сервере

        radio_btn = self.radio_group.checkedButton() # получаем текущую выбранную радиокнопку

        file_name = radio_btn.property("file_name") # получаем имя файла из свойства кнопки

        if server_path: # если путь к файлам на сервере указан

            file_server_path = os.path.join(server_path, file_name) # полный путь к файлу на сервере

            if os.path.exists(file_server_path): # если есть файл
                os.startfile(file_server_path) # открываем файл в системе

            else:
                local_file = os.path.join(program_directory, file_name) # полный путь к локальному файлу

                os.startfile(local_file) # открываем файл в системе
                self.message("Открыт локальный файл!") # сообщение, поверх всех окон с автоматическим закрытием

        else: # путь к файлам на сервере не указан
            os.startfile(os.path.join(program_directory, file_name)) # открываем файл в системе

    def open_settings(self): # открыть файл с настройками программы

        settings_file = os.path.join(program_directory, title + ".ini")
        print(settings_file)

        os.startfile(settings_file) # открываем файл в системе

        self.message("Введите необходимые значения! \nИ запустите приложение повторно!") # сообщение с названием файла

    def on_top(self): # поверх всех окон

        current_flags = self.windowFlags() # Получаем текущие флаги окна

        if current_flags & Qt.WindowType.WindowStaysOnTopHint: # проверяем, установлен ли уже флаг "поверх всех окон"
            new_flags = current_flags & ~Qt.WindowType.WindowStaysOnTopHint # если установлен, убираем его

            # self.on_top_btn.setIcon(QIcon(":/TTEditor/stud_button_normal.png")) # меняем иконку на обычную

            self.save_settings_val("on_top", False) # сохранить значения в словарь

        else:
            new_flags = current_flags | Qt.WindowType.WindowStaysOnTopHint # если не установлен, добавляем его

            # self.on_top_btn.setIcon(QIcon(":/TTEditor/stud_button_active.png")) # меняем иконку на активную

            self.save_settings_val("on_top", True) # сохранить значения в словарь

        self.setWindowFlags(new_flags) # применяем флаги
        self.show()  # Необходимо вызвать show() после изменения флагов

def resource_path(relative_path): # для сохранения картинки внутри exe файла

    try: # попытаться определить путь к папке
        base_path = sys._MEIPASS # путь к временной папки PyInstaller

    except Exception: # если ошибка
        base_path = os.path.abspath(".") # если не в exe, используем текущую директорию

    return os.path.join(base_path, relative_path) # объеденяем и возващаем полный путь

def message(text="Ошибка!", counter = 4, parent = None): # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия) с возможностью открытия до запуска основного окна!!!

    app = QApplication.instance() # Проверяем, существует ли экземпляр QApplication

    if app is None:  # Если приложение еще не создано
        app = QApplication(sys.argv) # новый экземпляр QApplication (sys.argv)
        need_exec = True # триггер

    else:
        need_exec = False # триггер

    msg = AutoCloseMessageBox(text, counter, parent) # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия)

    loop = QEventLoop() # создаем локальный цикл событий
    msg.closed.connect(loop.quit)

    msg.show() # показываем окно

    if need_exec: # если приложение было создано здесь, запускаем цикл событий
        loop.exec() # локальный цикл
        app.quit()  # завершаем приложение после закрытия сообщения

    else: # Если приложение уже существует, используем локальный цикл
        loop.exec() # локальный цикл

class AutoCloseMessageBox(QMessageBox): # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия)

    closed = pyqtSignal() # объявляем сигнал closed

    def __init__(self, text = "Ошибка!", counter = 4, parent=None):

        super().__init__(parent)

        self.duration = counter  # Время до автоматического закрытия (сек)
        self.text = text  # Исходный текст сообщения
        self.title = title # заголовок окна

        # Настройка окна
        self.setWindowTitle(self.title)
        self.setWindowIcon(QIcon(resource_path("cat.ico"))) # значок программы
        self.setIcon(QMessageBox.Icon.Information)
        self.setWindowFlags(self.windowFlags() | Qt.WindowType.WindowStaysOnTopHint)

        # Таймер для автоматического закрытия
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_countdown)

        # Обработчик ручного закрытия
        self.finished.connect(self.on_message_closed)

        self.update_text()  # Первоначальное обновление текста
        self.timer.start(1000)  # Запуск таймера (обновление каждую секунду)

    def update_countdown(self): # обновление текста по времени

        self.duration -= 1
        if self.duration > 0:
            self.update_text() # обновление текста
        else:
            self.timer.stop() # остановка таймера
            self.close() # закрываем окно
            self.closed.emit() # уведомляет внешний код о закрытии

    def update_text(self): # обновление текста
        self.setText(self.text)

    def on_message_closed(self, result): # обработчик закрытия сообщения (вызывается в любом случае)
        self.timer.stop()  # Останавливаем таймер при любом закрытии
        self.closed.emit() # Автоматическое закрытие

def save_ini_settings(): # cохраняет настройки в INI-файл

    config = configparser.ConfigParser()
    ini_path = os.path.join(program_directory, title + ".ini")

    # Читаем текущие настройки
    config.read(ini_path, encoding='utf-8')

    # Обновляем значения
    for key, value in settings.items():
        config.set('Settings', key, str(value))

    # Сохраняем обратно
    with open(ini_path, 'w', encoding='utf-8') as configfile:
        config.write(configfile)

#-------------------------------------------------------------------------------

if __name__ == "__main__":  # если мы запускаем файл напрямую, а не импортируем

    options = {"Дет." : "ТТ.txt", # название радиокнопок и список txt файлов к ним
               "СБ" : "ТТ СБ.txt",
               "УЧ" : "ТТ УЧ.txt",
               "КЭ" : "ТТ КЭ.txt",
               "Опт." : "ТТ Опт.txt"}

    try: # попытаться запустить интерфейс

        main() # запускаем функцию main()

##        save_ini_settings(): # cохраняет настройки в INI-файл

    except Exception as e:
        _, _, exc_tb = sys.exc_info()
        line_number = exc_tb.tb_lineno
        msg = "Произошла ошибка:\n" + traceback.format_exc()
        print(msg)
        message(msg, 10) # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия)