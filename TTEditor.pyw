#-------------------------------------------------------------------------------
# Author:      dimak222
#
# Created:     18.01.2022
# Copyright:   (c) dimak222 2022
# Licence:     No
#-------------------------------------------------------------------------------

title = "TTEditor" # наименование приложения
ver = "v2.1.0.0" # версия файла

#------------------------------Настройки!---------------------------------------

##parameters = {
##'check_update' : 'True # проверять обновление программы (True - да; False или "" - нет)',
##'beta' : 'False # скачивать бета версии программы (True - да; False или "" - нет)\n',
##
##'server_path' : '"" # путь к папке с txt файлами на сервере (для копирования файла рядом спрограммой), False или "" - использует файлы только рядом с программой\n',
##
##'more_options' : 'False # добавление доп. файлов через ";", пример: "ТТ МЧ : ТТ МЧ.txt; ТТ Л3 : ТТ Л3.txt" (Имя кнопки : Название txt файла). False или "" - не добавлять\n',
##
##'on_top' : 'True # запуск программы поверх всех окон, False или "" - выключить',
##
##'import_tt' : 'True # импортировать ТТ с сервера, False или "" - не импортировать\n',
##'import_tt_messages' : 'True # выдавать сообщения если на сервере нет txt файлов, False или "" - не выдавать\n',
##
##'last_choice_file' : 'Дет. # файл открываемый при запуске програмы',
##'window_geometry' : ' # размер и положение окна',
##}

#------------------------------Импорт модулей-----------------------------------

import psutil # модуль вывода запущеных процессов
import os # работа с файовой системой

import configparser # модуль для работы с INI-файлами
import ast # модуль преобразования переменных

import ctypes # для изменения атрибутов файла

from sys import exit # для выхода из приложения без ошибки

##import pythoncom # модуль для запуска без IDE
from win32com.client import Dispatch, gencache # библиотека API Windows  # noqa: F401
from pythoncom import connect, CoInitialize, CoUninitialize # подключаемся к запущенному экземпляру КОМПАС и работа в потоках

import filecmp # модуль сравнения файлов
from send2trash import send2trash # модуль для удаления файлов в корзину
import shutil # библиотека для копирования/перемещений/переименований

from PyQt6.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox, QDialog, QRadioButton, QCheckBox, QButtonGroup, QMenu
from PyQt6.QtGui import QIcon, QFont
from PyQt6.QtCore import Qt, QTimer, pyqtSignal, QByteArray, QThread

from main_window import Ui_MainWindow # импорт GUI окна
from settings_window import Ui_SettingsWindow # импорт GUI окна
import resources_rc  # импорт ресурсов (иконок)  # noqa: F401

import re # модуль регулярных выражений

import sys # для определения строки ошибки

import traceback # для вывода ошибок

#-------------------------------------------------------------------------------

def main(): # запуск интерфейса

    app = QApplication(sys.argv) # новый экземпляр QApplication (sys.argv)

    CheckUpdate() # проверить обновление приложение

    window = MyApp()  # создаём объект класса ExampleApp
    window.setWindowIcon(QIcon(resource_path("icon.ico"))) # значок программы
    window.setWindowTitle(f"{title} {ver}") # заголовок окна

    window.show()  # показываем окно

    app.exec() # запускаем приложение

    settings_loader.save_ini_settings() # cохраняем настройки в INI-файл

def resource_path(relative_path): # для сохранения картинки внутри exe файла

    try: # попытаться определить путь к папке
        base_path = sys._MEIPASS # путь к временной папки PyInstaller

    except Exception: # если ошибка
        base_path = os.path.abspath(".") # если не в exe, используем текущую директорию

    return os.path.join(base_path, relative_path) # объеденяем и возващаем полный путь

def message(text="Ошибка!", duration=4, message_type="info", exit_app=False): # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия, закрытие программы)

    print(f"[{message_type.upper()}][{exit_app}] {text}")

    message_manager.add_message(text, duration, message_type, exit_app) # используем менеджер сообщений

class MessageManager: # менеджер сообщений с очередью

    def __init__(self):
        self._initialized = False # триггер запуска основного окна

    def initialize(self, main_window=None):
        """Инициализация после создания главного окна"""

        if self._initialized: # триггер запуска основного окна
            return

        self.main_window = main_window
        self._message_queue = []
        self._is_showing = False
        self._initialized = True

    def add_message(self, text, duration=4, message_type="info", exit_app=False): # Добавление сообщения в очередь

        self.exit_app = exit_app # тригер закрытия программы

        if not self._initialized: # триггер запуска основного окна
            # Если менеджер не инициализирован, показываем сообщение напрямую
            self._show_message_direct(text, duration, message_type, exit_app)
            return

        self._message_queue.append({
            "text": text,
            "duration": duration,
            "type": message_type,
            "exit_app": exit_app,
        })

        self._process_queue() # обработка очереди сообщений

    def _process_queue(self): # обработка очереди сообщений

        if self._is_showing or not self._message_queue:
            return

        self._is_showing = True
        msg_data = self._message_queue.pop(0)

        msg = AutoCloseMessageBox(
            msg_data["text"],
            msg_data["duration"],
            msg_data["type"],
            self.main_window,
        )

##        msg.exit_app = msg_data["exit_app"]

        msg.closed.connect(self._on_message_closed)
        msg.show()

    def _on_message_closed(self): # когда окно закрылось

        self._is_showing = False
        self._process_queue # обработка очереди сообщений

        if self.exit_app: # если сообщение с закрытием программы
            exit() # завершаем программу

    def _show_message_direct(self, text, duration, message_type, exit_app):
        """Прямой показ сообщения (если менеджер не инициализирован)"""
        app = QApplication.instance() # если уже запущен

        if app is None:
            app = QApplication(sys.argv) # новый экземпляр QApplication (sys.argv)

        msg = AutoCloseMessageBox(text, duration, message_type) # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия)
        msg.exec() # показать

        if exit_app: # если сообщение с закрытием программы
            exit() # завершаем программу

class AutoCloseMessageBox(QMessageBox): # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия)

    closed = pyqtSignal() # объявляем сигнал closed

    def __init__(self, text = "Ошибка!", duration = 4, message_type="info", parent=None):

        super().__init__(parent)

        self.duration = duration # время до автоматического закрытия (сек)
        self.text = text # исходный текст сообщения

        self.title = f"{title} {ver}" # заголовок окна

        self.setWindowTitle(self.title) # Настройка окна
        self.setWindowIcon(QIcon(resource_path("icon.ico"))) # значок программы

        # Настраиваем иконку в зависимости от типа сообщения
        icon_map = {
            "error": QMessageBox.Icon.Critical,
            "warning": QMessageBox.Icon.Warning,
            "info": QMessageBox.Icon.Information,
            "success": QMessageBox.Icon.Information
        }
        self.setIcon(icon_map.get(message_type)) # иконка сообщения

        self.setWindowFlags(self.windowFlags() | Qt.WindowType.WindowStaysOnTopHint) # поверх всех окон

        self.update_text() # Первоначальное обновление текста

        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_countdown) # таймер для автоматического закрытия

        self.finished.connect(self.on_message_closed) # обработчик ручного закрытия

        self.timer.start(1000) # Запуск таймера (обновление каждую секунду)

    def update_countdown(self): # обновление текста по времени

        self.duration -= 1
        if self.duration > 0:
            self.update_text() # обновление текста
        else:
            self.close() # закрываем окно

    def update_text(self): # обновление текста
        self.setText(self.text)

    def on_message_closed(self, result): # обработчик закрытия сообщения (вызывается в любом случае)
        self.timer.stop() # останавливаем таймер при любом закрытии
        self.closed.emit() # уведомляет внешний код о закрытии

class SettingsLoader: # загружаем настройки из INI-файла

    def __init__(self): # запуск при инициализации

        self.doubleExe() # проверка на уже запущеное приложение

        self.load_settings() # загружает настройки из INI-файла

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
                            message("Приложение уже запущено!", 5, "warning", True) # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия, закрытие программы)

        if list == []: # если нет найденых названий программы
            program_directory = os.path.dirname(os.path.abspath(__file__)) # директория рядом с программой

        else: # если путь найден
            program_directory = os.path.dirname(psutil.Process().exe()) # директория программы

        program_directory = program_directory.replace("\\", "//", 1) # замена на слеш первого символа (при "\\192.168....")

    def load_settings(self): # загружает настройки из INI-файла

        global settings # делаем глобальным

        def load_ini_config(ini_path): # загружает параметры из INI-файла, автоматически удаляя комментарии (путь к INI-файлу)

            config = configparser.ConfigParser(inline_comment_prefixes=("#",)) # считываем настройки удаляя "#"

            if os.path.exists(ini_path): # если файл существует

                try:
                    config.read(ini_path, encoding='utf-8') # читаем файл
                    if 'Settings' in config:
                        return config['Settings']

                except Exception:
                    pass # пропускаем

            try: # создаём и повторно читаем файл

                create_default_ini(ini_path) # cоздаёт INI-файл с параметрами по умолчанию (путь к INI-файлу)

                config.read(ini_path, encoding='utf-8') # читаем файл

                return config['Settings']

            except Exception:
                msg = f"Не удалось создать или прочитать INI-файл настроек \"{ini_path}\""
                message(msg, 5, "error", True) # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия, закрытие программы)

        def create_default_ini(ini_path): # cоздаёт INI-файл с параметрами по умолчанию (путь к INI-файлу)

            config = configparser.ConfigParser()
            config['Settings'] = {} # создаем файл ini

            with open(ini_path, 'w', encoding='utf-8') as configfile:
                config.write(configfile)

        def convert_value(value): # функция для автоматического преобразования

            try:
                value = value.strip('"').strip() # удаляем '"' и пробелов по бокам если они есть
                return ast.literal_eval(value) # преобразуем значения

            except (ValueError, SyntaxError):
                return value  # Если преобразование невозможно, оставляем как есть

        ini_path = os.path.join(program_directory, title + ".ini") # путь к INI-файлу
        settings = load_ini_config(ini_path) # загружает параметры из INI-файла, автоматически удаляя комментарии (путь к INI-файлу)

        settings = {key: convert_value(value) for key, value in settings.items()} # создание и преобразование словаря

        print("Параметры загружены из INI файла!")

##        for key, value in settings.items(): # вывод загруженных параметров
##            print(f"{key} = {value}")

        print("#--------------------------------")

    def settings_val(self, key): # получить значение из словаря

        key = key.lower() # преобразуем ключ в нижний регистр для поиска

        value = settings.get(key) # берем переменную из словаря

        return value # возвращаем значение

    def save_settings_val(self, key, val): # сохранить значения в словарь

        key = key.lower() # преобразуем ключ в нижний регистр для поиска

        settings[key] = val # присваеваем новое значение

    def save_ini_settings(self): # cохраняем настройки в INI-файл

        try: # попытаться сохранить

            config = configparser.ConfigParser()
            ini_path = os.path.join(program_directory, title + ".ini")

            if os.path.exists(ini_path): # Если файл существует

                current_attrs = self.set_file_hidden(ini_path, 32) # если файл является скрытым, временно показываем его (32)

                config.read(ini_path, encoding='utf-8') # читаем текущие настройки

                save_flag = False # флаг сохранения

                if 'Settings' not in config: # если секция Settings не существует
                    config['Settings'] = {} # создадим её

                for key, value in settings.items(): # обновляем значения

                    if config['Settings'].get(key) != str(value): # ключа нет или значение отличается

                        print(f"{key} = {str(value)}")
                        config.set('Settings', key, str(value)) # обновляем

                        save_flag = True # флаг сохранения

                if save_flag: # если было изменение

                    with open(ini_path, 'w+', encoding='utf-8') as configfile: # сохраняем обратно
                        config.write(configfile)

                    print("#--------------------------------")
                    print("Параметры сохранены в INI файл!")

                self.set_file_hidden(ini_path, current_attrs) # возвращаем атрибут файла

        except Exception as e:
            msg = f"Ошибка при сохранении настроек: {e}"
            message(msg, 8, "error", True) # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия, закрытие программы)

    def set_file_hidden(self, file_path, attrs): # изменение атрибутов файла для сохранения

        if os.name == 'nt' and attrs:  # Проверяем, что это Windows и есть атрибут

            try:
                current_attrs = ctypes.windll.kernel32.GetFileAttributesW(file_path) # Получаем текущие атрибуты файла

                if current_attrs != attrs: # если они отличаются
                    ctypes.windll.kernel32.SetFileAttributesW(file_path, attrs) # применяем новый атрибкт

                return current_attrs # возвращаем исходное значение

            except Exception as e:
                print(f"Ошибка изменения атрибута файла: {e}")
                return False # вернуть значение

class CheckUpdate: # проверить обновление приложение

    def __init__(self): # запуск при инициализации
        self.check_update() # проверить обновление приложение

    def check_update(self): # проверить обновление приложение

        global url # значение делаем глобальным

        if settings_loader.settings_val("check_update"): # если проверка обновлений включена

            try: # попытаться импортировать модуль обновления

                from Updater import Updater # импортируем модуль обновления

                if "url" not in globals(): # если нет ссылки на программу
                    url = "" # без ссылки

                Updater.Update(title, ver, settings_loader.settings_val("beta"), url, program_directory, resource_path("icon.ico")) # проверяем обновление (имя программы, версия программы, скачивать бета версию, ссылка на программу, директория программы, путь к иконке)

            except SystemExit: # если закончили обновление (exit в Update)
                exit() # выходим из программы

            except Exception: # не удалось
                message("Ошибка обновления!", 4, "warning", False) # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия, закрытие программы)

        else:
            print("Проверка обновлений выключена")

class MyApp(QMainWindow, Ui_MainWindow): # основное окно

    def __init__(self): # нужно для доступа к переменным, методам и т.д. в файле main_window.pyw

        super().__init__()

        self.setupUi(self) # инициализация дизайна

        self.restore_window_geometry() # восстанавливаем положение, геометрию окна и поверх всех окон

        self.window_btn() # кнопки окна

        self.kompas_worker = KompasWorkerThread(self) # cоздание и настройка рабочего потока

        self.kompas_worker_signal() # сигналы потока

        self.kompas_worker.start() # запускаем поток (функцию run)

        message_manager.initialize(self) # инициализируем менеджер сообщений

    def restore_window_geometry(self): # восстанавливаем положение, геометрию окна и поверх всех окон

        try:
            on_top = settings_loader.settings_val('on_top') # значене положения и геометрии окон

            current_flags = self.windowFlags() # Получаем текущие флаги окна

            if on_top: # если есть значение
                new_flags = current_flags | Qt.WindowType.WindowStaysOnTopHint # если не установлен, добавляем его
                settings_loader.save_settings_val("on_top", True) # сохранить значения в словарь

            else:
                new_flags = current_flags & ~Qt.WindowType.WindowStaysOnTopHint # если не установлен, добавляем его
                settings_loader.save_settings_val("on_top", False) # сохранить значения в словарь

            self.setWindowFlags(new_flags) # применяем флаги

            window_geometry = settings_loader.settings_val('window_geometry') # значене положения и геометрии окна

            if window_geometry: # если есть значение
                geometry = QByteArray.fromBase64(window_geometry.encode())
                self.restoreGeometry(geometry)

        except Exception as e:
            msg = f"Ошибка восстановления состояния окна : {e}"
            message(msg, 10, "warning") # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия)

    def window_btn(self): # кнопки окна

        self.temp_states = {} # временное хранилище состояний чекбоксов

        self.listWidget.itemClicked.connect(self.send_command("select_item")) # для работы нажатия списка

        self.setup_radio_buttons() # настройка группы радиокнопок внутри doc_type_groupBoxx

        self.clear_line_btn.clicked.connect(self.send_command("clear_line")) # удалить последий пункт
        self.clear_btn.clicked.connect(self.send_command("clear")) # очистить ТТ

        self.edit_file_btn.clicked.connect(self.edit_file) # редактировать файл ТТ

        self.settings_btn.clicked.connect(self.open_settings) # открытие настроек
        self.settings_dialog = SettingsDialog(self) # окно с настройками (self как parent)

        self.on_top_btn.clicked.connect(self.on_top) # кнопка поверх всех окон

    def setup_radio_buttons(self): # настройка группы радиокнопок внутри doc_type_groupBox

        self.radio_group = QButtonGroup(self) # создаём группу кнопок (если ещё не создана)
        self.radio_group.buttonClicked.connect(self.on_radio_button_clicked) # подключаем сигнал

        self.bind_existing_radio_buttons() # обрабатываем уже существующие в doc_type_groupBox радиокнопки

        self.add_custom_radio_buttons() # добавляем пользовательские кнопки из more_options

        self.comparing_and_copying_files(settings_loader.settings_val("import_tt_messages")) # сравнение txt файлов с сервера и рядом с программой и их копирование (путь к файлам на сервере, выдавать сообщения если на сервере нет .txt файлов)

        self.restore_last_choice() # восстанавливаем последний выбор

        self.doc_type_groupBox.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu) # контекстное меню
        self.doc_type_groupBox.customContextMenuRequested.connect(self.on_groupbox_right_click) # нажатие правой кнопки

    def on_radio_button_clicked(self, button): # нажатие любой кнопки из группы

        btn_name = button.text() # имя кнопки

        if btn_name: # если есть имя
            settings_loader.save_settings_val("last_choice_file", btn_name) # сохранить значения в словарь

            self.update_list(button.property("file_name")) # обновляем список (имя файла, пересканировать файлы с сервера)

    def update_list(self, file_name): # обновляем список (имя файла, пересканировать файлы с сервера)

        txt_list = self.read_TT_file(file_name) # считывание с txt файла

        self.listWidget.clear() # удаляем весь список
        self.listWidget.addItems(txt_list) # добавляем все строки из списка

    def read_TT_file(self, txt_filename): # считываем ТТ с txt файла

        file_path = os.path.join(program_directory, txt_filename) # название текстового файла

        if not os.path.exists(file_path): # если нет txt файла
            txt_list = self.to_create_TT_file(file_path) # создать txt файл с записью в него значений
            message(f"Создан файл {txt_filename}, заполните его!", 8) # сообщение, поверх всех окон с автоматическим закрытием

        with open(file_path, 'r', encoding='utf-8') as txt_file: # используем менеджер контекста для автоматического закрытия файла
            txt_list = txt_file.read().splitlines() # читаем и сразу разбиваем на строки

        return txt_list # возвращаем считаный список

    def to_create_TT_file(self, file_path): # создать txt файл с записью в него значений

        txt = """Пример строки для вставки в ТТ. # заполните файл для отображения строк ТТ в программе. Значения после "#" работают как комментарии (не передаются в ТТ чертежа).
# для перевода текста ТТ на новую строку без нумерации, испольуйте "\\n". Пример: "Редактор\\nтехнических\\nтребований"
------------------------------------------------------------------------------------------- # для визуального отделения ТТ в программе можно использовать "----"
    """ # текст записываемый в .txt файл

        with open(file_path, "w", encoding="utf-8") as txt_file: # используем менеджер контекста для автоматического закрытия файла
            txt_file.write(txt) # записать файл

        os.startfile(file_path) # открываем файл в системе

        return txt # возвращаем значение текста

    def bind_existing_radio_buttons(self): # обрабатываем уже существующие в doc_type_groupBox радиокнопки

        options = {"Дет." : "ТТ.txt", # название радиокнопок и список txt файлов к ним
                   "СБ" : "ТТ СБ.txt",
                   "УЧ" : "ТТ УЧ.txt",
                   "КЭ" : "ТТ КЭ.txt",
                   "ПП" : "ТТ ПП.txt",
                   "Опт." : "ТТ Опт.txt",
                   }

        for radio_btn in self.doc_type_groupBox.findChildren(QRadioButton): # проходим по всем дочерним радиокнопкам внутри doc_type_groupBox

            name = radio_btn.text() # имя кнопкия

            if name in options: # привязка каждой кнопки к файлу
                file_name = options[name] # имя файа
                radio_btn.setProperty("file_name", file_name) # привязка сво-ва к кнопке

                self.radio_group.addButton(radio_btn) # добавляем кнопку в группу

            else:
                radio_btn.setVisible(False) # скрыть кнопку если её нет в списке

    def add_custom_radio_buttons(self): # добавляем пользовательские кнопки из more_options

        more_options = settings_loader.settings_val("more_options") # получить данные

        if not more_options: # если данных нет
            return # прервать

        more_options = more_options.split(";") # разделяем на список

        for btn in more_options: # перебор всех новых кнопок
            if ':' not in btn: # если нет разделения
                continue # проускаем

            btn_text, file_name = btn.split(":") # разделяем
            btn_text = btn_text.strip() # удаляем пробелы
            file_name = file_name.strip() # удаляем пробелы

            radio = QRadioButton(btn_text, self.doc_type_groupBox) # создаём кнопку
            radio.setProperty("file_name", file_name) # присваиваем сво-во

            self.horizontalLayout_3.addWidget(radio) # добавляем в layout
            self.radio_group.addButton(radio) # добавляем в группу

    def comparing_and_copying_files(self, import_tt_messages): # сравнение txt файлов с сервера и рядом с программой и их копирование (путь к файлам на сервере, выдавать сообщения если на сервере нет .txt файлов)

        if not settings_loader.settings_val("import_tt"): # если импорт с сервера выключен
            return # прервать

        server_path = settings_loader.settings_val("server_path") # путь к серверу с ТТ

        if not (server_path and os.path.exists(server_path)): # если путь к файлам на сервере не указан и он есть

            if import_tt_messages: # выдавать сообщения если на сервере нет .txt файлов
                message("Путь к папке на сервере не указан или не найден!") # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия)

            return # прервать

        summary = []  # список строк для итогового сообщения

        for radio_btn in self.radio_group.buttons(): # перебираем все кнопки

            file_name = radio_btn.property("file_name") # имя файла

            file_server  = os.path.join(server_path, file_name) # полный путь к файлу на сервере
            file_local = os.path.join(program_directory, file_name) # полный путь к файлу в папке с программой

            if not os.path.exists(file_server): # если нет файл на сервере

                if import_tt_messages: # выдавать сообщения если на сервере нет .txt файлов
##                    message(f"Файл \"{file_name}\" на сервере не найден!") # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия)
                    summary.append(f"❌ Файл \"{file_name}\" на сервере не найден") # сообщение

                continue # продолжить

            try: # попытаться заменить файлы

                if os.path.exists(file_local): # если есть файл в папке с программой

                    if not filecmp.cmp(file_server, file_local, shallow = True): # если файлы разные, обработать (сравнивает только метаданные файлов)

                        try: # попытаться удалить в корзину
                            send2trash(file_local) # старый файл удаляем в корзину

                        except Exception: # в случае ошибки
                            print(f"Файл \"{file_name}\" не может быть удалён в корзину!")

                        shutil.copy2(file_server, file_local) # копируем файл с сервера с сохранением методанных
##                        message(f"Файл \"{file_name}\" обновлён!") # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия)
                        summary.append(f"✅ Файл \"{file_name}\" обновлён")

                else: # нет файла в папке с программой
                    shutil.copy2(file_server, file_local) # копируем файл с сервера с сохранением методанных
##                    message(f"Файл \"{file_name}\" скопирован с сервера!", 2) # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия)
                    summary.append(f"📁 Файл \"{file_name}\" скопирован с сервера")

            except Exception: # в случае ошибки
##                message(f"Нет прав на замену файла \"{file_name}\".\nЗапустите программу от администратора\nили переместите её в другую папку!", message_type = "warning") # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия)
                summary.append(f"⛔ Нет прав на замену файла \"{file_name}\".\nЗапустите программу от администратора\nили переместите её в другую папку!")

        if summary and import_tt_messages: # если есть сообщения и они включены
            final_text = "Результаты синхронизации с сервера:\n" + "\n".join(summary) # итоговое сообщение
            message(final_text, duration=8, message_type="info") # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия)

    def restore_last_choice(self): # восстанавливаем последний выбор

        last_choice_file = settings_loader.settings_val("last_choice_file") # получаем значение из настроек

        for radio_btn in self.radio_group.buttons(): # ищем кнопку с таким file_name
            if radio_btn.text() == last_choice_file: # если имя кнопки совпала с мпоследней кнопкой
                radio_btn.setChecked(True) # сделать её активной
                self.update_list(radio_btn.property("file_name")) # обновить список
                return # прервать

        else: # если перебрали все кнопки
            for radio_btn in self.radio_group.buttons(): # перебрать и найти активную кнопку
                 if radio_btn.isChecked(): # если радиокнопка активна
                    self.update_list(radio_btn.property("file_name")) # обновить список
                    return # прервать

    def on_groupbox_right_click(self, pos): # нажатие правой кнопки

        menu = QMenu() # контекстное меню

        font = QFont() # шрифт
        font.setPointSize(10) # размер шрифта
        menu.setFont(font) # применить шрифт

        sync_action = menu.addAction("🔄 Обновить файл ТТ с сервера") # текст в меню

        action = menu.exec(self.doc_type_groupBox.mapToGlobal(pos)) # показать (позиция)

        if action == sync_action: # если действие выполнено

            button = self.radio_group.checkedButton() # берём активную радиокнопку

            file_name = button.property("file_name") # имя файла

            self.comparing_and_copying_files(False) # сравнение txt файлов с сервера и рядом с программой и их копирование (без сообщений)

            self.update_list(file_name) # обновляем список (имя файла, пересканировать файлы с сервера)

            message(f"Файл \"{file_name}\" Обновлён!", message_type = "info") # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия)

    def edit_file(self): # открыть .txt файл для редактирования

        server_path = settings_loader.settings_val("server_path") # путь к серверу с ТТ

        radio_btn = self.radio_group.checkedButton() # получаем текущую выбранную радиокнопку
        file_name = radio_btn.property("file_name") # получаем имя файла из свойства кнопки
        local_file = os.path.join(program_directory, file_name) # полный путь к локальному файлу

        if not settings_loader.settings_val("import_tt"): # если импорт с сервера выключен
            os.startfile(local_file) # открываем локальный файл в системе
            return # прервать

        file_server = os.path.join(server_path, file_name) # полный путь к файлу на сервере

        if os.path.exists(file_server): # если есть файл
            os.startfile(file_server) # открываем файл в системе

        else:
            os.startfile(local_file) # открываем файл в системе
            message("Открыт локальный файл!") # сообщение, поверх всех окон с автоматическим закрытием

    def open_settings(self): # открываем окно с настройками

        self.settings_dialog.exec() # показываем как модальное окно

    def on_top(self): # поверх всех окон

        current_flags = self.windowFlags() # Получаем текущие флаги окна

        if current_flags & Qt.WindowType.WindowStaysOnTopHint: # проверяем, установлен ли уже флаг "поверх всех окон"
            new_flags = current_flags & ~Qt.WindowType.WindowStaysOnTopHint # если установлен, убираем его

            settings_loader.save_settings_val("on_top", False) # сохранить значения в словарь

        else:
            new_flags = current_flags | Qt.WindowType.WindowStaysOnTopHint # если не установлен, добавляем его

            settings_loader.save_settings_val("on_top", True) # сохранить значения в словарь

        self.setWindowFlags(new_flags) # применяем флаги
        self.show() # Необходимо вызвать show() после изменения флагов

    def kompas_worker_signal(self): # сигналы потока

        self.kompas_worker.text_line_msg_signal.connect(self.text_line_msg) # сигнал подключеня к КОМПАСу

        self.kompas_worker.initialized_signal.connect(self.on_kompas_initialized) # сигнал подключеня к КОМПАСу

    def text_line_msg(self, msg="", color=""): # изменение текстовой строки (текст, цвет)

        self.text_line.setStyleSheet(color) # сброс к стандартному цвету
        self.text_line.setText(msg) # прописать текст

    def on_kompas_initialized(self, success, msg): # сообщение о подключении к КОМПАСу

        if not success: # если подключились
            message(msg, 8, "error", True) # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия, закрытие программы)

    def send_command(self, command, data=None): # универсальная отправка команды в поток

        def caller(*args): # функцию-обработчик

            if args: # если сигнал передал аргументы, используем их

                first_arg = args[0] # первый аргумент

                if hasattr(first_arg, 'text'): # если есть атрибут текст
                    data = first_arg.text() # выдиляем текст, например, QListWidgetItem

                else:
                    data = first_arg # иначе передаём напрямую

            self.kompas_worker.add_command(command, data) # добавляем команду в поток

        return caller # возвращаем функцию

    def closeEvent(self, event): # действия после закрытия окна

        if hasattr(self, 'kompas_worker') and self.kompas_worker.isRunning(): # Останавливаем рабочий поток

            self.kompas_worker.quit() # запускаем завершение потока
            self.kompas_worker.wait(2000) # ожидаем завершения 2с
            if self.kompas_worker.isRunning(): # если всё ещё запущен
                self.kompas_worker.terminate() # принудительно завершаем
                print("Поток завершён принудительно!")
                self.kompas_worker.wait() # ожидаем завершения

        self.save_window_geometry() # сохранить геометрию и положение окна

        for checkbox_name, checked in self.temp_states.items(): # запись всех изменённых чекбоксов
            settings_loader.save_settings_val(checkbox_name, checked) # cохраняем состояние в настройки

        message_manager._initialized = False # используем менеджер сообщений
        super().closeEvent(event) # закрытие основного окна

    def save_window_geometry(self): # сохраняем положение и геометрию окна

        geometry = self.saveGeometry()
        window_geometry = geometry.toBase64().data().decode() # положение и размер окна
        settings_loader.save_settings_val('window_geometry', window_geometry) # сохранить значения в словарь

class SettingsDialog(QDialog, Ui_SettingsWindow): # окно с настройками

    def __init__(self, parent=None):

        super().__init__(parent)

        # Загружаем интерфейс из Qt Designer
        self.setupUi(self)

        self.setWindowTitle(f"{title} {ver}") # Устанавливаем заголовок окна

        # Устанавливаем иконку
        self.setWindowIcon(QIcon(resource_path("icon.ico")))

        # Делаем окно модальным (блокирует главное окно)
        self.setModal(True)

        # Временное хранилище состояний чекбоксов
        self.temp_states = {}

        # Настраиваем чекбоксы в окне настроек
        self.setup_checkbox()

##        self.setup_radio_group(self.coefficient_middle_layer_groupBox) # настройка группы радиокнопок внутри groupBox

        self.set_lines() # заполнить путь предыдущими значениями и проверка изменения

        self.path_btn.clicked.connect(self.open_file_directory) # открытть указание папки с ТТ

    def setup_checkbox(self): # считывание и установка состояния чекбокса

        checkboxes = self.findChildren(QCheckBox) # автоматический поиск всех чекбоксов (QCheckBox)

        for checkbox in checkboxes: # перебираем все чекбоксы

            checkbox.clicked.connect(self.save_checkBox) # подключаем сигнал

            checkbox_name = checkbox.objectName() # получаем имя свойства

            value = settings_loader.settings_val(checkbox_name) # получаем значение из настроек

            if value is not None: # если значение существует (не None)

                try: # попытаться применить настройки
                    checkbox.setChecked(value) # устанавливаем состояние чекбокса

                except Exception: # в случае ошибки
                    self.temp_states[checkbox_name] = checkbox.isChecked() # сохраняем дефолтное состояние чекбокса во временный словарь

            else:
                settings_loader.save_settings_val(checkbox_name, checkbox.isChecked()) # cохраняем состояние в настройки

    def save_checkBox(self): # сохранить статус чекбокса

        sender = self.sender() # получаем объект, который отправил сигнал

        checkbox_name = sender.objectName() # получаем имя чекбокса

        self.temp_states[checkbox_name] = sender.isChecked() # сохраняем состояние чекбокса во временный словарь

    def setup_radio_group(self, group_box): # настройка группы радиокнопок внутри groupBox

        radio_btns = group_box.findChildren(QRadioButton) # автоматический поиск радиокнопок в groupBox

        self.group_name = group_box.objectName() # имя группы

        value = settings_loader.settings_val(self.group_name) # получаем значение из настроек

        for radio_btn in radio_btns: # перебираем все радиокнопки

            if value is None: # если значение не существует (не None)

                if radio_btn.isChecked(): # если радиокнопка активна
                    settings_loader.save_settings_val(self.group_name, radio_btn.objectName()) # cохраняем состояние в настройки
                    break # прервать

            radio_btn.clicked.connect(self.save_radio_group) # подключаем сигнал

            if radio_btn.objectName() == value: # если имя радиокнопки записано в настрйках
                radio_btn.setChecked(True) # делаем её активной

    def save_radio_group(self): # сохранить статус группы

        sender = self.sender() # получаем объект, который отправил сигнал

        radio_btn = sender.objectName() # получаем имя радиокнопки

        self.temp_states[self.group_name] = radio_btn # сохраняем состояние радиокнопки во временный словарь

    def set_lines(self): # заполнить путь предыдущими значениями и проверка изменения

        self.path_lineEdit.textChanged.connect(self.check_folder_path) # проверка изменений в поле и активация кнопки

        old_path_lineEdit = str(settings_loader.settings_val("server_path")) # сохранённая старая строк

        if old_path_lineEdit not in ("None", ""): # если есть сохранённые значения
            self.path_lineEdit.setText(old_path_lineEdit) # прописать значение

    def check_folder_path(self): # проверка только пути к папке

        current_font = self.path_lineEdit.font() # cохраняем текущий шрифт

        path = self.sanitize_path(self.path_lineEdit.text()) # очищаем путь от запрещённых символов

        if path != self.path_lineEdit.text(): # обновляем поле ввода очищенным путем (если он изменился)
            self.update_line_edit_text(self.path_lineEdit, path) # обновление текста с сохранением позиции курсора

        self.is_folder_valid = self.checking_empty_folder(path)
        self.path_lineEdit.setStyleSheet("" if self.is_folder_valid else "background-color: #ffdddd;") # меняем цвет строки

        self.temp_states["server_path"] = path # сохраняем состояние радиокнопки во временный словарь

        self.path_lineEdit.setFont(current_font) # восстанавливаем шрифт

    def sanitize_path(self, path): # очищаем путь от запрещённых символов

        path = path.strip()
        path = re.sub(r'[*?"<>|]', '', path)  # Удаляем запрещённые символы
        path = path.replace('file:///', '')  # Удаляем URL-префикс

        return path.strip('"')  # Удаляем обрамляющие кавычки

    def update_line_edit_text(self, line_edit, new_text): # обновление текста с сохранением позиции курсора

        line_edit.blockSignals(True) # временно блокируем сигналы
        cursor_pos = line_edit.cursorPosition() # Сохранение позиции курсора
        line_edit.setText(new_text) # прописать значение
        line_edit.setCursorPosition(cursor_pos) # Сохранение позиции курсор
        line_edit.blockSignals(False) # разблокируем сигналы

    def checking_empty_folder(self, path): # проверка пустай ли папка

        if path =="": # если строка не заполнена
            return False # прервать

        if not os.path.isdir(path) and not re.match(r'^[\.\/\\]+$', path): # если папка существует и нет символов в начале (для правильной проверки)
            return False # прервать

        return True # возвращаем значение

    def open_file_directory(self): # Вызов QFileDialog для выбора папки

        directory_path = QFileDialog.getExistingDirectory(
            self,
            "Выберите папку",  # Заголовок окна
            ""                 # Начальная директория (пустая строка — текущая директория)
        )

        if directory_path:
            directory_path = directory_path.replace('/', '\\')  # изменяем наклон слэша
            print("Выбрана папка:", directory_path)  # Действия с выбранным файлом
            self.path_lineEdit.setText(directory_path) # прописать в строку

    def closeEvent(self, event): # действия после закрытия окна на крестик

        for checkbox_name, checked in self.temp_states.items(): # запись всех изменённых чекбоксов
            settings_loader.save_settings_val(checkbox_name, checked) # cохраняем состояние в настройки

        self.reject() # закрываем модальное диалоговое окно

class KompasWorkerThread(QThread):

    text_line_msg_signal = pyqtSignal(str, str) # сигнал об изменении строки (текст, цвет)

    initialized_signal = pyqtSignal(bool, str) # сигнал об инициализации

    def __init__(self, parent=None, settings_val=None):

        super().__init__(parent)

        self.command_queue = []  # очередь команд

    def run(self): # основной цикл потока

        try:
            CoInitialize() # инициализация COM в потоке

            if self.kompasAPI(): # подключение API компаса

                self.initialized_signal.emit(True, "") # посылаем сигнал

                while True: # Основной цикл обработки команд

                    if self.command_queue: # если есть очередь команд
                        command, data, callback = self.command_queue.pop(0) # извлечь данные
                        result = self.execute_command(command, data) # выполнение команы

                        if result == "close_kompas": # если закрыли программу
                            break # прервать цикл

                    else:
                        self.msleep(200) # небольшая пауза чтобы не грузить CPU

        except Exception:
            msg = f"Произошла ошибка:\n {traceback.format_exc()}"
            self.initialized_signal.emit(False, msg)

        finally:
            self.cleanup() # очистка при завершении потока

    def kompasAPI(self): # подключение API компаса

        try: # попытаться подключиться к КОМПАСу

            self.text_line_msg_signal.emit("Идёт подключение к КОМПАСу!", "") # посылаем сигнал (текст, цвет)

            try: # попытаться подключиться к КОМПАСу
                connect('Kompas.Application.7') # если подключились
                self.kompas_run = True # был запущен

            except Exception:
                self.kompas_run = False # не был запущен

            KAPI5 = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0) # API5 КОМПАСа
            self.iKompasObject = KAPI5.KompasObject(connect('Kompas.Application.5')) # подключение к запущенному экземпляру КОМПАСа
##            self.iKompasObject = Dispatch("Kompas.Application.5", None, KAPI5.KompasObject.CLSID) # интерфейс API КОМПАС

            self.KAPI7 = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0) # API7 КОМПАСа
            self.iApplication = self.KAPI7.IApplication(connect('Kompas.Application.7')) # подключение к запущенному экземпляру КОМПАСа
##            self.iApplication = Dispatch("Kompas.Application.7") # интерфейс приложения КОМПАС-3D.

            self.iApplication.Visible = True # сделать КОМПАС-3D видимым

            if self.iApplication.HideMessage != 0: # диалоговые кокна КОМПАСа (0 - показывать все сообщения; 1 - с перестроением/сохраненем; 2 - без перестроения/сохранения)
                self.iApplication.HideMessage = 0 # диалоговые кокна КОМПАСа (0 - показывать все сообщения; 1 - с перестроением/сохраненем; 2 - без перестроения/сохранения)

            self.text_line_msg_signal.emit("", "") # посылаем сигнал (текст, цвет)

            return True

        except Exception: # если не получилось подключиться к КОМПАСу
            msg = "Не удалось подключиться к КОМПАС-3D!\nЗапустите или перезапустите КОМПАС-3D!"
            self.initialized_signal.emit(False, msg) # посылаем сигнал

    def add_command(self, command, data=None, callback=None): # добавление команды в очередь (вызывается из основного потока)
        self.command_queue.append((command, data, callback)) # добавляем в очередь

    def execute_command(self, command, data): # выполнение команы

        if command == "select_item": # команда прописапть ТТ
            return self.select_item(data) # прописываем в строку ТТ

        elif command == "clear_line": # очистка послетней строки ТТ
            return self.clear_line() # очистка послетней строки ТТ

        elif command == "clear": # очистка ТТ
            return self.clear() # очистка ТТ

        elif command == "close_kompas": # команда завершить поток
            return "close_kompas"

        else:
            return f"Unknown command: {command}"

    def select_item(self, text): # выбираем и прописываем в строку ТТ

        line = self.listbox_text_processing(text) # обработка текста списка (не вводим лишнее в ТТ)

        if not line: # если нет строки
            return # прервать

        if isinstance(line, list): # если значение параметра список, обработать каждое значение
            line_count = len(line) # количество строк
            line = "\n".join(line) # строки разделённые знаком переноса

        else:
            line_count = 0 # нет строк

        iKompasDocument = self.iApplication.ActiveDocument # делаем активный открытый документ

        if iKompasDocument is None or iKompasDocument.DocumentType != 1: # если нет открытого документа и этот документ не чертёж
            self.kompas_message("Откройте чертёж!") # сообщение в окне компаса если он открыт

        else:
            iDrawingDocument = self.KAPI7.IDrawingDocument(iKompasDocument) # интерфейс чертежа
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
##            iTextFont = self.KAPI7.ITextFont(iTextItem) # интерфейс параметров шрифта
##            iTextFont.Bold = False # жирный шрифт
##            iTextFont.Color = 0 # цвет
##            iTextFont.FontName = 'GOST type A' # имя шрифта
##            iTextFont.Height = 5 # высота текта
##            iTextFont.Italic = True # курсив
##            iTextFont.Underline = False # подчёркивание
##            iTextFont.WidthFactor = 1 # коэффициент сужения
            iTextItem.Update() # обновить данные компонента

            if line_count > 1:
                iCount = iText.Count # количество строк

                for i in range(1, line_count): # для каждой строки
                    line_index = iCount - i

                    if line_index >= 0:  # Проверяем, что индекс существует
                        iTextLine = iText.TextLine(line_index) # указываем на последнюю строку (отсчёт с 0)
                        iTextLine.Numbering = 3 # тип нумерации абзаца (-1 - Тип не определенный; 0 - Cрока без нумерации; 1 - Cтрока с нумерацией уровня level; 2 - На строке начинается новая нумерация пунктов; 3 - Cтрока не должна нумероваться никогда)

            iTechnicalDemand.Update() # обновить данные ТТ

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

    def kompas_message(self, text): # сообщение в окне КОМПАСа если он открыт

        if self.iApplication.Visible: # если компас видимый
            self.iApplication.MessageBoxEx(text, 'Message:', 64) # сообщение в КОМПАСе

    def clear_line(self): # очистка послетней строки ТТ

        iKompasDocument = self.iApplication.ActiveDocument # делаем активный открытый документ

        if iKompasDocument is None or iKompasDocument.DocumentType != 1: # если нет открытого документа и этот документ не чертёж
            self.kompas_message("Откройте чертёж!") # сообщение в окне компаса если он открыт

        else:
            iDrawingDocument = self.KAPI7.IDrawingDocument(iKompasDocument) # интерфейс чертежа
            iTechnicalDemand = iDrawingDocument.TechnicalDemand # интерфейс технических требований
            iText = iTechnicalDemand.Text # интерфейс текста ТТ
            Count = iText.Count # количество строк
            iTextLine = iText.TextLine(Count - 1) # указываем на последнюю строку (отсчёт с 0)

            iTextItem = iTextLine.TextItem(0) # указываем на последнюю строку (отсчёт с 0)
            NewLine = iTextItem.NewLine # признак начала строки

            iTextLine.Delete() # удалить строку текста (всю строку с текстом)
            iTechnicalDemand.Update() # обновить данные ТТ

            while not NewLine: # если строка не начало, удаляем её

                Count = iText.Count # количество строк
                iTextLine = iText.TextLine(Count - 1) # указываем на последнюю строку (отсчёт с 0)

                iTextItem = iTextLine.TextItem(0) # указываем на последнюю строку (отсчёт с 0)
                NewLine = iTextItem.NewLine # признак начала строки
                iTextLine.Delete() # удалить строку текста (всю строку с текстом)

            iTechnicalDemand.Update() # обновить данные ТТ

    def clear(self): # очистка ТТ

        iKompasDocument = self.iApplication.ActiveDocument # делаем активный открытый документ

        if iKompasDocument is None or iKompasDocument.DocumentType != 1: # если нет открытого документа и этот документ не чертёж
            self.kompas_message("Откройте чертёж!") # сообщение в окне компаса если он открыт

        else:
            iDrawingDocument = self.KAPI7.IDrawingDocument(iKompasDocument) # интерфейс чертежа
            iTechnicalDemand = iDrawingDocument.TechnicalDemand # интерфейс технических требований
            iText = iTechnicalDemand.Text # интерфейс текста ТТ
            iText.Clear() # очистить ТТ
            iTechnicalDemand.Update() # обновить данные ТТ

    def cleanup(self): # очистка при завершении потока

        try:
            if not self.kompas_run and self.iApplication: # если КОМПАС-3D был запущен программой и активен
                self.iApplication.HideMessage = 2 # диалоговые кокна КОМПАСа (0 - показывать все сообщения; 1 - с перестроением/сохраненем; 2 - без перестроения/сохранения)
                self.iApplication.Quit() # закрыть КОМПАС

            CoUninitialize() # закрытие инициализации COM в потоке

        except Exception:
            pass

    def quit(self): # остановка рабочего потока
        self.add_command("close_kompas") # запустить команду закрытия

#-------------------------------------------------------------------------------

if __name__ == "__main__":  # если мы запускаем файл напрямую, а не импортируем

    try: # попытаться запустить интерфейс

        message_manager = MessageManager() # глобальный экземпляр менеджера сообщений

        settings_loader = SettingsLoader() # загружаем настройки из INI-файла

        main() # запускаем функцию main()

    except Exception:
        msg = f"Произошла ошибка:\n {traceback.format_exc()}"
        message(msg, 10, "error") # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия)