import sys
import aiohttp
import os
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QLineEdit, QPushButton, QVBoxLayout, QFileDialog, QStyle, \
    QTextEdit, QGridLayout, QMenu, QAction, QMainWindow, QMenuBar, QMessageBox, QDialog, QInputDialog
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import QThread, pyqtSignal, QLocale
import fitz
import asyncio
import pandas as pd
import csv
from docx import Document
import getpass
from datetime import datetime
import requests
import subprocess

def check_for_updates(current_version):
    # Получение последней версии из репозитория на GitHub
    repo_url = 'https://api.github.com/repos/yourusername/yourrepository/releases/latest'
    response = requests.get(repo_url)
    latest_version = response.json()['tag_name']

    if latest_version > current_version:
        return latest_version
    else:
        return None

def update_application():
    # Скачивание и применение обновлений с использованием Git
    subprocess.run(['git', 'pull'])

if __name__ == '__main__':
    current_version = '1.0.0'  # Версия вашего текущего приложения

    latest_version = check_for_updates(current_version)

    if latest_version:
        print(f'Доступно обновление ({latest_version})! Хотите обновить приложение? (да/нет)')
        response = input()
        if response.lower() == 'да':
            update_application()
            print('Приложение обновлено!')
        else:
            print('Обновление отменено.')
    else:
        print('У вас установлена последняя версия приложения.')

class WorkerThread(QThread):
    finished = pyqtSignal(str)

    def __init__(self, func, *args, **kwargs):
        super().__init__()
        self.func = func
        self.args = args
        self.kwargs = kwargs

    def run(self):
        result = self.func(*self.args, **self.kwargs)
        self.finished.emit(result)

class PDFclass(QDialog):
    def __init__(self, parent=None):
        super(PDFclass, self).__init__(parent)

        # Создание кнопки и привязка действий к кнопкам
        self.btn_count_folders = QPushButton('Роли в разных папках', self)
        self.btn_count_folders.clicked.connect(self.count_pages_for_folders)

        self.btn_count_single_folder = QPushButton('Роли в одной папке', self)
        self.btn_count_single_folder.clicked.connect(self.process_single_folder)

        self.result_text = QTextEdit(self)
        self.result_text.setLineWrapMode(QTextEdit.WidgetWidth)
        self.result_text.setReadOnly(True)

        # добавление кнопок на графику
        layout = QVBoxLayout(self)
        layout.addWidget(self.btn_count_folders)
        layout.addWidget(self.btn_count_single_folder)
        layout.addWidget(self.result_text)


    async def count_total_pages(self, file):
        pdf_document = fitz.open(file)
        total_pages = pdf_document.page_count
        pdf_document.close()
        return total_pages

    async def count_pages_in_single_folder(self, folder):
        files = [os.path.join(folder, file_name) for file_name in os.listdir(folder) if
                 file_name.endswith('.pdf') and os.path.isfile(os.path.join(folder, file_name))]
        page_counts = {}
        for file in files:
            pdf_document = fitz.open(file)
            page_counts[os.path.basename(file)] = pdf_document.page_count
            pdf_document.close()
        return page_counts

    async def async_count_pages_in_single_folder(self, folder):
        folder_name = os.path.basename(folder)
        page_counts = await self.count_pages_in_single_folder(folder)
        result_text = f'{folder_name}\n'
        for file, pages in page_counts.items():
            result_text += f'{file}: {pages}\n'
        self.result_text.insertPlainText(result_text)
        self.save_to_file(result_text)

    def process_single_folder(self):
        folder = QFileDialog.getExistingDirectory(self, 'Выбрать папку с PDF-файлами')
        if folder:
            loop = asyncio.get_event_loop()
            loop.run_until_complete(self.async_count_pages_in_single_folder(folder))

    def select_folders(self):
        folders = []
        folder = QFileDialog.getExistingDirectory(self, 'Выбрать папку с PDF-файлами')
        while folder:
            folders.append(folder)
            folder = QFileDialog.getExistingDirectory(self, 'Выбрать еще папку с PDF-файлами')
        return folders

    async def async_count_pages_for_folders(self):
        folders = self.select_folders()
        self.result_text.clear()
        result_text = ''
        for folder in folders:
            folder_name = os.path.basename(folder)
            files = [os.path.join(folder, file_name) for file_name in os.listdir(folder) if
                     file_name.endswith('.pdf') and os.path.isfile(os.path.join(folder, file_name))]
            total_pages = await asyncio.gather(*(self.count_total_pages(file) for file in files))
            result_text += f'{folder_name}: {sum(total_pages)}\n'
        self.result_text.insertPlainText(result_text)
        self.save_to_file(result_text)

    def count_pages_for_folders(self):
        loop = asyncio.get_event_loop()
        loop.run_until_complete(self.async_count_pages_for_folders())

    def save_to_file(self, text):
        if not text:  # Проверяем, что текст не пустой
            return

        desktop_path = os.path.join(os.path.join(os.path.expanduser('~'), 'Desktop'))
        file_name = f"Результат_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.txt"
        file_path = os.path.join(desktop_path, file_name)
        with open(file_path, 'w', encoding='utf-8') as file:
            file.write(text)
        self.result_text.clear()
        self.result_text.insertPlainText(f'Результат сохранен в файле: {file_name}')



class SearchAndCopyDialog(QDialog):
    def __init__(self, parent=None):
        super(SearchAndCopyDialog, self).__init__(parent)

        self.label_search = QLabel('Введите номера пачек для поиска (через пробел или диапазон через тире):')
        self.search_entry = QLineEdit()

        self.label_save = QLabel('Введите имя файла для сохранения:')
        self.save_entry = QLineEdit()

        self.btn_search_copy = QPushButton('Выполнить поиск и копирование')
        self.btn_search_copy.clicked.connect(self.search_and_copy)

        layout = QVBoxLayout(self)
        layout.addWidget(self.label_search)
        layout.addWidget(self.search_entry)
        layout.addWidget(self.label_save)
        layout.addWidget(self.save_entry)
        layout.addWidget(self.btn_search_copy)

    def search_and_copy(self):
            file_path, _ = QFileDialog.getOpenFileName(self, 'Выберите текстовый файл')
            if not file_path:
                QMessageBox.warning(self, 'Ошибка', 'Файл не выбран.')
                return

            with open(file_path, 'r') as file:
                content = file.readlines()

            search_numbers = self.search_entry.text().strip().split()
            matching_lines = []

            for line in content:
                columns = line.strip().split(';')  # Разбиваем строку на столбцы с разделителем ;
                if len(columns) >= 2:  # Проверяем, что в строке есть как минимум два столбца
                    # Поиск номеров пачек только в первом и втором столбце
                    for search_number in search_numbers:
                        if "-" in search_number:
                            start, end = map(int, search_number.split("-"))
                            if any(start <= int(col) <= end for col in columns[:2]):
                                matching_lines.append(line)
                        else:
                            if any(search_number == col for col in columns[:2]):
                                matching_lines.append(line)

            if not matching_lines:
                QMessageBox.information(self, 'Результат', 'Совпадений не найдено.')
                return

            directory_path = QFileDialog.getExistingDirectory(self, 'Выберите папку для сохранения файла')
            if not directory_path:
                QMessageBox.warning(self, 'Ошибка', 'Папка не выбрана.')
                return

            file_name = self.save_entry.text()
            new_file_path = os.path.join(directory_path, f"{file_name}.txt")

            with open(new_file_path, 'w') as new_file:
                new_file.writelines(matching_lines)

            QMessageBox.information(self, 'Успех', 'Новый файл успешно создан и сохранен.')

            # После сохранения файла спрашиваем пользователя о повторении операции
            reply = QMessageBox.question(self, 'Повторить поиск', 'Хотите выполнить поиск и копирование еще раз?',
                                         QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)

            if reply == QMessageBox.Yes:
                self.clear_input_fields()
                self.search_and_copy()
            else:
                self.close()

class IntegratedApp(QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()

    # Создание кнопок, лейблов
    def init_ui(self):

        self.setWindowTitle('doesnt work')
        self.setWindowIcon(QIcon('logobbs.ico'))

        self.btn_night_shift = QPushButton('Ночная смена', self)
        self.btn_night_shift.clicked.connect(self.night_shift)

        self.btn_search_copy = QPushButton('Поиск пачек для лото', self)
        self.btn_search_copy.clicked.connect(self.open_search_and_copy_dialog)

        self.btn_PDFWORKED = QPushButton('Работа с PDF-файлами', self)
        self.btn_PDFWORKED.clicked.connect(self.PDFWork)

        self.btn_create_csv = QPushButton('Создать базу данных для старого портала', self)
        self.btn_create_csv.clicked.connect(self.create_csv_file)

        self.result_text = QTextEdit(self)
        self.result_text.setReadOnly(True)

        # Расположение layout'ов
        grid_layout = QGridLayout()
        grid_layout.addWidget(self.btn_night_shift, 1, 0, 1, 2)
        grid_layout.addWidget(self.btn_search_copy, 2, 0, 1, 2)
        grid_layout.addWidget(self.btn_PDFWORKED, 3, 0, 1, 2)
        grid_layout.addWidget(self.btn_create_csv, 4, 0, 1, 2)
        grid_layout.addWidget(self.result_text, 5, 0, 1, 2)
        self.setLayout(grid_layout)

    # Создание docx ночной смены
    def night_shift(self):
        import os
        from docx import Document
        import getpass
        from datetime import datetime

        # Определяем рабочий стол пользователя
        desktop_path = os.path.join(os.path.join(os.path.expanduser('~'), 'Desktop'))

        # Открываем документ
        doc = Document(r'C:\Users\kozlov\PycharmProjects\dogovornoch\BBS_prikaz_noch.docx')

        # Словарь для преобразования английских названий месяцев в русские
        months_translation = {
            'January': 'января',
            'February': 'февраля',
            'March': 'марта',
            'April': 'апреля',
            'May': 'мая',
            'June': 'июня',
            'July': 'июля',
            'August': 'августа',
            'September': 'сентября',
            'October': 'октября',
            'November': 'ноября',
            'December': 'декабря'
        }

        # Получаем текущую дату
        today = datetime.today()
        day = today.strftime('%d')  # Получаем текущий день
        month = months_translation[
            today.strftime('%B')]  # Получаем текущий месяц, полное название с переводом и склонением
        year = today.strftime('%Y')  # Получаем текущий год
        file_name = f'{day}_{month}.docx'  # Формируем имя файла

        # Получаем имя пользователя
        user_name = getpass.getuser()

        # Если юзернейм такой то, то и фио такое же
        if user_name == 'Evsyukov':
            surname = 'Евсюков Александр Юрьевич'
        elif user_name == 'kozlov':
            surname = 'Козлов Максим Александрович'
        elif user_name == 'vdovidchenko':
            surname = 'Вдовидченко Артем Романович'
        elif user_name == 'kondrashov':
            surname = 'Кондрашов Сергей Витальевич'
        elif user_name == 'mamaev':
            surname = 'Мамаев Виктор А.'

        # Заменяем метки в документе на значения
        for para in doc.paragraphs:
            para.text = para.text.replace('{day}', day)
            para.text = para.text.replace('{month}', month)
            para.text = para.text.replace('{year}', year)
            para.text = para.text.replace('{surname}', surname)

        # Сохраняем изменения
        output_path = os.path.join(desktop_path, file_name)
        doc.save(output_path)

    # Вызов окна для работы с PDF-файлами
    def PDFWork(self):
        dialog = PDFclass(self)
        dialog.exec_()

    # Вызов окна для работы с Лото
    def open_search_and_copy_dialog(self):
        dialog = SearchAndCopyDialog(self)
        dialog.exec_()

    # Работа с .csv файлами
    def create_csv_file(self):
            # Получаем ввод от пользователя
            max_number = int(self.input_dialog("Введите количество коробов: "))
            second_number = int(self.input_dialog("Введите данные ШК 1: "))
            third_number = int(self.input_dialog("Введите данные ШК 2: "))
            fourth_number = int(self.input_dialog("Введите подсказ ШК 1: "))
            fifth_number = int(self.input_dialog("Введите подсказ ШК 2: "))
            step = int(self.input_dialog("Введите количество в коробе: "))
            secondd_number = second_number
            thirdd_number = third_number

            # Создаем пустой список для данных
            data = []

            for i in range(2, max_number + 2):  # Начинаем с 2, чтобы первое число было 2
                row = [i,
                       second_number + (i - 1) * step,
                       third_number + (i - 1) * step,
                       secondd_number + (i - 1) * step,
                       thirdd_number + (i - 1) * step,
                       fourth_number + (i - 1) * step,
                       fifth_number + (i - 1) * step,
                       step]
                data.append(row)

            # Записываем данные в CSV файл
            csv_file_path, _ = QFileDialog.getSaveFileName(self, 'Сохранить CSV файл', filter="CSV Files (*.csv)")
            if not csv_file_path:
                self.result_text.setPlainText("Файл не сохранен.")
                return

            with open(csv_file_path, 'w', newline='') as csvfile:
                csvwriter = csv.writer(csvfile, delimiter=';')  # Используем точку с запятой как разделитель
                csvwriter.writerow([
                    '1', second_number, third_number, secondd_number, thirdd_number, fourth_number, fifth_number,
                    step
                ])  # Запись данных пользователя
                csvwriter.writerows(data)

            # Записываем данные в CSV файл
            with open(csv_file_path, 'a', newline='') as csvfile:
                csvwriter = csv.writer(csvfile, delimiter=';')  # Используем точку с запятой как разделитель
                csvwriter.writerow([
                    '1', second_number, third_number, secondd_number, thirdd_number, fourth_number, fifth_number,
                    step
                ])  # Запись данных пользователя

            # Удаляем последнюю строку из CSV файла
            df = pd.read_csv(csv_file_path, delimiter=';')  # Используем точку с запятой как разделитель
            df = df[:-2]  # Удаляем лишние строки
            df.to_csv(csv_file_path, index=False, sep=';')  # Используем точку с запятой как разделитель

            self.result_text.setPlainText('Файл создан.(Не забудьте убрать с первой строки десятичные значения)')

    def input_dialog(self, text):
        input_dialog = QInputDialog(self)
        input_dialog.setInputMode(QInputDialog.TextInput)

        # Получаем поле ввода (QLineEdit)
        line_edit = input_dialog.findChild(QLineEdit)

        # максимальная длина ввода
        line_edit.setMaxLength(14)

        input_dialog.setWindowTitle('Ввод данных')
        input_dialog.setLabelText(text)
        input_dialog.resize(400, 200)
        input_dialog.exec_()

        value = input_dialog.textValue()
        if value:
            return int(value)
        else:
            self.result_text.setPlainText("Ввод отменен")
            self.close()
            return None

    def save_to_file(self, text):
            desktop_path = os.path.join(os.path.join(os.path.expanduser('~'), 'Desktop'))
            file_name = f"Результат_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.txt"
            file_path = os.path.join(desktop_path, file_name)
            with open(file_path, 'w', encoding='utf-8') as file:
                file.write(text)
            self.result_text.setPlainText(f'Результат сохранен в файле: {file_name}')

class MainMenu(QMainWindow):
    def __init__(self):
        super().__init__()

        # Create actions
        exit_action = QAction('Выход', self)
        exit_action.setShortcut('Ctrl+Q')
        exit_action.setStatusTip('Выход из приложения')
        exit_action.triggered.connect(self.close)
        helps = QAction('Помощь', self)
        helps.setStatusTip('Инструкции')
        helps.triggered.connect(self.show_help_dialog)

        # Create menu bar
        menubar = self.menuBar()
        file_menu = menubar.addMenu('Меню')
        file_menu.addAction(exit_action)
        file_menu.addAction(helps)

        self.central_widget = IntegratedApp()
        self.setCentralWidget(self.central_widget)

    def show_help_dialog(self):
        help_dialog = QDialog(self)
        help_dialog.setWindowTitle('Помощь')

        layout = QVBoxLayout()

        label = QLabel('Привет', help_dialog)
        layout.addWidget(label)

        help_dialog.setLayout(layout)
        help_dialog.exec_()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    locale = QLocale(QLocale.Russian)
    QLocale.setDefault(locale)
    main_menu = MainMenu()
    main_menu.show()
    sys.exit(app.exec_())
