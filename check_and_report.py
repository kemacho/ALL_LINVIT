import os
import glob
import threading
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
                               QLineEdit, QPushButton, QComboBox, QProgressBar, QMessageBox, QFileDialog)
from PySide6.QtCore import Qt, Signal, QObject

# Названия подпапок СИ
nameSI = ['0 Заявка и приложение', '1 Распоряжение по заявке', '2 Решение по заявке', '3 Заключения по ОМД и ТД',
          '4 Акт выбора ПК', '5 Протоколы СИ', '6 Заключение СИ', '7 Программа проверки произ', '8 Акт ПП',
          '9 Распоряжение на анализ', '10 Решение о выдаче', '11 Сертификат', '12 Доп.материалы']

# Названия подпапок ИК
nameIK = ['0 Распоряжение', '1 Письмо-уведомление', '2 Программа ИК', '3 Программа проверки произ', '4 Акт выбора ПК',
          '5 Акт проверки производства', '6 Протоколы ИК', '7 Акт по результатам ИК', '8 Распоряжение на анализ',
          '9 Решение по ИК', '10 Доп. материалы']

# Шаблоны
SI_TEMPLATES = {
    "Шаблон РЖД": ['2', '1', '1', '4', '1', 'Any', '1', '1', '1', '1', '2', '2', 'Any'],
    "Шаблон 2024": ['1', '1', '1', '4', '1', 'Any', '1', '1', '1', '1', '1', '1', 'Any']
}

IK_TEMPLATES = {
    "Шаблон РЖД": ['1', '1', '1', '1', '1', '1', 'Any', '1', '1', '1', 'Any'],
    "Общий шаблон": ['1', '1', '1', '1', '1', '1', 'Any', '1', '1', '1', 'Any']
}


class WorkerSignals(QObject):
    progress = Signal(int)
    message = Signal(str)
    finished = Signal(str)
    error = Signal(str)


class FolderProcessor(threading.Thread):
    def __init__(self, inpath, fileqnt, fileqnt_ik):
        super().__init__()
        self.inpath = inpath
        self.fileqnt = fileqnt
        self.fileqnt_ik = fileqnt_ik
        self.signals = WorkerSignals()
        self.worry = []

    def run(self):
        try:
            self.process_folders()
            excel_path = self.create_excel_report()
            self.signals.finished.emit(excel_path)
        except Exception as e:
            self.signals.error.emit(str(e))

    def worrymessage(self, pt, nm1, nm2):
        self.worry.append(pt)
        self.worry.append(nm1)
        self.worry.append(nm2)

    def process_folder_with_one_file(self, old_folder, contents, Pos_dest, Neg_dest):
        if len(contents) == 1 and old_folder != Pos_dest:
            os.rename(old_folder, Pos_dest)
        elif len(contents) == 0 and old_folder != Neg_dest:
            os.rename(old_folder, Neg_dest)
        elif len(contents) > 1:
            os.rename(old_folder, Pos_dest)
            self.worrymessage(old_folder, len(contents), '1')

    def process_folder_with_four_files(self, old_folder, contents, Pos_dest, Neg_dest):
        if len(contents) == 4 and old_folder != Pos_dest:
            os.rename(old_folder, Pos_dest)
        elif len(contents) < 4 and old_folder != Neg_dest:
            os.rename(old_folder, Neg_dest)
        elif len(contents) > 4:
            os.rename(old_folder, Pos_dest)
            self.worrymessage(old_folder, len(contents), '4')

    def process_folder_with_any_files(self, old_folder, contents, Pos_dest, Neg_dest):
        if len(contents) > 0 and old_folder != Pos_dest:
            os.rename(old_folder, Pos_dest)
        elif len(contents) == 0 and old_folder != Neg_dest:
            os.rename(old_folder, Neg_dest)
        elif len(contents) > 20:
            os.rename(old_folder, Pos_dest)
            self.worrymessage(old_folder, len(contents), 'возможно не так много')

    def process_folder_with_two_files(self, old_folder, contents, Pos_dest, Norm_dest, Neg_dest):
        if len(contents) == 2 and old_folder != Pos_dest:
            os.rename(old_folder, Pos_dest)
        elif len(contents) == 1 and old_folder != Norm_dest:
            os.rename(old_folder, Norm_dest)
        elif len(contents) == 0 and old_folder != Neg_dest:
            os.rename(old_folder, Neg_dest)
        elif len(contents) > 2:
            os.rename(old_folder, Pos_dest)
            self.worrymessage(old_folder, len(contents), '2')

    def check(self, FileQNT, old_folder, contents, Pos_dest, Norm_dest, Neg_dest):
        filtered_contents = [f for f in contents if f != "Thumbs.db"]

        if FileQNT == "1":
            self.process_folder_with_one_file(old_folder, filtered_contents, Pos_dest, Neg_dest)
        if FileQNT == "2":
            self.process_folder_with_two_files(old_folder, filtered_contents, Pos_dest, Norm_dest, Neg_dest)
        if FileQNT == "4":
            self.process_folder_with_four_files(old_folder, filtered_contents, Pos_dest, Neg_dest)
        elif FileQNT == 'Any':
            self.process_folder_with_any_files(old_folder, filtered_contents, Pos_dest, Neg_dest)

    def process_folders(self):
        try:
            folder_names = os.listdir(self.inpath)
        except Exception as e:
            self.signals.error.emit(f"Ошибка при чтении папки: {e}")
            return

        total_folders = len(folder_names)
        for i, name in enumerate(folder_names):
            pathSI = os.path.join(self.inpath, name, '0. СИ')
            pathIK1 = os.path.join(self.inpath, name, '1. ИК-1')
            pathIK2 = os.path.join(self.inpath, name, '2. ИК-2')

            # Проверка для папки СИ
            for j in range(len(self.fileqnt)):
                nameSIzv = nameSI[j] + '*'
                try:
                    old_folder = glob.glob(os.path.join(pathSI, nameSIzv))
                    if old_folder:
                        old_folder = old_folder[0]
                        contents = os.listdir(old_folder)

                        folder = os.path.join(pathSI, nameSI[j])
                        Pos_dest = str(folder) + ' (+)'
                        Norm_dest = str(folder) + ' (+—)'
                        Neg_dest = str(folder) + ' (—)'

                        self.check(self.fileqnt[j], old_folder, contents, Pos_dest, Norm_dest, Neg_dest)
                except Exception as e:
                    print(f"Ошибка при обработке {nameSI[j]} в {name}: {e}")

            # Проверка для папки ИК1
            for j in range(len(self.fileqnt_ik)):
                nameIK1zv = nameIK[j] + '*'
                try:
                    old_folder = glob.glob(os.path.join(pathIK1, nameIK1zv))
                    if old_folder:
                        old_folder = old_folder[0]
                        contents = os.listdir(old_folder)

                        folder = os.path.join(pathIK1, nameIK[j])
                        Pos_dest = str(folder) + ' (+)'
                        Norm_dest = str(folder) + ' (+—)'
                        Neg_dest = str(folder) + ' (—)'

                        self.check(self.fileqnt_ik[j], old_folder, contents, Pos_dest, Norm_dest, Neg_dest)
                except Exception as e:
                    print(f"Ошибка при обработке {nameIK[j]} в {name} (ИК1): {e}")

            # Проверка для папки ИК2
            for j in range(len(self.fileqnt_ik)):
                nameIK2zv = nameIK[j] + '*'
                try:
                    old_folder = glob.glob(os.path.join(pathIK2, nameIK2zv))
                    if old_folder:
                        old_folder = old_folder[0]
                        contents = os.listdir(old_folder)

                        folder = os.path.join(pathIK2, nameIK[j])
                        Pos_dest = str(folder) + ' (+)'
                        Norm_dest = str(folder) + ' (+—)'
                        Neg_dest = str(folder) + ' (—)'

                        self.check(self.fileqnt_ik[j], old_folder, contents, Pos_dest, Norm_dest, Neg_dest)
                except Exception as e:
                    print(f"Ошибка при обработке {nameIK[j]} в {name} (ИК2): {e}")

            progress = int((i + 1) / total_folders * 100)
            self.signals.progress.emit(progress)
            self.signals.message.emit(f"Обработано папок: {i + 1} из {total_folders}")

        if self.worry:
            worry_messages = []
            for i in range(0, len(self.worry), 3):
                msg = f'Пожалуйста проверьте папку: {self.worry[i]}, там находится {self.worry[i + 1]} файла, вместо {self.worry[i + 2]}'
                worry_messages.append(msg)
            self.signals.message.emit("\n".join(worry_messages))

    def create_excel_report(self):
        wb = Workbook()
        if 'Sheet' in wb.sheetnames:
            wb.remove(wb['Sheet'])

        sheets = {
            'СИ': nameSI,
            'ИК-1': nameIK,
            'ИК-2': nameIK
        }

        for sheet_name, headers in sheets.items():
            ws = wb.create_sheet(title=sheet_name)
            ws.column_dimensions['A'].width = 50
            for col in range(2, len(headers) + 2):
                ws.column_dimensions[chr(64 + col)].width = 15

            headers_with_name = ['Название папки'] + headers
            ws.row_dimensions[1].height = 30

            for col_num, header in enumerate(headers_with_name, start=1):
                cell = ws.cell(row=1, column=col_num, value=header)
                cell.font = Font(bold=True)
                cell.alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')

        folder_names = os.listdir(self.inpath)
        for folder_name in folder_names:
            pathSI = os.path.join(self.inpath, folder_name, '0. СИ')
            if os.path.exists(pathSI):
                self.process_folder_for_excel(wb['СИ'], folder_name, pathSI, nameSI)

            pathIK1 = os.path.join(self.inpath, folder_name, '1. ИК-1')
            if os.path.exists(pathIK1):
                self.process_folder_for_excel(wb['ИК-1'], folder_name, pathIK1, nameIK)

            pathIK2 = os.path.join(self.inpath, folder_name, '2. ИК-2')
            if os.path.exists(pathIK2):
                self.process_folder_for_excel(wb['ИК-2'], folder_name, pathIK2, nameIK)

        output_path = os.path.join(self.inpath, "results.xlsx")
        wb.save(output_path)
        return output_path

    def process_folder_for_excel(self, worksheet, folder_name, base_path, subfolder_names):
        row_num = worksheet.max_row + 1 if worksheet.max_row > 1 else 2
        worksheet.cell(row=row_num, column=1, value=folder_name)

        for i, subfolder in enumerate(subfolder_names, start=2):
            subfolder_path = os.path.join(base_path, subfolder + '*')
            matched_folders = glob.glob(subfolder_path)

            if not matched_folders:
                worksheet.cell(row=row_num, column=i, value="Нет папки")
                continue

            folder = matched_folders[0]
            status = self.get_folder_status(folder)
            worksheet.cell(row=row_num, column=i, value=status)

    def get_folder_status(self, folder_path):
        folder_name = os.path.basename(folder_path)
        if ' (+)' in folder_name:
            return '+'
        elif ' (—)' in folder_name:
            return '-'
        elif ' (+—)' in folder_name:
            return '+-'
        else:
            return '?'


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Обработчик папок")
        self.setMinimumSize(800, 600)

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.layout = QVBoxLayout(self.central_widget)

        # Выбор корневой папки
        self.folder_frame = QWidget()
        self.folder_layout = QHBoxLayout(self.folder_frame)

        self.folder_label = QLabel("Корневая папка:")
        self.folder_entry = QLineEdit()
        self.folder_entry.setPlaceholderText("Выберите папку...")
        self.browse_button = QPushButton("Обзор")
        self.browse_button.clicked.connect(self.browse_folder)

        self.folder_layout.addWidget(self.folder_label)
        self.folder_layout.addWidget(self.folder_entry)
        self.folder_layout.addWidget(self.browse_button)

        # Выбор шаблона СИ
        self.template_si_frame = QWidget()
        self.template_si_layout = QHBoxLayout(self.template_si_frame)

        self.template_si_label = QLabel("Шаблон СИ:")
        self.template_si_combo = QComboBox()
        self.template_si_combo.addItems(list(SI_TEMPLATES.keys()) + ["Пользовательский"])
        self.template_si_combo.setCurrentText("Пользовательский")
        self.template_si_combo.currentTextChanged.connect(self.apply_si_template)  # Автоматическое применение

        self.template_si_layout.addWidget(self.template_si_label)
        self.template_si_layout.addWidget(self.template_si_combo)

        # Выбор шаблона ИК
        self.template_ik_frame = QWidget()
        self.template_ik_layout = QHBoxLayout(self.template_ik_frame)

        self.template_ik_label = QLabel("Шаблон ИК:")
        self.template_ik_combo = QComboBox()
        self.template_ik_combo.addItems(list(IK_TEMPLATES.keys()) + ["Пользовательский"])
        self.template_ik_combo.setCurrentText("Общий шаблон")  # Установка по умолчанию
        self.template_ik_combo.currentTextChanged.connect(self.apply_ik_template)  # Автоматическое применение

        self.template_ik_layout.addWidget(self.template_ik_label)
        self.template_ik_layout.addWidget(self.template_ik_combo)

        # Параметры СИ и ИК
        self.params_frame = QWidget()
        self.params_layout = QHBoxLayout(self.params_frame)

        # СИ параметры
        self.si_params_frame = QWidget()
        self.si_params_layout = QVBoxLayout(self.si_params_frame)
        self.si_label = QLabel("СИ:")
        self.si_params_layout.addWidget(self.si_label)

        self.si_combos = []
        options = ["1", "2", "4", "Any"]

        for name in nameSI:
            frame = QWidget()
            layout = QHBoxLayout(frame)
            label = QLabel(f"{name}:")
            combo = QComboBox()
            combo.addItems(options)
            combo.setCurrentText("1")
            layout.addWidget(label)
            layout.addWidget(combo)
            self.si_params_layout.addWidget(frame)
            self.si_combos.append(combo)

        # ИК параметры
        self.ik_params_frame = QWidget()
        self.ik_params_layout = QVBoxLayout(self.ik_params_frame)
        self.ik_label = QLabel("ИК:")
        self.ik_params_layout.addWidget(self.ik_label)

        self.ik_combos = []

        for name in nameIK:
            frame = QWidget()
            layout = QHBoxLayout(frame)
            label = QLabel(f"{name}:")
            combo = QComboBox()
            combo.addItems(options)
            combo.setCurrentText("1")
            layout.addWidget(label)
            layout.addWidget(combo)
            self.ik_params_layout.addWidget(frame)
            self.ik_combos.append(combo)

        self.params_layout.addWidget(self.si_params_frame)
        self.params_layout.addWidget(self.ik_params_frame)

        # Прогресс и сообщения
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.message_label = QLabel()
        self.message_label.setAlignment(Qt.AlignCenter)

        # Кнопка запуска
        self.start_button = QPushButton("Начать обработку")
        self.start_button.clicked.connect(self.start_processing)

        # Добавляем все виджеты в основной layout
        self.layout.addWidget(self.folder_frame)
        self.layout.addWidget(self.template_si_frame)
        self.layout.addWidget(self.template_ik_frame)
        self.layout.addWidget(self.params_frame)
        self.layout.addWidget(self.progress_bar)
        self.layout.addWidget(self.message_label)
        self.layout.addWidget(self.start_button)

    def browse_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Выберите корневую папку")
        if folder:
            self.folder_entry.setText(folder)

    def apply_si_template(self):
        template_name = self.template_si_combo.currentText()
        if template_name != "Пользовательский":
            template = SI_TEMPLATES.get(template_name)
            if template:
                for i, value in enumerate(template):
                    if i < len(self.si_combos):
                        self.si_combos[i].setCurrentText(value)

    def apply_ik_template(self):
        template_name = self.template_ik_combo.currentText()
        if template_name != "Пользовательский":
            template = IK_TEMPLATES.get(template_name)
            if template:
                for i, value in enumerate(template):
                    if i < len(self.ik_combos):
                        self.ik_combos[i].setCurrentText(value)

    def start_processing(self):
        folder_path = self.folder_entry.text()
        if not folder_path:
            QMessageBox.critical(self, "Ошибка", "Необходимо выбрать корневую папку.")
            return

        fileqnt = [combo.currentText() for combo in self.si_combos]
        fileqnt_ik = [combo.currentText() for combo in self.ik_combos]

        self.progress_bar.setValue(0)
        self.message_label.setText("Подготовка к обработке...")
        self.start_button.setEnabled(False)

        self.worker = FolderProcessor(folder_path, fileqnt, fileqnt_ik)
        self.worker.signals.progress.connect(self.update_progress)
        self.worker.signals.message.connect(self.update_message)
        self.worker.signals.finished.connect(self.processing_finished)
        self.worker.signals.error.connect(self.show_error)
        self.worker.start()

    def update_progress(self, value):
        self.progress_bar.setValue(value)

    def update_message(self, message):
        self.message_label.setText(message)

    def processing_finished(self, excel_path):
        self.start_button.setEnabled(True)
        QMessageBox.information(self, "Готово",
                                f"Обработка завершена!\nОтчет сохранен в:\n{excel_path}")

    def show_error(self, error_msg):
        self.start_button.setEnabled(True)
        QMessageBox.critical(self, "Ошибка", error_msg)


if __name__ == "__main__":
    app = QApplication([])
    window = MainWindow()
    window.show()
    app.exec()