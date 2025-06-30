
import os
import glob
from pathlib import Path
from PySide6.QtWidgets import (
    QWidget, QLabel, QPushButton, QComboBox, QVBoxLayout,
    QHBoxLayout, QGridLayout, QProgressBar, QMessageBox,
    QGroupBox, QSizePolicy
)
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication

ARCHIVE_PATH = Path(r"\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!АРХИВ")

NAME_SI = [
    '0 Заявка и приложение', '1 Распоряжение по заявке', '2 Решение по заявке',
    '3 Заключения по ОМД и ТД', '4 Акт выбора ПК', '5 Протоколы СИ', '6 Заключение СИ',
    '7 Программа проверки произ', '8 Акт ПП', '9 Распоряжение на анализ',
    '10 Решение о выдаче', '11 Сертификат', '12 Доп.материалы'
]

NAME_IK = [
    '0 Распоряжение', '1 Письмо-уведомление', '2 Программа ИК', '3 Программа проверки произ',
    '4 Акт выбора ПК', '5 Акт проверки производства', '6 Протоколы ИК',
    '7 Акт по результатам ИК', '8 Распоряжение на анализ', '9 Решение по ИК', '10 Доп. материалы'
]

SI_TEMPLATES = {
    "Шаблон РЖД": ['2', '1', '1', '4', '1', 'Any', '1', '1', '1', '1', '2', '2', 'Any'],
    "Общий шаблон": ['1', '1', '1', '4', '1', 'Any', '1', '1', '1', '1', '1', '1', 'Any']
}

IK_TEMPLATES = {
    "Шаблон РЖД": ['1', '1', '1', '1', '1', '1', 'Any', '1', '1', '1', 'Any'],
    "Общий шаблон": ['1', '1', '1', '1', '1', '1', 'Any', '1', '1', '1', 'Any']
}

class FolderProcessorApp(QWidget):
    def __init__(self):
        super().__init__()
        self.si_qnt = []
        self.ik_qnt = []
        self.selected_folder = ""
        self.setup_ui()

    def setup_ui(self):
        self.setWindowTitle("Обработчик папок")
        layout = QVBoxLayout(self)

        folder_group = QGroupBox("Выбор года")
        folder_layout = QHBoxLayout()
        self.year_selector = QComboBox()
        years = sorted([f.name for f in ARCHIVE_PATH.iterdir() if f.is_dir()])
        self.year_selector.addItems(years)
        self.year_selector.currentTextChanged.connect(self.update_selected_path)
        self.path_display = QLabel("")
        folder_layout.addWidget(QLabel("Год:"))
        folder_layout.addWidget(self.year_selector)
        folder_group.setLayout(folder_layout)
        layout.addWidget(folder_group)
        layout.addWidget(self.path_display)

        template_group = QGroupBox("Выбор шаблонов")
        template_layout = QHBoxLayout()
        self.template_si = QComboBox()
        self.template_ik = QComboBox()
        self.template_si.addItems(list(SI_TEMPLATES.keys()) + ["Пользовательский"])
        self.template_ik.addItems(list(IK_TEMPLATES.keys()) + ["Пользовательский"])
        self.template_si.currentTextChanged.connect(self.apply_si_template)
        self.template_ik.currentTextChanged.connect(self.apply_ik_template)
        template_layout.addWidget(QLabel("Шаблон СИ:"))
        template_layout.addWidget(self.template_si)
        template_layout.addSpacing(20)
        template_layout.addWidget(QLabel("Шаблон ИК:"))
        template_layout.addWidget(self.template_ik)
        template_group.setLayout(template_layout)
        layout.addWidget(template_group)

        grid_group = QGroupBox("Индивидуальные настройки")
        grid = QGridLayout()
        for i, name in enumerate(NAME_SI):
            grid.addWidget(QLabel(name), i, 0)
            cb = QComboBox()
            cb.addItems(["1", "2", "4", "Any"])
            self.si_qnt.append(cb)
            grid.addWidget(cb, i, 1)

        for i, name in enumerate(NAME_IK):
            grid.addWidget(QLabel(name), i, 2)
            cb = QComboBox()
            cb.addItems(["1", "2", "4", "Any"])
            self.ik_qnt.append(cb)
            grid.addWidget(cb, i, 3)
        grid_group.setLayout(grid)
        layout.addWidget(grid_group)

        self.progress_bar = QProgressBar()
        self.message_label = QLabel("")
        layout.addWidget(self.progress_bar)
        layout.addWidget(self.message_label)

        start_btn = QPushButton("Начать обработку")
        start_btn.clicked.connect(self.start_processing)
        layout.addWidget(start_btn)

        self.template_si.setCurrentText("Общий шаблон")
        self.template_ik.setCurrentText("Общий шаблон")
        self.update_selected_path(self.year_selector.currentText())

    def update_selected_path(self, year):
        selected_path = ARCHIVE_PATH / year
        self.selected_folder = str(selected_path)
        self.path_display.setText(f"Путь: {self.selected_folder}")

    def apply_si_template(self, template_name):
        if template_name in SI_TEMPLATES:
            for i, val in enumerate(SI_TEMPLATES[template_name]):
                if i < len(self.si_qnt):
                    index = self.si_qnt[i].findText(val)
                    if index >= 0:
                        self.si_qnt[i].setCurrentIndex(index)

    def apply_ik_template(self, template_name):
        if template_name in IK_TEMPLATES:
            for i, val in enumerate(IK_TEMPLATES[template_name]):
                if i < len(self.ik_qnt):
                    index = self.ik_qnt[i].findText(val)
                    if index >= 0:
                        self.ik_qnt[i].setCurrentIndex(index)

    def start_processing(self):
        inpath = self.selected_folder
        fileqnt = [cb.currentText() for cb in self.si_qnt]
        fileqnt_ik = [cb.currentText() for cb in self.ik_qnt]
        process_folders(inpath, fileqnt, fileqnt_ik, self.progress_bar, self.message_label)

def process_folders(inpath, fileqnt, fileqnt_ik, progress_bar, message_label):
    try:
        folder_names = os.listdir(inpath)
    except Exception as e:
        QMessageBox.critical(None, "Ошибка", str(e))
        return

    worry = []

    def worrymessage(pt, nm1, nm2):
        worry.append((pt, nm1, nm2))

    def check(FileQNT, old_folder, contents, Pos_dest, Norm_dest, Neg_dest):
        filtered_contents = [f for f in contents if f != "Thumbs.db"]
        if FileQNT == "1":
            if len(filtered_contents) == 1:
                os.rename(old_folder, Pos_dest)
            elif len(filtered_contents) == 0:
                os.rename(old_folder, Neg_dest)
            else:
                os.rename(old_folder, Pos_dest)
                worrymessage(old_folder, len(filtered_contents), '1')
        elif FileQNT == "2":
            if len(filtered_contents) == 2:
                os.rename(old_folder, Pos_dest)
            elif len(filtered_contents) == 1:
                os.rename(old_folder, Norm_dest)
            elif len(filtered_contents) == 0:
                os.rename(old_folder, Neg_dest)
            else:
                os.rename(old_folder, Pos_dest)
                worrymessage(old_folder, len(filtered_contents), '2')
        elif FileQNT == "4":
            if len(filtered_contents) == 4:
                os.rename(old_folder, Pos_dest)
            elif len(filtered_contents) < 4:
                os.rename(old_folder, Neg_dest)
            else:
                os.rename(old_folder, Pos_dest)
                worrymessage(old_folder, len(filtered_contents), '4')
        elif FileQNT == "Any":
            if len(filtered_contents) > 0:
                os.rename(old_folder, Pos_dest)
            else:
                os.rename(old_folder, Neg_dest)
            if len(filtered_contents) > 20:
                worrymessage(old_folder, len(filtered_contents), 'возможно не так много')

    nameSI = NAME_SI
    nameIK = NAME_IK
    total = len(folder_names)

    for i, name in enumerate(folder_names):
        pathSI = os.path.join(inpath, name, "0. СИ")
        pathIK1 = os.path.join(inpath, name, "1. ИК-1")
        pathIK2 = os.path.join(inpath, name, "2. ИК-2")

        for j, folder_pattern in enumerate(nameSI):
            try:
                old_folder = glob.glob(os.path.join(pathSI, folder_pattern + "*"))
                if old_folder:
                    old_folder = old_folder[0]
                    contents = os.listdir(old_folder)
                    base = os.path.join(pathSI, folder_pattern)
                    check(fileqnt[j], old_folder, contents, base + " (+)", base + " (+—)", base + " (—)")
            except Exception as e:
                print(f"Ошибка СИ {folder_pattern} в {name}: {e}")

        for j, folder_pattern in enumerate(nameIK):
            for pathIK in (pathIK1, pathIK2):
                try:
                    old_folder = glob.glob(os.path.join(pathIK, folder_pattern + "*"))
                    if old_folder:
                        old_folder = old_folder[0]
                        contents = os.listdir(old_folder)
                        base = os.path.join(pathIK, folder_pattern)
                        check(fileqnt_ik[j], old_folder, contents, base + " (+)", base + " (+—)", base + " (—)")
                except Exception as e:
                    print(f"Ошибка ИК {folder_pattern} в {name}: {e}")

        progress = int((i + 1) / total * 100)
        progress_bar.setValue(progress)
        message_label.setText(f"Обработано: {i + 1} из {total}")
        QApplication.processEvents()

    if worry:
        msg = "".join([f"{pt} — {nm1} файла(ов), ожидалось {nm2}" for pt, nm1, nm2 in worry])
        QMessageBox.warning(None, "Предупреждения", msg)
    else:
        QMessageBox.information(None, "Готово", "Проверка завершена.")
