from pathlib import Path
from PySide6.QtWidgets import (
    QDialog, QLineEdit, QComboBox, QFormLayout,
    QDialogButtonBox, QMessageBox
)

ARCHIVE_PATH = Path(r"\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!АРХИВ")

class CreateCertDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Создание сертификата")

        layout = QFormLayout()

        self.cert_number_input = QLineEdit()
        self.cert_name_input = QLineEdit()

        self.year_selector = QComboBox()
        self.year_selector.addItems(sorted(
            [folder.name for folder in ARCHIVE_PATH.iterdir() if folder.is_dir()]
        ))
        self.year_selector.currentTextChanged.connect(self.suggest_next_number)

        layout.addRow("Номер сертификата:", self.cert_number_input)
        layout.addRow("Доп. название:", self.cert_name_input)
        layout.addRow("Год:", self.year_selector)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)

        layout.addWidget(buttons)
        self.setLayout(layout)

        self.suggest_next_number(self.year_selector.currentText())

    def suggest_next_number(self, year: str):
        year_path = ARCHIVE_PATH / year
        numbers = []

        if year_path.exists():
            for folder in year_path.iterdir():
                if folder.is_dir():
                    name = folder.name
                    if name[:3].isdigit():
                        numbers.append(int(name[:3]))

        next_number = max(numbers, default=0) + 1
        self.cert_number_input.setText(f"{next_number:03}")

    def get_data(self):
        return (
            self.cert_number_input.text().strip(),
            self.cert_name_input.text().strip(),
            self.year_selector.currentText()
        )

def create_certificate_structure(new_cert_path):
    """Создает структуру папок для нового сертификата"""
    si_folders = ['0 Заявка и приложение', '1 Распоряжение по заявке', '2 Решение по заявке',
                  '3 Заключения по ОМД и ТД',
                  '4 Акт выбора ПК', '5 Протоколы СИ', '6 Заключение СИ', '7 Программа проверки произ',
                  '8 Акт ПП', '9 Распоряжение на анализ', '10 Решение о выдаче', '11 Сертификат',
                  '12 Доп.материалы']

    ik_folders = ['0 Распоряжение', '1 Письмо-уведомление', '2 Программа ИК', '3 Программа проверки произ',
                  '4 Акт выбора ПК', '5 Акт проверки производства', '6 Протоколы ИК', '7 Акт по результатам ИК',
                  '8 Распоряжение на анализ', '9 Решение по ИК', '10 Доп. материалы']

    structure = {
        "0. СИ": si_folders,
        "1. ИК-1": ik_folders,
        "2. ИК-2": ik_folders,
    }

    for main_folder, subfolders in structure.items():
        main_path = new_cert_path / main_folder
        main_path.mkdir(parents=True, exist_ok=True)
        for subfolder in subfolders:
            subfolder_path = main_path / subfolder
            subfolder_path.mkdir(parents=True, exist_ok=True)

            # Особая обработка для "3 Заключения по ОМД и ТД"
            if subfolder == '3 Заключения по ОМД и ТД':
                for inner_folder in [
                    '3.1 Заключение ОМД',
                    '3.2 Заключение ТД и РПН',
                    '3.3 Заключение ПМ'
                ]:
                    (subfolder_path / inner_folder).mkdir(parents=True, exist_ok=True)