import sys
import os
from PySide6.QtWidgets import (
    QApplication, QWidget, QPushButton, QVBoxLayout,
    QFileDialog, QLabel, QListWidget, QMessageBox, QHBoxLayout, QFrame
)
from PySide6.QtCore import Qt

VALID_EXTENSIONS = ['.pdf', '.docx']

class FileRenamerApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Переименование файлов по шаблону")
        self.resize(800, 600)

        self.destination_folder = ""
        self.source_folder = ""

        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        # Кнопки выбора папок
        btn_layout = QHBoxLayout()
        self.btn_select_dest = QPushButton("Выбрать папку c файлами для переименования")
        self.btn_select_source = QPushButton("Выбрать папку с правильно названными файлами")
        self.btn_select_dest.clicked.connect(self.select_destination_folder)
        self.btn_select_source.clicked.connect(self.select_source_folder)
        btn_layout.addWidget(self.btn_select_dest)
        btn_layout.addWidget(self.btn_select_source)

        # Метки и списки файлов
        self.label_dest = QLabel("Файлы назначения:")
        self.list_dest = QListWidget()
        self.label_source = QLabel("Файлы источника:")
        self.list_source = QListWidget()

        # Кнопка переименования
        self.btn_rename = QPushButton("Переименовать")
        self.btn_rename.clicked.connect(self.rename_files)

        # Предупреждающее сообщение в рамке
        warning_frame = QFrame()
        warning_frame.setFrameShape(QFrame.Box)
        warning_frame.setStyleSheet(
            "QFrame { border: 2px solid darkred; border-radius: 6px; background-color: #fff0f0; }")

        warning_layout = QVBoxLayout()
        warning_label = QLabel(
            "⚠️ Обратите внимание!\n"
            "• Количество файлов в обеих папках должно совпадать.\n"
            "• Порядок важен — файлы будут переименованы по алфавиту (от А до Я).\n"
            "• Операция необратима. Если не уверены, создайте резервную копию."
        )
        warning_label.setWordWrap(True)
        warning_label.setStyleSheet("color: darkred; font-weight: bold;")

        warning_layout.addWidget(warning_label)
        warning_frame.setLayout(warning_layout)

        layout.addWidget(warning_frame)

        # Добавление в макет
        layout.addLayout(btn_layout)
        layout.addWidget(self.label_dest)
        layout.addWidget(self.list_dest)
        layout.addWidget(self.label_source)
        layout.addWidget(self.list_source)
        layout.addWidget(self.btn_rename)

        self.setLayout(layout)

    def select_destination_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Выберите папку назначения")
        if folder:
            self.destination_folder = folder
            self.update_file_list()

    def select_source_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Выберите папку источника")
        if folder:
            self.source_folder = folder
            self.update_file_list()

    def update_file_list(self):
        if self.destination_folder:
            dest_files = self.get_filtered_files(self.destination_folder)
            self.list_dest.clear()
            self.list_dest.addItems(dest_files)

        if self.source_folder:
            source_files = self.get_filtered_files(self.source_folder)
            self.list_source.clear()
            self.list_source.addItems(source_files)

    def get_filtered_files(self, folder):
        files = [
            f for f in os.listdir(folder)
            if os.path.isfile(os.path.join(folder, f)) and os.path.splitext(f)[1].lower() in VALID_EXTENSIONS
        ]
        return sorted(files, key=lambda x: x.lower())

    def rename_files(self):
        dest_files = self.get_filtered_files(self.destination_folder)
        source_files = self.get_filtered_files(self.source_folder)

        if len(dest_files) != len(source_files):
            QMessageBox.critical(self, "Ошибка", "Количество файлов в папках не совпадает.")
            return


        reply = QMessageBox.question(self, "Подтвердите переименование", "Вы уверены",
                                     QMessageBox.Yes | QMessageBox.No)

        if reply == QMessageBox.Yes:
            for dest, src in zip(dest_files, source_files):
                dest_path = os.path.join(self.destination_folder, dest)
                new_name = os.path.splitext(src)[0] + os.path.splitext(dest)[1]
                new_path = os.path.join(self.destination_folder, new_name)
                try:
                    os.rename(dest_path, new_path)
                except Exception as e:
                    QMessageBox.critical(self, "Ошибка", f"Не удалось переименовать {dest}: {e}")
                    return
            QMessageBox.information(self, "Успех", "Файлы успешно переименованы.")
            self.update_file_list()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = FileRenamerApp()
    window.show()
    sys.exit(app.exec())