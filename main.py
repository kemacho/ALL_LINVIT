
import sys
from pathlib import Path
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QWidget, QComboBox,
    QListWidget, QTabWidget, QTreeView, QFileSystemModel, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QDialog, QFormLayout, QDialogButtonBox, QMessageBox, QCheckBox
)
from PySide6.QtCore import Qt, QUrl, QThread, Signal
from PySide6.QtGui import QDesktopServices, QKeyEvent, QFont

ARCHIVE_PATH = Path(r"\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!АРХИВ")  # Путь к папке архива

class CustomTreeView(QTreeView):
    """ Свой QTreeView для поддержки горячих клавиш """
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent_viewer = parent

    def keyPressEvent(self, event: QKeyEvent):
        if event.key() == Qt.Key_Return or event.key() == Qt.Key_Enter:
            index = self.currentIndex()
            self.parent_viewer.open_file(index)
        else:
            super().keyPressEvent(event)

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

class ProjectLoaderThread(QThread):
    projects_loaded = Signal(list)

    def __init__(self, selected_year, parent=None):
        super().__init__(parent)
        self.selected_year = selected_year

    def run(self):
        projects = []
        if self.selected_year == "Все годы":
            for year_folder in sorted(ARCHIVE_PATH.iterdir()):
                if year_folder.is_dir():
                    for project in sorted(year_folder.iterdir()):
                        if project.is_dir():
                            projects.append((year_folder.name, project.name, project))
        else:
            year_path = ARCHIVE_PATH / self.selected_year
            if year_path.exists():
                for project in sorted(year_path.iterdir()):
                    if project.is_dir():
                        projects.append((self.selected_year, project.name, project))

        self.projects_loaded.emit(projects)

class ArchiveViewer(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Просмотр Архива")
        self.resize(1400, 800)

        # Основные виджеты
        self.year_selector = QComboBox()
        self.year_selector.addItem("Все годы")
        self.year_selector.addItems(sorted([folder.name for folder in ARCHIVE_PATH.iterdir() if folder.is_dir()]))
        self.year_selector.currentTextChanged.connect(self.update_projects)

        self.search_bar = QLineEdit()
        self.search_bar.setPlaceholderText("Поиск проекта...")
        self.search_bar.textChanged.connect(self.filter_projects)

        self.create_cert_button = QPushButton("Создать сертификат")
        self.create_cert_button.clicked.connect(self.create_certificate)

        self.expand_folders_checkbox = QCheckBox("Раскрывать папки по умолчанию")
        self.expand_folders_checkbox.setChecked(False)  # По умолчанию выключено

        self.project_list = QListWidget()
        self.project_list.itemClicked.connect(self.update_tabs)

        self.tabs = QTabWidget()

        self.project_title = QLabel("")
        self.project_title.setAlignment(Qt.AlignCenter)
        self.project_title.setStyleSheet("font-size: 18px; font-weight: bold; padding: 5px;")

        # Разметка
        left_layout = QVBoxLayout()
        left_layout.addWidget(QLabel("Выберите год:"))
        left_layout.addWidget(self.year_selector)
        left_layout.addWidget(QLabel("Поиск проекта:"))
        left_layout.addWidget(self.search_bar)

        left_layout.addWidget(self.expand_folders_checkbox)

        left_layout.addWidget(self.create_cert_button)
        left_layout.addWidget(QLabel("Проекты:"))
        left_layout.addWidget(self.project_list)

        left_widget = QWidget()
        left_widget.setLayout(left_layout)

        right_layout = QVBoxLayout()
        right_layout.addWidget(self.project_title)
        right_layout.addWidget(self.tabs)

        right_widget = QWidget()
        right_widget.setLayout(right_layout)

        main_layout = QHBoxLayout()
        main_layout.addWidget(left_widget, 2)
        main_layout.addWidget(right_widget, 5)

        container = QWidget()

        container.setLayout(main_layout)
        self.setCentralWidget(container)

        # Инициализация
        self.all_projects = []  # [(год, имя проекта, путь)]
        self.current_project_path = None
        self.update_projects("Все годы")

    def select_created_project(self, projects, folder_name):
        """Автоматически выбрать только что созданный проект"""
        for year, name, path in projects:
            if name == folder_name:
                self.all_projects = projects
                self.display_projects(projects)
                items = self.project_list.findItems(name, Qt.MatchExactly)
                if items:
                    item = items[0]
                    self.project_list.setCurrentItem(item)
                    self.update_tabs(item)
                break

    def on_projects_loaded(self, projects):
        self.all_projects = projects
        self.display_projects(projects)

    def update_projects(self, selected_year):
        """ Запускаем поток загрузки проектов """
        self.project_list.clear()
        self.all_projects.clear()
        self.project_list.addItem("Загрузка...")

        self.loader_thread = ProjectLoaderThread(selected_year)
        self.loader_thread.projects_loaded.connect(self.on_projects_loaded)
        self.loader_thread.start()

    def display_projects(self, projects):
        """ Показывает список проектов (только имена) """
        self.project_list.clear()
        for _, project_name, _ in projects:
            self.project_list.addItem(project_name)

    def filter_projects(self, text):
        """ Фильтрация списка проектов """
        filtered = [(year, name, path) for (year, name, path) in self.all_projects if text.lower() in name.lower()]
        self.display_projects(filtered)

    def update_tabs(self, item):
        """ Обновляем вкладки при выборе проекта """
        # Сохраняем текущую выбранную вкладку
        current_tab_index = self.tabs.currentIndex() if self.tabs.count() > 0 else -1

        self.tabs.clear()

        # Снимаем выделение со всех
        for i in range(self.project_list.count()):
            list_item = self.project_list.item(i)
            list_item.setFont(QFont())

        # Выделяем жирным выбранный проект
        item.setFont(QFont("", weight=QFont.Bold))

        selected_project_name = item.text()

        # Ищем путь к выбранному проекту
        for year, name, path in self.all_projects:
            if name == selected_project_name:
                self.current_project_path = path
                self.project_title.setText(f"{year} - {name}")
                break

        if self.current_project_path is None:
            return

        # Проверяем наличие каждой папки
        folders = {
            "СИ": self.current_project_path / "0. СИ",
            "ИК-1": self.current_project_path / "1. ИК-1",
            "ИК-2": self.current_project_path / "2. ИК-2",
        }

        for name, path in folders.items():
            if path.exists() and path.is_dir():
                view = CustomTreeView(self)
                model = QFileSystemModel()
                model.setRootPath(str(path))
                view.setModel(model)
                view.setRootIndex(model.index(str(path)))
                view.setColumnWidth(0, 400)

                view.doubleClicked.connect(self.open_file)

                # 📌 Раскрыть все папки, если чекбокс активен
                if self.expand_folders_checkbox.isChecked():
                    view.expandAll()

                self.tabs.addTab(view, name)

        # Восстанавливаем выбранную вкладку, если она существует в новом наборе
        if 0 <= current_tab_index < self.tabs.count():
            self.tabs.setCurrentIndex(current_tab_index)

    def open_file(self, index):
        """ Открыть файл по двойному клику или Enter """
        view = self.tabs.currentWidget()
        if isinstance(view, QTreeView):
            model = view.model()
            file_path = model.filePath(index)
            if Path(file_path).is_file():
                QDesktopServices.openUrl(QUrl.fromLocalFile(file_path))

    def create_certificate(self):
        dialog = CreateCertDialog(self)
        if dialog.exec() == QDialog.Accepted:
            cert_number, cert_name, selected_year = dialog.get_data()

            if not cert_number:
                QMessageBox.warning(self, "Ошибка", "Введите номер сертификата.")
                return

            folder_name = cert_number
            if cert_name:
                folder_name += f" - {cert_name}"

            year_path = ARCHIVE_PATH / selected_year
            new_cert_path = year_path / folder_name

            # Проверка на дубликат по первым 3 цифрам
            existing_numbers = {p.name[:3] for p in year_path.iterdir() if p.is_dir()}
            if cert_number in existing_numbers:
                QMessageBox.warning(self, "Ошибка", f"Сертификат с номером {cert_number} уже существует.")
                return

            # Структура папок
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

                    if subfolder == '3 Заключения по ОМД и ТД':
                        for inner_folder in [
                            '3.1 Заключение ОМД',
                            '3.2 Заключение ТД и РПН',
                            '3.3 Заключение ПМ'
                        ]:
                            (subfolder_path / inner_folder).mkdir(parents=True, exist_ok=True)

            QMessageBox.information(self, "Готово", f"Сертификат '{folder_name}' успешно создан.")

            # Обновим список и выберем только что созданный сертификат
            self.loader_thread = ProjectLoaderThread(selected_year)
            self.loader_thread.projects_loaded.connect(
                lambda projects: self.select_created_project(projects, folder_name))
            self.loader_thread.start()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ArchiveViewer()
    window.show()
    sys.exit(app.exec())
