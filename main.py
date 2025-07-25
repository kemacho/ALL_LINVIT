import sys
from pathlib import Path
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QWidget, QComboBox,
    QListWidget, QTabWidget, QTreeView, QFileSystemModel, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QDialog, QFormLayout, QDialogButtonBox,
    QMessageBox, QCheckBox
)
from PySide6.QtCore import Qt, QUrl, QThread, Signal, QTimer
from PySide6.QtGui import QDesktopServices, QKeyEvent, QFont, QMouseEvent

ARCHIVE_PATH = Path(r"\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!АРХИВ")


class CustomTreeView(QTreeView):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent_viewer = parent
        self.setUniformRowHeights(True)
        self.setAnimated(False)
        self.setExpandsOnDoubleClick(True)  # Включаем стандартное поведение для папок

    def keyPressEvent(self, event: QKeyEvent):
        if event.key() in (Qt.Key_Return, Qt.Key_Enter):
            index = self.currentIndex()
            model = self.model()
            if model.isDir(index):
                self.setExpanded(index, not self.isExpanded(index))
            else:
                self.parent_viewer.open_file(index)
        else:
            super().keyPressEvent(event)

    def mouseDoubleClickEvent(self, event: QMouseEvent):
        index = self.indexAt(event.pos())
        if not index.isValid():
            return

        model = self.model()
        if not model.isDir(index):
            # Только для файлов вызываем нашу обработку
            self.parent_viewer.open_file(index)
        else:
            # Для папок - стандартное поведение
            super().mouseDoubleClickEvent(event)


class CreateCertDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Создание сертификата")
        self.setup_ui()

    def setup_ui(self):
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
        numbers = [int(folder.name[:3]) for folder in year_path.iterdir()
                   if folder.is_dir() and folder.name[:3].isdigit()]
        next_num = max(numbers, default=0) + 1
        self.cert_number_input.setText(f"{next_num:03}")

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
                    projects.extend(
                        (year_folder.name, p.name, p)
                        for p in sorted(year_folder.iterdir())
                        if p.is_dir()
                    )
        else:
            year_path = ARCHIVE_PATH / self.selected_year
            if year_path.exists():
                projects.extend(
                    (self.selected_year, p.name, p)
                    for p in sorted(year_path.iterdir())
                    if p.is_dir()
                )
        self.projects_loaded.emit(projects)


class ArchiveViewer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setup_ui()
        self.all_projects = []
        self.current_project_path = None
        self.update_projects("Все годы")

    def setup_ui(self):
        self.setWindowTitle("Просмотр Архива")
        self.resize(1400, 800)

        # Виджеты
        self.year_selector = QComboBox()
        self.year_selector.addItem("Все годы")
        self.year_selector.addItems(sorted(
            [f.name for f in ARCHIVE_PATH.iterdir() if f.is_dir()]
        ))
        self.year_selector.currentTextChanged.connect(self.update_projects)

        self.search_bar = QLineEdit()
        self.search_bar.setPlaceholderText("Поиск проекта...")
        self.search_bar.textChanged.connect(self.filter_projects)

        self.create_cert_button = QPushButton("Создать сертификат")
        self.create_cert_button.clicked.connect(self.create_certificate)

        self.expand_folders_checkbox = QCheckBox("Раскрывать папки по умолчанию")
        self.expand_folders_checkbox.setChecked(False)
        self.expand_folders_checkbox.stateChanged.connect(self.update_folder_expansion)

        self.project_list = QListWidget()
        self.project_list.itemClicked.connect(self.update_tabs)

        self.tabs = QTabWidget()
        self.project_title = QLabel()
        self.project_title.setAlignment(Qt.AlignCenter)
        self.project_title.setStyleSheet("""
            font-size: 18px; 
            font-weight: bold; 
            padding: 5px;
        """)

        # Layout
        left_layout = QVBoxLayout()
        left_layout.addWidget(QLabel("Выберите год:"))
        left_layout.addWidget(self.year_selector)
        left_layout.addWidget(QLabel("Поиск проекта:"))
        left_layout.addWidget(self.search_bar)
        left_layout.addWidget(self.expand_folders_checkbox)
        left_layout.addWidget(self.create_cert_button)
        left_layout.addWidget(QLabel("Проекты:"))
        left_layout.addWidget(self.project_list)

        right_layout = QVBoxLayout()
        right_layout.addWidget(self.project_title)
        right_layout.addWidget(self.tabs)

        main_layout = QHBoxLayout()
        main_layout.addWidget(QWidget(layout=left_layout), 2)
        main_layout.addWidget(QWidget(layout=right_layout), 5)

        container = QWidget()
        container.setLayout(main_layout)
        self.setCentralWidget(container)

    def update_folder_expansion(self):
        """Обновляет раскрытие папок с небольшой задержкой"""
        state = self.expand_folders_checkbox.isChecked()
        for i in range(self.tabs.count()):
            view = self.tabs.widget(i)
            if isinstance(view, CustomTreeView):
                QTimer.singleShot(50, lambda v=view, s=state:
                v.expandAll() if s else v.collapseAll())

    def select_created_project(self, projects, folder_name):
        for year, name, path in projects:
            if name == folder_name:
                self.all_projects = projects
                self.display_projects(projects)
                if items := self.project_list.findItems(name, Qt.MatchExactly):
                    self.project_list.setCurrentItem(items[0])
                    self.update_tabs(items[0])
                break

    def on_projects_loaded(self, projects):
        self.all_projects = projects
        self.display_projects(projects)

    def update_projects(self, selected_year):
        self.project_list.clear()
        self.project_list.addItem("Загрузка...")

        self.loader_thread = ProjectLoaderThread(selected_year)
        self.loader_thread.projects_loaded.connect(self.on_projects_loaded)
        self.loader_thread.start()

    def display_projects(self, projects):
        self.project_list.clear()
        for _, name, _ in projects:
            self.project_list.addItem(name)

    def filter_projects(self, text):
        filtered = [
            (y, n, p) for y, n, p in self.all_projects
            if text.lower() in n.lower()
        ]
        self.display_projects(filtered)

    def update_tabs(self, item):
        current_idx = self.tabs.currentIndex() if self.tabs.count() else -1
        self.tabs.clear()

        for i in range(self.project_list.count()):
            self.project_list.item(i).setFont(QFont())

        item.setFont(QFont("", weight=QFont.Bold))
        selected_name = item.text()

        for year, name, path in self.all_projects:
            if name == selected_name:
                self.current_project_path = path
                self.project_title.setText(f"{year} - {name}")
                break

        if not self.current_project_path:
            return

        folders = {
            "СИ": self.current_project_path / "0. СИ",
            "ИК-1": self.current_project_path / "1. ИК-1",
            "ИК-2": self.current_project_path / "2. ИК-2",
        }

        for tab_name, folder_path in folders.items():
            if folder_path.exists() and folder_path.is_dir():
                view = CustomTreeView(self)
                model = QFileSystemModel()
                model.setRootPath(str(folder_path))
                view.setModel(model)
                view.setRootIndex(model.index(str(folder_path)))
                view.setColumnWidth(0, 400)
                view.setIndentation(15)
                view.setStyleSheet("QTreeView::item { height: 25px; }")
                view.doubleClicked.connect(lambda idx, v=view: self.open_file(idx))

                if self.expand_folders_checkbox.isChecked():
                    QTimer.singleShot(100, view.expandAll)

                self.tabs.addTab(view, tab_name)

        if 0 <= current_idx < self.tabs.count():
            self.tabs.setCurrentIndex(current_idx)

    def open_file(self, index):
        """Открывает только файлы, папки обрабатываются автоматически"""
        view = self.tabs.currentWidget()
        if isinstance(view, CustomTreeView):
            model = view.model()
            path = model.filePath(index)
            if Path(path).is_file():
                QDesktopServices.openUrl(QUrl.fromLocalFile(path))

    def create_certificate(self):
        dialog = CreateCertDialog(self)
        if dialog.exec() == QDialog.Accepted:
            num, name, year = dialog.get_data()
            if not num:
                QMessageBox.warning(self, "Ошибка", "Введите номер сертификата.")
                return

            folder_name = f"{num}{f' - {name}' if name else ''}"
            year_path = ARCHIVE_PATH / year
            new_path = year_path / folder_name

            if num in {p.name[:3] for p in year_path.iterdir() if p.is_dir()}:
                QMessageBox.warning(self, "Ошибка", f"Сертификат {num} уже существует.")
                return

            # Создаем структуру папок
            self.create_folder_structure(new_path)
            QMessageBox.information(self, "Готово", f"Создан сертификат: {folder_name}")

            self.loader_thread = ProjectLoaderThread(year)
            self.loader_thread.projects_loaded.connect(
                lambda ps: self.select_created_project(ps, folder_name))
            self.loader_thread.start()

    def create_folder_structure(self, base_path):
        structure = {
            "0. СИ": [
                '0 Заявка и приложение', '1 Распоряжение по заявке',
                '2 Решение по заявке', '3 Заключения по ОМД и ТД',
                '4 Акт выбора ПК', '5 Протоколы СИ', '6 Заключение СИ',
                '7 Программа проверки произ', '8 Акт ПП',
                '9 Распоряжение на анализ', '10 Решение о выдаче',
                '11 Сертификат', '12 Доп.материалы'
            ],
            "1. ИК-1": [
                '0 Распоряжение', '1 Письмо-уведомление', '2 Программа ИК',
                '3 Программа проверки произ', '4 Акт выбора ПК',
                '5 Акт проверки производства', '6 Протоколы ИК',
                '7 Акт по результатам ИК', '8 Распоряжение на анализ',
                '9 Решение по ИК', '10 Доп. материалы'
            ],
            "2. ИК-2": [
                '0 Распоряжение', '1 Письмо-уведомление', '2 Программа ИК',
                '3 Программа проверки произ', '4 Акт выбора ПК',
                '5 Акт проверки производства', '6 Протоколы ИК',
                '7 Акт по результатам ИК', '8 Распоряжение на анализ',
                '9 Решение по ИК', '10 Доп. материалы'
            ]
        }

        for main_folder, subfolders in structure.items():
            main_dir = base_path / main_folder
            main_dir.mkdir(parents=True, exist_ok=True)

            for sub in subfolders:
                sub_dir = main_dir / sub
                sub_dir.mkdir(exist_ok=True)

                if sub == '3 Заключения по ОМД и ТД':
                    for inner in ['3.1 Заключение ОМД', '3.2 Заключение ТД и РПН', '3.3 Заключение ПМ']:
                        (sub_dir / inner).mkdir(exist_ok=True)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ArchiveViewer()
    window.show()
    sys.exit(app.exec())