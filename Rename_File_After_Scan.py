# file_renamer_dnd.py
import sys
import os
import re
from typing import List, Dict
from PySide6.QtWidgets import (
    QApplication, QWidget, QPushButton, QVBoxLayout, QFileDialog,
    QLabel, QListWidget, QMessageBox, QHBoxLayout, QFrame, QAbstractItemView,
    QCheckBox, QGroupBox, QLineEdit, QComboBox, QScrollArea
)
from PySide6.QtCore import Qt

VALID_EXTENSIONS = {'.pdf', '.docx'}

# База данных документов для каждого типа (с номерами и форматами как в примере)
DOCUMENT_TYPES: Dict[str, List[str]] = {
    "СИ": [
        "01. ОС-{номер} - Распоряжение по заявке",
        "02. ОС-{номер} - Решение по заявке",
        "02.1 ОС-{номер} - Документы для процедуры СИ",
        "02.2 ОС-{номер} - Вопросник по АСП СИ",
        "03. ОС-{номер} - Заключение ОМТД",
        "04. ОС-{номер} - Акт выбора ПК",
        "04.1 ОС-{номер} - Направление-заявка СИ"
        "05. ОС-{номер} - Программа СИ",
        "06. ОС-{номер} - Заключение протоколы СИ",
        "07. ОС-{номер} - Программа АСП",
        "08. ОС-{номер} - Акт АСП",
        "09. ОС-{номер} - Распоряжение (анализ результатов)",
        "10. ОС-{номер} - Решение о выдаче сертификата",
        "11. СИ-{номер} - Сертификат №___",
        "12. CИ-{номер} - Прил. к сертификату №___"
    ],
    "ИК-1": [
        "01. ИК-{номер} - Распоряжение ИК №1",
        "02. ИК-{номер} - Извещение ИК №1",
        "03. ИК-{номер} - Программа ИК №1",
        "03.1 ИК-{номер} - Документы для процедуры ИК №1",
        "03.2 ИК-{номер} - Вопросник по АСП ИК №1",
        "04.1 ИК-{номер} - Акт по ОМТД ИК №1",
        "04.2 ИК-{номер} - Акт по МК ИК №1",
        "05. ИК-{номер} - Акт выбора ПК ИК №1",
        "05.2 ИК-{номер} - Направление-заявка ИК №1",
        "06. ИК-{номер} - Программа испытаний ИК №1",
        "07. ИК-{номер} - Акт по испытаниям ИК №1",
        "08. ИК-{номер} - Программа АСП ИК №1",
        "09. ИК-{номер} - Акт АСП ИК №1",
        "10. ИК-{номер} - Акт по ИК №1",
        "11.1 ИК-{номер} - Решение по ИК №1 (подтверждение)",
        "11.2 ИК-{номер} - Решение по ИК №1 (приостановка)",
        "11.3 ИК-{номер} - Решение по ИК №1 (прекращение)",
        "12. ИК-{номер} - Решение по ИК №1 (возобновление)",
        "13. ИК-{номер} - Решение по ИК №1 (отмена)"
    ],
    "ИК-2": [
        "01. ИК-{номер} - Распоряжение ИК №2",
        "02. ИК-{номер} - Извещение ИК №2",
        "03 ИК-{номер} - Программа ИК №2",
        "03.1 ИК-{номер} - Документы для процедуры ИК №2",
        "03.2 ИК-{номер} - Вопросник по АСП ИК №2",
        "04.1 ИК-{номер} - Акт по ОМТД ИК №2",
        "04.2 ИК-{номер} - Акт по МК ИК №2",
        "05. ИК-{номер} - Акт выбора ПК ИК №2",
        "05.2 ИК-{номер} - Направление-заявка ИК №2",
        "06. ИК-{номер} - Программа испытаний ИК №2",
        "07. ИК-{номер} - Акт по испытаниям ИК №2",
        "08. ИК-{номер} - Программа АСП ИК №2",
        "09. ИК-{номер} - Акт АСП ИК №2",
        "10. ИК-{номер} - Акт по ИК №2",
        "11.1 ИК-{номер} - Решение по ИК №2 (подтверждение)",
        "11.2 ИК-{номер} - Решение по ИК №2 (приостановка)",
        "11.3 ИК-{номер} - Решение по ИК №2 (прекращение)",
        "12. ИК-{номер} - Решение по ИК №2 (возобновление)",
        "13. ИК-{номер} - Решение по ИК №2 (отмена)"
    ]
}


def unique_path(path: str) -> str:
    """Если путь существует — добавить суффикс ' (копия N)' перед расширением."""
    if not os.path.exists(path):
        return path
    base, ext = os.path.splitext(path)
    n = 1
    while True:
        candidate = f"{base} (копия {n}){ext}"
        if not os.path.exists(candidate):
            return candidate
        n += 1


class DropListWidget(QListWidget):
    """QListWidget с корректным обработчиком внешнего drag&drop."""

    def __init__(self, parent=None, role="dest"):
        super().__init__(parent)
        self.role = role  # "dest" или "source"
        self.setAcceptDrops(True)
        self.setDragEnabled(False)
        self.setDropIndicatorShown(True)
        self.setDragDropMode(QAbstractItemView.DropOnly)
        self.setSelectionMode(QAbstractItemView.ExtendedSelection)

        self.default_style = """
            QListWidget {
                border: 2px dashed #bdbdbd;
                border-radius: 6px;
                background-color: #fcfcfc;
            }
            QListWidget::item { padding: 4px; }
        """
        self.highlight_style = """
            QListWidget {
                border: 2px dashed #0078d7;
                border-radius: 6px;
                background-color: #f0fbff;
            }
        """
        self.setStyleSheet(self.default_style)

    # При перетаскивании — принять, если есть URLs (файлы/папки)
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self.setStyleSheet(self.highlight_style)
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragLeaveEvent(self, event):
        self.setStyleSheet(self.default_style)

    def dropEvent(self, event):
        self.setStyleSheet(self.default_style)
        if not event.mimeData().hasUrls():
            event.ignore()
            return

        urls = event.mimeData().urls()
        paths = [u.toLocalFile() for u in urls if u.isLocalFile()]

        folders = [p for p in paths if os.path.isdir(p)]
        files = [p for p in paths if os.path.isfile(p)]

        if folders:
            # Если бросили папку(и) — возьмём первую и установим как папку
            folder = folders[0]
            self.parent().set_folder(folder, self.role)
        elif files:
            # Добавляем только выбранные файлы как "индивидуальные файлы"
            self.parent().add_individual_files(files, self.role)

        event.acceptProposedAction()


class DocumentSelector(QWidget):
    """Виджет для выбора документов через чекбоксы"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.checkboxes = []
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        # Выбор типа документа
        type_layout = QHBoxLayout()
        type_layout.addWidget(QLabel("Тип дела:"))
        self.doc_type_combo = QComboBox()
        self.doc_type_combo.addItems(["СИ", "ИК-1", "ИК-2"])
        self.doc_type_combo.currentTextChanged.connect(self.update_document_list)
        type_layout.addWidget(self.doc_type_combo)
        type_layout.addStretch()
        layout.addLayout(type_layout)

        # Поле для ввода номера дела (только цифры)
        case_layout = QHBoxLayout()
        case_layout.addWidget(QLabel("Номер дела (только цифры):"))
        self.case_number_edit = QLineEdit()
        self.case_number_edit.setPlaceholderText("Например: 534")
        self.case_number_edit.textChanged.connect(self.update_document_preview)
        case_layout.addWidget(self.case_number_edit)
        layout.addLayout(case_layout)

        # Группа с чекбоксами документов
        self.doc_group = QGroupBox("Выберите документы для переименования:")
        doc_layout = QVBoxLayout()

        # Scroll area для чекбоксов
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFixedHeight(600)

        self.checkbox_widget = QWidget()
        self.checkbox_layout = QVBoxLayout(self.checkbox_widget)

        scroll.setWidget(self.checkbox_widget)
        doc_layout.addWidget(scroll)

        # Кнопки выбора всех/очистки
        btn_layout = QHBoxLayout()
        self.select_all_btn = QPushButton("Выбрать все")
        self.select_all_btn.clicked.connect(self.select_all)
        self.clear_all_btn = QPushButton("Очистить все")
        self.clear_all_btn.clicked.connect(self.clear_all)
        btn_layout.addWidget(self.select_all_btn)
        btn_layout.addWidget(self.clear_all_btn)
        btn_layout.addStretch()
        doc_layout.addLayout(btn_layout)

        self.doc_group.setLayout(doc_layout)
        layout.addWidget(self.doc_group)

        self.setLayout(layout)

        # Загружаем первоначальный список документов
        self.update_document_list()

    def update_document_list(self):
        """Обновляет список чекбоксов в соответствии с выбранным типом"""
        # Очищаем старые чекбоксы
        for i in reversed(range(self.checkbox_layout.count())):
            widget = self.checkbox_layout.itemAt(i).widget()
            if widget:
                widget.deleteLater()

        self.checkboxes = []
        doc_type = self.doc_type_combo.currentText()
        documents = DOCUMENT_TYPES.get(doc_type, [])

        for i, doc_template in enumerate(documents, 1):
            # Показываем шаблон без номера дела
            checkbox = QCheckBox(doc_template.replace("{номер}", "XXX"))
            checkbox.template = doc_template  # сохраняем шаблон
            self.checkbox_layout.addWidget(checkbox)
            self.checkboxes.append(checkbox)

        # Обновляем предпросмотр с текущим номером дела
        self.update_document_preview()

    def update_document_preview(self):
        """Обновляет отображение названий документов с текущим номером дела"""
        case_number = self.case_number_edit.text().strip()
        # Оставляем только цифры
        clean_number = re.sub(r'\D', '', case_number)

        for checkbox in self.checkboxes:
            template = getattr(checkbox, 'template', '')
            if template:
                if clean_number:
                    display_name = template.replace("{номер}", clean_number)
                else:
                    display_name = template.replace("{номер}", "XXX")
                checkbox.setText(display_name)

    def select_all(self):
        """Выбирает все чекбоксы"""
        for checkbox in self.checkboxes:
            checkbox.setChecked(True)

    def clear_all(self):
        """Снимает выделение со всех чекбоксов"""
        for checkbox in self.checkboxes:
            checkbox.setChecked(False)

    def get_selected_documents(self):
        """Возвращает список выбранных документов с подставленным номером дела"""
        selected = []
        case_number = self.get_case_number()

        for checkbox in self.checkboxes:
            if checkbox.isChecked():
                template = getattr(checkbox, 'template', '')
                if template and case_number:
                    doc_name = template.replace("{номер}", case_number)
                    selected.append(doc_name)

        return selected

    def get_case_number(self):
        """Возвращает очищенный номер дела (только цифры)"""
        case_number = self.case_number_edit.text().strip()
        return re.sub(r'\D', '', case_number)


class FileRenamerApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Переименование файлов по шаблону")
        self.resize(950, 650)
        self.setAcceptDrops(True)

        # Папки (если пользователь выбрал папку)
        self.destination_folder = ""
        self.source_folder = ""

        # Списки "индивидуальных" файлов (полные пути)
        self.individual_dest_files: List[str] = []
        self.individual_source_files: List[str] = []

        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        # Верхняя подсказка
        top_label = QLabel(
            "Перетащи папку или файлы в левый список. Файлы должны быть отсортированы поо порядку их имен слева.\n"
            "В правой части выберите тип документов, введите номер дела и отметьте нужные документы."
        )
        top_label.setWordWrap(True)
        layout.addWidget(top_label)

        # Кнопки выбора папок
        btn_layout = QHBoxLayout()
        self.btn_select_dest = QPushButton("Выбрать папку с файлами для переименования")
        self.btn_select_dest.clicked.connect(lambda: self.select_folder("dest"))
        self.btn_select_dest.setToolTip("Выбрать папку с файлами, которые нужно переименовать")
        btn_layout.addWidget(self.btn_select_dest)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)

        # Статус строка с текущими путями
        self.status_label = QLabel("Папка с файлами: не задана")
        layout.addWidget(self.status_label)

        # Основная область с списком файлов и выбором документов
        main_layout = QHBoxLayout()

        # Левый список - файлы для переименования
        left_frame = QVBoxLayout()
        self.label_dest = QLabel("Файлы для переименования (должны быть отсортированы):")

        # Кнопки управления файлами
        file_buttons_layout = QHBoxLayout()
        self.btn_clear_all = QPushButton("Очистить все файлы")
        self.btn_clear_all.clicked.connect(self.clear_all_files)
        self.btn_remove_selected = QPushButton("Удалить выбранные")
        self.btn_remove_selected.clicked.connect(self.remove_selected_files)

        file_buttons_layout.addWidget(self.btn_clear_all)
        file_buttons_layout.addWidget(self.btn_remove_selected)
        file_buttons_layout.addStretch()

        left_frame.addWidget(self.label_dest)
        left_frame.addLayout(file_buttons_layout)

        self.list_dest = DropListWidget(self, role="dest")
        left_frame.addWidget(self.list_dest)

        # Правая часть - выбор документов
        right_frame = QVBoxLayout()
        self.document_selector = DocumentSelector()
        right_frame.addWidget(self.document_selector)

        main_layout.addLayout(left_frame)
        main_layout.addLayout(right_frame)
        layout.addLayout(main_layout)

        # Кнопки действий
        action_layout = QHBoxLayout()
        self.btn_preview = QPushButton("Показать соответствие (превью)")
        self.btn_preview.clicked.connect(self.show_preview)
        self.btn_rename = QPushButton("Переименовать")
        self.btn_rename.clicked.connect(self.rename_files)
        action_layout.addWidget(self.btn_preview)
        action_layout.addWidget(self.btn_rename)
        action_layout.addStretch()
        layout.addLayout(action_layout)

        # Предупреждение
        warning_frame = QFrame()
        warning_frame.setFrameShape(QFrame.Box)
        warning_frame.setStyleSheet("QFrame { border: 2px solid darkred; background:#fff6f6; }")
        w_layout = QVBoxLayout()
        w_label = QLabel(
            "⚠️ ВАЖНО: Файлы должны быть отсортированы по порядку их имен слева. Проверьте превью перед переименованием!")
        w_label.setWordWrap(True)
        w_layout.addWidget(w_label)
        warning_frame.setLayout(w_layout)
        layout.addWidget(warning_frame)

        self.setLayout(layout)

    def clear_all_files(self):
        """Очищает все файлы из списка"""
        if self.list_dest.count() == 0:
            return

        reply = QMessageBox.question(
            self,
            "Очистка файлов",
            "Вы уверены, что хотите удалить все файлы из списка?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            self.destination_folder = ""
            self.individual_dest_files = []
            self.update_file_list()
            self.update_status()

    def remove_selected_files(self):
        """Удаляет выбранные файлы из списка"""
        selected_items = self.list_dest.selectedItems()
        if not selected_items:
            QMessageBox.information(self, "Нет выбора", "Выберите файлы для удаления из списка.")
            return

        if self.individual_dest_files:
            # Удаляем выбранные файлы из списка индивидуальных файлов
            selected_names = {item.text() for item in selected_items}
            self.individual_dest_files = [
                f for f in self.individual_dest_files
                if os.path.basename(f) not in selected_names
            ]
        elif self.destination_folder:
            # При работе с папкой просто очищаем выделение
            # (не удаляем физические файлы, только из интерфейса)
            for item in selected_items:
                self.list_dest.takeItem(self.list_dest.row(item))

        self.update_file_list()
        self.update_status()

    def select_folder(self, role):
        folder = QFileDialog.getExistingDirectory(self, "Выберите папку")
        if folder:
            self.set_folder(folder, role)

    def set_folder(self, folder: str, role: str):
        if role == "dest":
            self.destination_folder = folder
            self.individual_dest_files = []
        else:
            self.source_folder = folder
            self.individual_source_files = []

        self.update_file_list()
        self.update_status()

    def add_individual_files(self, paths: List[str], role: str):
        clean = [p for p in paths if os.path.isfile(p) and os.path.splitext(p)[1].lower() in VALID_EXTENSIONS]
        if not clean:
            QMessageBox.warning(self, "Файлы не добавлены", "Нет допустимых файлов для добавления (pdf/docx).")
            return
        if role == "dest":
            self.destination_folder = ""
            # Добавляем только выбранные файлы
            self.individual_dest_files = clean
        else:
            self.source_folder = ""
            self.individual_source_files = clean

        self.update_file_list()
        self.update_status()

    def update_file_list(self):
        """Обновление списка файлов с сортировкой по номерам"""
        self.list_dest.clear()
        if self.individual_dest_files:
            # Сортируем файлы по номеру в начале имени
            sorted_files = self.sort_files_by_number(self.individual_dest_files)
            for p in sorted_files:
                self.list_dest.addItem(os.path.basename(p))
        elif self.destination_folder:
            items = self.get_filtered_files(self.destination_folder)
            self.list_dest.addItems(items)

    def sort_files_by_number(self, file_paths: List[str]) -> List[str]:
        """Сортирует файлы по номеру в начале имени"""

        def extract_number(filename):
            # Ищем число в начале имени файла
            match = re.match(r'^(\d+)', os.path.basename(filename))
            return int(match.group(1)) if match else 9999

        return sorted(file_paths, key=extract_number)

    def update_status(self):
        if self.destination_folder:
            status = f"Папка с файлами: {self.destination_folder} ({self.list_dest.count()} шт.)"
        elif self.individual_dest_files:
            status = f"Файлы: индивидуально ({len(self.individual_dest_files)} шт.)"
        else:
            status = "Файлы: не заданы"

        self.status_label.setText(status)

    def get_filtered_files(self, folder: str) -> List[str]:
        """Получает файлы и сортирует их по номерам"""
        files = [
            f for f in os.listdir(folder)
            if os.path.isfile(os.path.join(folder, f)) and os.path.splitext(f)[1].lower() in VALID_EXTENSIONS
        ]
        # Сортируем по номеру в начале имени
        return self.sort_files_by_number([os.path.join(folder, f) for f in files])

    def displayed_full_paths(self, list_widget: QListWidget, folder: str, individual_files: List[str]) -> List[str]:
        """Получает полные пути к файлам в порядке отображения"""
        displayed = [list_widget.item(i).text() for i in range(list_widget.count())]
        if individual_files:
            fulls = []
            used = set()
            for name in displayed:
                found = None
                for p in individual_files:
                    if os.path.basename(p) == name and p not in used:
                        found = p
                        break
                if found:
                    fulls.append(found)
                    used.add(found)
                else:
                    fulls.append(os.path.join(folder, name) if folder else name)
            return fulls
        else:
            return [os.path.join(folder, name) for name in displayed]

    def show_preview(self):
        """Показывает превью переименования"""
        if self.list_dest.count() == 0:
            QMessageBox.warning(self, "Пустой список", "Список файлов для переименования пуст.")
            return

        case_number = self.document_selector.get_case_number()
        if not case_number:
            QMessageBox.warning(self, "Не указан номер дела", "Введите номер дела (только цифры) для продолжения.")
            return

        selected_docs = self.document_selector.get_selected_documents()
        if not selected_docs:
            QMessageBox.warning(self, "Не выбраны документы", "Выберите хотя бы один документ для переименования.")
            return

        dest_fulls = self.displayed_full_paths(self.list_dest, self.destination_folder, self.individual_dest_files)

        if len(dest_fulls) != len(selected_docs):
            QMessageBox.warning(self, "Несоответствие количества",
                                f"Количество файлов ({len(dest_fulls)}) не соответствует количеству выбранных документов ({len(selected_docs)}).")
            return

        lines = []
        limit = min(20, len(dest_fulls))
        for i in range(limit):
            old_name = os.path.basename(dest_fulls[i])
            new_name = f"{selected_docs[i]}{os.path.splitext(dest_fulls[i])[1]}"
            lines.append(f"{old_name}\n  → {new_name}")

        if len(dest_fulls) > limit:
            lines.append(f"... и ещё {len(dest_fulls) - limit} файлов")

        QMessageBox.information(self, "Превью переименований", "\n\n".join(lines))

    def rename_files(self):
        """Основная операция переименования"""
        if self.list_dest.count() == 0:
            QMessageBox.warning(self, "Пустой список", "Список файлов для переименования пуст.")
            return

        case_number = self.document_selector.get_case_number()
        if not case_number:
            QMessageBox.warning(self, "Не указан номер дела", "Введите номер дела (только цифры) для продолжения.")
            return

        selected_docs = self.document_selector.get_selected_documents()
        if not selected_docs:
            QMessageBox.warning(self, "Не выбраны документы", "Выберите хотя бы один документ для переименования.")
            return

        dest_fulls = self.displayed_full_paths(self.list_dest, self.destination_folder, self.individual_dest_files)

        if len(dest_fulls) != len(selected_docs):
            QMessageBox.critical(self, "Ошибка",
                                 f"Количество файлов ({len(dest_fulls)}) не соответствует количеству выбранных документов ({len(selected_docs)}).")
            return

        # Собираем пары (old_full, new_full)
        pairs = []
        for i, old_full in enumerate(dest_fulls):
            old_ext = os.path.splitext(old_full)[1]
            new_name = f"{selected_docs[i]}{old_ext}"
            new_full = os.path.join(os.path.dirname(old_full), new_name)
            pairs.append((old_full, new_full))

        # Предварительный просмотр с подтверждением
        preview_lines = [f"{os.path.basename(a)}\n  → {os.path.basename(b)}" for a, b in pairs[:10]]
        if len(pairs) > 10:
            preview_lines.append(f"... и ещё {len(pairs) - 10} файлов")
        preview_text = "\n\n".join(preview_lines)

        reply = QMessageBox.question(self, "Подтвердите переименование",
                                     f"Будет переименовано {len(pairs)} файлов. Превью:\n\n{preview_text}\n\nПродолжить?",
                                     QMessageBox.Yes | QMessageBox.No,
                                     QMessageBox.No)
        if reply != QMessageBox.Yes:
            return

        # Выполняем переименование
        errors = []
        success_count = 0

        for old_full, new_full in pairs:
            try:
                if not os.path.exists(old_full):
                    errors.append(f"Не найден исходный: {old_full}")
                    continue

                final_new = unique_path(new_full)
                os.rename(old_full, final_new)
                success_count += 1
                QApplication.processEvents()

            except Exception as e:
                errors.append(f"Ошибка при переименовании {os.path.basename(old_full)}: {str(e)}")

        # Показываем результат
        if errors:
            error_text = "\n".join(errors[:10])
            if len(errors) > 10:
                error_text += f"\n... и ещё {len(errors) - 10} ошибок"
            QMessageBox.critical(self, "Частично завершено",
                                 f"Успешно переименовано: {success_count} из {len(pairs)} файлов\n\nОшибки:\n{error_text}")
        else:
            QMessageBox.information(self, "Готово", f"Все {success_count} файлов успешно переименованы.")

        # Обновляем интерфейс после операции
        self.individual_dest_files = []
        self.update_file_list()
        self.update_status()

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            new_files = []
            for url in event.mimeData().urls():
                path = url.toLocalFile()
                if os.path.isfile(path):
                    ext = os.path.splitext(path)[1].lower()
                    if ext in VALID_EXTENSIONS:
                        new_files.append(path)
                elif os.path.isdir(path):
                    # Если перетащили папку - устанавливаем ее как папку назначения
                    self.set_folder(path, "dest")
                    event.acceptProposedAction()
                    return

            if new_files:
                self.add_individual_files(new_files, "dest")
                event.acceptProposedAction()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = FileRenamerApp()
    win.show()
    sys.exit(app.exec())