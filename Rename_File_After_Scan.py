import sys
import os
import re
import json
from typing import List, Dict
from PySide6.QtWidgets import (
    QApplication, QWidget, QPushButton, QVBoxLayout, QFileDialog,
    QLabel, QListWidget, QMessageBox, QHBoxLayout, QFrame, QAbstractItemView,
    QCheckBox, QGroupBox, QLineEdit, QComboBox, QScrollArea, QDialog,
    QPlainTextEdit, QDialogButtonBox
)
from PySide6.QtCore import Qt

VALID_EXTENSIONS = {'.pdf', '.docx', '.doc', '.jpg', '.png'}

# --- Стандартные настройки (Заводские) ---
DEFAULT_RULES: Dict[str, List[str]] = {
    "СИ": [
        "01. ОС-{номер} - Распоряжение по заявке",
        "02. ОС-{номер} - Решение по заявке",
        "02.1 ОС-{номер} - Документы для процедуры СИ",
        "02.2 ОС-{номер} - Вопросник по АСП СИ",
        "03. ОС-{номер} - Заключение ОМТД",
        "04. ОС-{номер} - Акт выбора ПК",
        "04.1 ОС-{номер} - Направление-заявка СИ",
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
        "11.1 ИК-{номер} - Решение по ИК №1 (подтверждение)"
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
        "11.1 ИК-{номер} - Решение по ИК №2 (подтверждение)"
    ]
}


class RuleManager:
    """Класс для управления сохранением и загрузкой правил в JSON"""
    FILENAME = "rules.json"

    def __init__(self):
        self.rules = self.load_rules()

    def load_rules(self) -> Dict[str, List[str]]:
        if not os.path.exists(self.FILENAME):
            return DEFAULT_RULES.copy()
        try:
            with open(self.FILENAME, 'r', encoding='utf-8') as f:
                data = json.load(f)
                # Простая валидация
                if isinstance(data, dict) and all(isinstance(v, list) for v in data.values()):
                    return data
                return DEFAULT_RULES.copy()
        except Exception:
            return DEFAULT_RULES.copy()

    def save_rules(self):
        try:
            with open(self.FILENAME, 'w', encoding='utf-8') as f:
                json.dump(self.rules, f, ensure_ascii=False, indent=4)
        except Exception as e:
            print(f"Ошибка сохранения правил: {e}")

    def reset_to_defaults(self):
        self.rules = DEFAULT_RULES.copy()
        self.save_rules()

    def get_types(self):
        return list(self.rules.keys())

    def get_rules_for(self, doc_type):
        return self.rules.get(doc_type, [])

    def update_rules_for(self, doc_type, new_list):
        self.rules[doc_type] = new_list
        self.save_rules()


class RulesEditorDialog(QDialog):
    """Окно редактирования шаблонов"""

    def __init__(self, rule_manager: RuleManager, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Редактор шаблонов документов")
        self.resize(600, 500)
        self.rule_manager = rule_manager
        self.current_type = "СИ"

        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        # Выбор типа для редактирования
        top_layout = QHBoxLayout()
        top_layout.addWidget(QLabel("Выберите тип пакета:"))
        self.type_combo = QComboBox()
        self.type_combo.addItems(self.rule_manager.get_types())
        self.type_combo.currentTextChanged.connect(self.load_type_to_editor)
        top_layout.addWidget(self.type_combo)
        layout.addLayout(top_layout)

        layout.addWidget(
            QLabel("Список шаблонов (каждый с новой строки):\nИспользуйте {номер} для подстановки номера дела."))

        # Редактор текста
        self.editor = QPlainTextEdit()
        layout.addWidget(self.editor)

        # Кнопки
        btn_layout = QHBoxLayout()

        self.btn_reset = QPushButton("Сбросить все на заводские")
        self.btn_reset.setStyleSheet("color: red;")
        self.btn_reset.clicked.connect(self.reset_defaults)

        self.btn_save = QPushButton("Сохранить изменения")
        self.btn_save.clicked.connect(self.save_current_changes)

        self.btn_close = QPushButton("Закрыть")
        self.btn_close.clicked.connect(self.accept)

        btn_layout.addWidget(self.btn_reset)
        btn_layout.addStretch()
        btn_layout.addWidget(self.btn_save)
        btn_layout.addWidget(self.btn_close)

        layout.addLayout(btn_layout)
        self.setLayout(layout)

        # Загружаем начальное состояние
        self.load_type_to_editor(self.type_combo.currentText())

    def load_type_to_editor(self, type_name):
        """Загружает список строк в текстовое поле"""
        # Сначала сохраним текущее состояние (если переключились)
        # self.save_current_changes_internal() # Можно раскомментировать для автосохранения при переключении

        self.current_type = type_name
        rules = self.rule_manager.get_rules_for(type_name)
        text = "\n".join(rules)
        self.editor.setPlainText(text)

    def save_current_changes(self):
        """Сохраняет текст из редактора в менеджер"""
        text = self.editor.toPlainText()
        # Разбиваем на строки и убираем пустые
        lines = [line.strip() for line in text.split('\n') if line.strip()]

        self.rule_manager.update_rules_for(self.current_type, lines)
        QMessageBox.information(self, "Сохранено", f"Шаблоны для '{self.current_type}' обновлены!")

    def reset_defaults(self):
        reply = QMessageBox.question(self, "Сброс", "Вы уверены? Все ваши изменения шаблонов будут потеряны.",
                                     QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            self.rule_manager.reset_to_defaults()
            self.load_type_to_editor(self.current_type)
            QMessageBox.information(self, "Сброс", "Настройки возвращены к заводским.")


def unique_path(path: str) -> str:
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
    def __init__(self, parent=None, role="dest"):
        super().__init__(parent)
        self.role = role
        self.setAcceptDrops(True)
        self.setDragEnabled(True)  # Разрешаем драг
        self.setDropIndicatorShown(True)

        # Разрешаем внутреннее перемещение (Reordering)
        self.setDragDropMode(QAbstractItemView.InternalMove)
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

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self.setStyleSheet(self.highlight_style)
        else:
            # Для внутреннего перемещения
            super().dragEnterEvent(event)

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            super().dragMoveEvent(event)

    def dragLeaveEvent(self, event):
        self.setStyleSheet(self.default_style)
        super().dragLeaveEvent(event)

    def dropEvent(self, event):
        self.setStyleSheet(self.default_style)

        # Если это файлы извне
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            paths = [u.toLocalFile() for u in urls if u.isLocalFile()]

            folders = [p for p in paths if os.path.isdir(p)]
            files = [p for p in paths if os.path.isfile(p)]

            if folders:
                self.parent().set_folder(folders[0], self.role)
            elif files:
                self.parent().add_individual_files(files, self.role)

            event.acceptProposedAction()
        else:
            # Если это внутреннее перемещение строк
            super().dropEvent(event)
            # После перемещения нужно обновить внутренний список файлов в родителе
            self.parent().sync_list_order()


class DocumentSelector(QWidget):
    def __init__(self, rule_manager: RuleManager, parent=None):
        super().__init__(parent)
        self.rule_manager = rule_manager
        self.checkboxes = []
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        # Верхняя панель: Тип + Кнопка настроек
        top_bar = QHBoxLayout()

        type_layout = QHBoxLayout()
        type_layout.addWidget(QLabel("Тип дела:"))
        self.doc_type_combo = QComboBox()
        self.doc_type_combo.addItems(self.rule_manager.get_types())
        self.doc_type_combo.currentTextChanged.connect(self.update_document_list)
        type_layout.addWidget(self.doc_type_combo)

        top_bar.addLayout(type_layout)
        top_bar.addStretch()

        # Кнопка настроек
        self.btn_settings = QPushButton("⚙️ Редактор шаблонов")
        self.btn_settings.clicked.connect(self.open_settings)
        top_bar.addWidget(self.btn_settings)

        layout.addLayout(top_bar)

        # Поле для ввода номера дела
        case_layout = QHBoxLayout()
        case_layout.addWidget(QLabel("Номер дела (только цифры):"))
        self.case_number_edit = QLineEdit()
        self.case_number_edit.setPlaceholderText("Например: 534")
        self.case_number_edit.textChanged.connect(self.update_document_preview)
        case_layout.addWidget(self.case_number_edit)

        self.btn_decrease = QPushButton("−")
        self.btn_decrease.setFixedWidth(30)
        self.btn_decrease.clicked.connect(self.decrease_case_number)

        self.btn_increase = QPushButton("+")
        self.btn_increase.setFixedWidth(30)
        self.btn_increase.clicked.connect(self.increase_case_number)

        case_layout.addWidget(self.btn_decrease)
        case_layout.addWidget(self.btn_increase)
        case_layout.addStretch()
        layout.addLayout(case_layout)

        # Группа с чекбоксами
        self.doc_group = QGroupBox("Выберите документы:")
        doc_layout = QVBoxLayout()

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFixedHeight(500)

        self.checkbox_widget = QWidget()
        self.checkbox_layout = QVBoxLayout(self.checkbox_widget)
        # Прижимаем чекбоксы к верху
        self.checkbox_layout.setAlignment(Qt.AlignTop)

        scroll.setWidget(self.checkbox_widget)
        doc_layout.addWidget(scroll)

        # Кнопки выбора
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
        self.update_document_list()

    def open_settings(self):
        """Открывает окно редактирования шаблонов"""
        dialog = RulesEditorDialog(self.rule_manager, self)
        if dialog.exec():
            # Если нажали закрыть, обновляем список в главном окне
            self.update_document_list()

    def decrease_case_number(self):
        current = self.get_case_number()
        if current:
            val = int(current)
            if val > 1:
                self.case_number_edit.setText(str(val - 1))
        else:
            self.case_number_edit.setText("1")

    def increase_case_number(self):
        current = self.get_case_number()
        val = int(current) if current else 0
        self.case_number_edit.setText(str(val + 1))

    def update_document_list(self):
        # Очистка
        for i in reversed(range(self.checkbox_layout.count())):
            w = self.checkbox_layout.itemAt(i).widget()
            if w: w.deleteLater()
        self.checkboxes = []

        doc_type = self.doc_type_combo.currentText()
        # Загружаем из менеджера правил
        documents = self.rule_manager.get_rules_for(doc_type)

        for doc_template in documents:
            cb = QCheckBox(doc_template)
            cb.template = doc_template
            self.checkbox_layout.addWidget(cb)
            self.checkboxes.append(cb)

        self.update_document_preview()

    def update_document_preview(self):
        case_number = self.get_case_number()
        for cb in self.checkboxes:
            tpl = getattr(cb, 'template', '')
            if tpl:
                display = tpl.replace("{номер}", case_number if case_number else "XXX")
                cb.setText(display)

    def select_all(self):
        for cb in self.checkboxes: cb.setChecked(True)

    def clear_all(self):
        for cb in self.checkboxes: cb.setChecked(False)

    def get_selected_documents(self):
        selected = []
        case_number = self.get_case_number()
        for cb in self.checkboxes:
            if cb.isChecked():
                tpl = getattr(cb, 'template', '')
                if tpl:
                    selected.append(tpl.replace("{номер}", case_number if case_number else "XXX"))
        return selected

    def get_case_number(self):
        txt = self.case_number_edit.text().strip()
        return re.sub(r'\D', '', txt)


class FileRenamerApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Мастер переименования файлов")
        self.resize(1000, 700)

        # Инициализация менеджера правил
        self.rule_manager = RuleManager()

        self.destination_folder = ""
        # Теперь храним список файлов как список кортежей (имя, полный_путь) или просто полных путей
        # Для упрощения Reordering будем хранить пути в widget, а здесь только исходный источник
        self.current_files_map = {}  # path -> original_path

        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        # Инструкция
        layout.addWidget(QLabel("1. Перетащите файлы слева (можно менять их порядок перетаскиванием).\n"
                                "2. Выберите тип (СИ/ИК) и настройте шаблоны при необходимости справа.\n"
                                "3. Введите номер и нажмите переименовать."))

        # Главная область
        main_layout = QHBoxLayout()

        # ЛЕВАЯ ЧАСТЬ
        left_frame = QVBoxLayout()
        left_frame.addWidget(QLabel("Файлы (Drag&Drop для сортировки):"))

        btns_file = QHBoxLayout()
        self.btn_clear = QPushButton("Очистить")
        self.btn_clear.clicked.connect(self.clear_files)
        self.btn_del_sel = QPushButton("Удалить выбранные")
        self.btn_del_sel.clicked.connect(self.remove_selected)
        btns_file.addWidget(self.btn_clear)
        btns_file.addWidget(self.btn_del_sel)
        left_frame.addLayout(btns_file)

        self.list_dest = DropListWidget(self, role="dest")
        left_frame.addWidget(self.list_dest)
        main_layout.addLayout(left_frame)

        # ПРАВАЯ ЧАСТЬ
        right_frame = QVBoxLayout()
        self.document_selector = DocumentSelector(self.rule_manager)
        right_frame.addWidget(self.document_selector)
        main_layout.addLayout(right_frame)

        layout.addLayout(main_layout)

        # НИЖНЯЯ ПАНЕЛЬ
        action_layout = QHBoxLayout()

        self.btn_preview = QPushButton("Предпросмотр")
        self.btn_preview.clicked.connect(self.show_preview)

        self.btn_rename = QPushButton("🚀 Переименовать")
        self.btn_rename.setStyleSheet("font-weight: bold; font-size: 14pt; padding: 10px; background-color: #d1e7dd;")
        self.btn_rename.clicked.connect(self.rename_files)

        action_layout.addWidget(self.btn_preview)
        action_layout.addStretch()
        action_layout.addWidget(self.btn_rename)

        layout.addLayout(action_layout)
        self.setLayout(layout)

    def set_folder(self, folder, role):
        self.list_dest.clear()
        self.current_files_map = {}

        files = [f for f in os.listdir(folder) if os.path.splitext(f)[1].lower() in VALID_EXTENSIONS]
        # Сортировка по числам в имени
        files.sort(key=lambda x: (self.extract_number(x), x))

        for f in files:
            full_path = os.path.join(folder, f)
            self.list_dest.addItem(f)
            self.current_files_map[f] = full_path

    def add_individual_files(self, paths, role):
        # Если добавляем файлы, а не папку, просто докидываем в список
        # Сортируем добавляемую пачку
        paths.sort(key=lambda x: (self.extract_number(os.path.basename(x)), x))

        for p in paths:
            if os.path.splitext(p)[1].lower() in VALID_EXTENSIONS:
                name = os.path.basename(p)
                # Избегаем дублей визуально, если нужно (здесь разрешим, но лучше проверять)
                self.list_dest.addItem(name)
                self.current_files_map[name] = p

    def extract_number(self, text):
        match = re.search(r'(\d+)', text)
        return int(match.group(1)) if match else 999999

    def sync_list_order(self):
        """Метод вызывается ListWidget после перетаскивания строк"""
        pass  # Логика уже в визуальном порядке элементов list_dest

    def clear_files(self):
        self.list_dest.clear()
        self.current_files_map = {}

    def remove_selected(self):
        for item in self.list_dest.selectedItems():
            self.list_dest.takeItem(self.list_dest.row(item))
            # Из map не удаляем, так как имя файла уникально для пути, пусть висит в памяти

    def get_ordered_files(self):
        files = []
        for i in range(self.list_dest.count()):
            name = self.list_dest.item(i).text()
            if name in self.current_files_map:
                files.append(self.current_files_map[name])
        return files

    def show_preview(self):
        files = self.get_ordered_files()
        if not files:
            QMessageBox.warning(self, "Пусто", "Нет файлов для предпросмотра.")
            return

        docs = self.document_selector.get_selected_documents()
        if not docs:
            QMessageBox.warning(self, "Внимание", "Не выбраны документы справа!")
            return

        lines = []
        limit = min(len(files), len(docs), 20)
        for i in range(limit):
            old = os.path.basename(files[i])
            ext = os.path.splitext(files[i])[1]
            new = docs[i] + ext
            lines.append(f"{old}  ->  {new}")

        if len(files) > limit and len(docs) > limit:
            lines.append(f"... и ещё {min(len(files), len(docs)) - limit} строк")

        preview_text = "\n".join(lines) + f"\n\nВсего файлов: {len(files)}, Шаблонов: {len(docs)}"

        msg = QMessageBox(self)
        msg.setWindowTitle("Превью переименований")
        msg.setText(preview_text)

        # Увеличиваем ширину окна за счёт ширины QLabel внутри
        msg.setStyleSheet("""
            QMessageBox QLabel {
                min-width: 800px;
            }
        """)

        msg.exec()

    def rename_files(self):
        files = self.get_ordered_files()
        docs = self.document_selector.get_selected_documents()

        if not files or not docs:
            QMessageBox.warning(self, "Ошибка", "Нет файлов или не выбраны шаблоны.")
            return

        if len(files) != len(docs):
            reply = QMessageBox.question(self, "Несовпадение",
                                         f"Файлов: {len(files)}\nШаблонов: {len(docs)}\n\nПродолжить (лишнее будет проигнорировано)?",
                                         QMessageBox.Yes | QMessageBox.No)
            if reply != QMessageBox.Yes:
                return

        count = 0
        limit = min(len(files), len(docs))

        for i in range(limit):
            old_path = files[i]
            folder = os.path.dirname(old_path)
            ext = os.path.splitext(old_path)[1]
            new_name = docs[i] + ext
            new_path = os.path.join(folder, new_name)

            try:
                final_path = unique_path(new_path)
                os.rename(old_path, final_path)
                count += 1
            except Exception as e:
                print(f"Error renaming {old_path}: {e}")

        QMessageBox.information(self, "Готово", f"Переименовано {count} файлов.")
        self.list_dest.clear()
        self.current_files_map = {}


if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = FileRenamerApp()
    win.show()
    sys.exit(app.exec())
