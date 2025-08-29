import sys
import re
import tempfile
from pathlib import Path

from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QLabel,
    QFileDialog, QTextEdit
)
from PySide6.QtCore import Qt

from PyPDF2 import PdfMerger
from docx2pdf import convert

# Наборы документов
PACKAGE_RULES = {
    "Заказчик": ["02", "03", "04.1", "04.2", "05", "05.1", "06", "07",
                 "08", "09", "10", "11.1", "11.2", "11.3", "12", "13", "15"],
    "Линвит_подписание": ["02", "05", "05", "05.1", "05.1", "06", "08", "09", "10"],
    "Линвит_наши": ["01", "03", "04.1", "04.2", "05.2", "07",
                    "11.1", "11.2", "11.3", "12", "13", "14"],
}

def extract_doc_number(filename: str) -> str | None:
    """Извлекаем номер документа из имени файла (например, '03.1 договор.pdf')."""
    match = re.match(r"^(\d+(?:\.\d+)?)", filename.strip())
    return match.group(1) if match else None

def parse_number(num_str: str) -> tuple[int, int]:
    """Преобразует строку '03.2' → (3, 2), '05' → (5, 0)."""
    if "." in num_str:
        main, sub = num_str.split(".")
        return int(main), int(sub)
    return int(num_str), 0

class DropArea(QTextEdit):
    """Зона для Drag&Drop и хранения файлов."""
    def __init__(self, status_label: QLabel):
        super().__init__()
        self.setAcceptDrops(True)
        self.setPlaceholderText("Перетащите сюда PDF или Word файлы")
        self.setReadOnly(True)

        self.files: list[Path] = []                # оригиналы (docx/pdf)
        self.original_to_pdf: dict[Path, Path] = {} # сопоставление оригинал → pdf
        self.temp_files: list[Path] = []           # временные pdf после конвертации
        self.status_label = status_label

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                f = Path(url.toLocalFile())
                if f.suffix.lower() in [".pdf", ".docx", ".doc"]:
                    pdf_path = self.ensure_pdf(f)
                    self.files.append(f)
                    self.original_to_pdf[f] = pdf_path
            self.refresh()
            event.acceptProposedAction()
        else:
            event.ignore()

    def ensure_pdf(self, path: Path) -> Path:
        """Если файл Word — конвертируем его во временный PDF и возвращаем путь."""
        if path.suffix.lower() in [".docx", ".doc"]:
            tmp_dir = Path(tempfile.gettempdir())
            out_path = tmp_dir / (path.stem + "_converted.pdf")

            self.status_label.setText(f"⏳ Конвертация {path.name} ...")
            convert(str(path), str(out_path))
            self.status_label.setText(f"✅ {path.name} преобразован в PDF")

            self.temp_files.append(out_path)
            return out_path
        return path

    def refresh(self):
        # сортировка по номерам документов
        def sort_key(f: Path):
            num = extract_doc_number(f.name)
            return parse_number(num) if num else (999, 999)

        self.files.sort(key=sort_key)
        self.setPlainText("\n".join(f.name for f in self.files))

    def clear_files(self):
        """Очищаем список файлов и удаляем временные PDF"""
        for tmp in self.temp_files:
            try:
                if tmp.exists():
                    tmp.unlink()
            except Exception:
                pass
        self.temp_files.clear()
        self.files.clear()
        self.original_to_pdf.clear()
        self.clear()
        self.setPlaceholderText("Перетащите сюда PDF или Word файлы")
        self.status_label.setText("🗑 Список файлов очищен")

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Формирование пакетов документов")
        self.resize(800, 600)
        layout = QVBoxLayout(self)

        self.label = QLabel("Выберите файлы для работы:")
        layout.addWidget(self.label)

        # Кнопки управления
        self.btn_choose = QPushButton("Выбрать файлы")
        self.btn_choose.setMinimumHeight(40)
        layout.addWidget(self.btn_choose)

        self.btn_clear = QPushButton("Очистить список файлов")
        self.btn_clear.setMinimumHeight(40)
        layout.addWidget(self.btn_clear)

        # Статус
        self.status = QLabel("")
        layout.addWidget(self.status)

        # Зона Drag&Drop
        self.drop_area = DropArea(self.status)
        self.drop_area.setMinimumHeight(300)
        layout.addWidget(self.drop_area)

        # Кнопки пакетов
        self.btn_customer = QPushButton("Сформировать пакет для Заказчика")
        self.btn_sign = QPushButton("Сформировать пакет Линвит (подписание)")
        self.btn_ours = QPushButton("Сформировать пакет Линвит (наши подписи)")

        for btn in (self.btn_customer, self.btn_sign, self.btn_ours):
            btn.setMinimumHeight(40)
            layout.addWidget(btn)

        # Сигналы
        self.btn_choose.clicked.connect(self.choose_files)
        self.btn_clear.clicked.connect(self.clear_files)
        self.btn_customer.clicked.connect(lambda: self.make_package("Заказчик"))
        self.btn_sign.clicked.connect(lambda: self.make_package("Линвит_подписание"))
        self.btn_ours.clicked.connect(lambda: self.make_package("Линвит_наши"))

    def choose_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "Выберите PDF или Word файлы", "", "Документы (*.pdf *.docx *.doc)"
        )
        if files:
            for f in files:
                path = Path(f)
                pdf_path = self.drop_area.ensure_pdf(path)
                self.drop_area.files.append(path)
                self.drop_area.original_to_pdf[path] = pdf_path
            self.drop_area.refresh()

    def clear_files(self):
        self.drop_area.clear_files()

    def make_package(self, package_name: str):
        files_map: dict[str, list[Path]] = {}

        for orig in self.drop_area.files:
            num = extract_doc_number(orig.name)
            if num:
                pdf_path = self.drop_area.original_to_pdf[orig]
                files_map.setdefault(num, []).append(pdf_path)

        rules = PACKAGE_RULES[package_name]
        missing = []
        merger = PdfMerger()

        for doc in rules:
            if doc in files_map and files_map[doc]:
                merger.append(str(files_map[doc][0]))
            else:
                missing.append(doc)

        if not merger.pages:
            self.status.setText(f"❌ Нет документов для пакета {package_name}")
            return

        out_dir = self.drop_area.files[0].parent if self.drop_area.files else Path.cwd()
        out_file = out_dir / f"{package_name}.pdf"
        merger.write(str(out_file))
        merger.close()

        msg = f"✅ Пакет {package_name} сформирован: {out_file.name}"
        if missing:
            msg += f"\n⚠️ Отсутствуют документы: {', '.join(missing)}"
        self.status.setText(msg)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())