# save as packer_app.py
import sys
import re
import tempfile
from pathlib import Path
from typing import List, Dict

from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QFileDialog, QTextEdit, QProgressBar
)
from PySide6.QtCore import Qt, QThread, Signal

from PyPDF2 import PdfMerger
from docx2pdf import convert

# ---- правила пакетов
PACKAGE_RULES = {
    "Заказчик": ["02", "02.1", "02.2", "03", "04", "05", "06", "07",
                 "08", "10", "11", "12"],
    "Линвит_подписание": ["04", "04", "05", "07", "08"],
    "Линвит_наши": ["01", "03", "04.1", "04.1", "06", "09",
                    "10", "11", "12", "13", "15"],
}
# PACKAGE_RULES = {
#     "Заказчик": ["02", "02.1", "02.2", "03", "04", "05", "06", "07",
#                  "08", "10", "11", "12"],
#     "Линвит_подписание": ["04", "04", "05", "07", "08"],
#     "Линвит_наши": ["01", "03", "04.1", "04.1", "06", "09",
#                     "10", "11", "12", "13", "15"],
# }

# ---- утилиты для номеров
def extract_doc_number(filename: str) -> str | None:
    match = re.match(r"^(\d+(?:\.\d+)?)", filename.strip())
    return match.group(1) if match else None

def parse_number(num_str: str) -> tuple[int, int]:
    if "." in num_str:
        main, sub = num_str.split(".")
        return int(main), int(sub)
    return int(num_str), 0

# ---- рабочий поток: конвертация + сборка пакетов
class WorkerThread(QThread):
    progress = Signal(int)            # 0..100
    status = Signal(str)              # текстовый статус
    finished_signal = Signal(str)     # сообщение о завершении
    temp_files_signal = Signal(list)  # список временных pdf, чтобы UI сохранил их

    def __init__(self, originals: List[Path], mode: str, package_name: str | None = None, out_dir: Path | None = None):
        """
        originals: список оригинальных Path (docx/doc/pdf)
        mode: "single" или "all"
        package_name: если mode == "single" — имя пакета (ключ в PACKAGE_RULES)
        out_dir: куда сохранять итоговые PDF (если None — рядом с первым оригиналом или cwd)
        """
        super().__init__()
        self.originals = originals
        self.mode = mode
        self.package_name = package_name
        self.out_dir = out_dir
        self._stop_requested = False

    def run(self):
        try:
            # 1) Конвертация Word -> временные PDF (шкала 0..50)
            self.status.emit("Начинается конвертация (если есть Word-файлы)...")
            total_files = len(self.originals) if self.originals else 1
            converted_map: Dict[Path, Path] = {}
            temp_files: List[Path] = []

            for i, orig in enumerate(self.originals, start=1):
                if self._stop_requested:
                    self.status.emit("Операция прервана")
                    return
                if orig.suffix.lower() in [".docx", ".doc"]:
                    self.status.emit(f"⏳ Конвертация {orig.name} ...")
                    # сохраняем с приставкой _converted, в temp
                    tmp_dir = Path(tempfile.gettempdir())
                    out_path = tmp_dir / (orig.stem + "_converted.pdf")
                    try:
                        convert(str(orig), str(out_path))
                    except Exception as e:
                        # conversion failed
                        self.status.emit(f"❌ Ошибка конвертации {orig.name}: {e}")
                        # но продолжаем дальше, отмечая этот файл как несуществующий
                        converted_map[orig] = out_path  # возможно пустой/битый — будет ловиться при склейке
                    else:
                        converted_map[orig] = out_path
                        temp_files.append(out_path)
                        self.status.emit(f"✅ {orig.name} → PDF")
                else:
                    # если это уже PDF, просто используем оригинал как pdf
                    converted_map[orig] = orig

                # прогресс конвертации: 0..50
                self.progress.emit(int(i / total_files * 50))

            # сигнализируем UI, чтобы он знал о временных файлах
            if temp_files:
                self.temp_files_signal.emit(temp_files)

            # 2) Сборка пакета(ов) (шкала 50..100)
            if self.mode == "single" and self.package_name:
                self._build_single_package(converted_map, self.package_name, start_progress=50)
            elif self.mode == "all":
                self._build_all_packages(converted_map, start_progress=50)
            else:
                self.status.emit("Неверный режим работы воркера")
                self.finished_signal.emit("Ошибка")
        except Exception as e:
            self.status.emit(f"Ошибка в потоке: {e}")
            self.finished_signal.emit("Ошибка")

    def _build_single_package(self, converted_map: Dict[Path, Path], package_key: str, start_progress: int = 50):
        rules = PACKAGE_RULES.get(package_key, [])
        self.status.emit(f"Сборка пакета: {package_key} ...")
        # подготовим files_map: номер -> list[pdf_paths]
        files_map: Dict[str, List[Path]] = {}
        for orig, pdf in converted_map.items():
            num = extract_doc_number(orig.name)
            if num:
                files_map.setdefault(num, []).append(pdf)

        merger = PdfMerger()
        total = len(rules) if rules else 1
        processed = 0
        missing = []

        for doc in rules:
            if self._stop_requested:
                self.status.emit("Операция прервана")
                return
            if doc in files_map and files_map[doc]:
                try:
                    merger.append(str(files_map[doc][0]))
                except Exception as e:
                    # ошибка при добавлении — отметим как пропуск
                    self.status.emit(f"⚠️ Ошибка при добавлении {doc}: {e}")
                    missing.append(doc)
            else:
                missing.append(doc)

            processed += 1
            # прогресс 50..100
            self.progress.emit(start_progress + int(processed / total * (100 - start_progress)))

        # если ничего не добавлено — сообщим и не создаём файл
        try:
            if not merger.pages:
                self.status.emit(f"❌ Нет документов для пакета {package_key}")
                self.finished_signal.emit(f"Пакет {package_key} не сформирован")
                return
        except Exception:
            # старые версии PyPDF2 могут не иметь .pages; проверим len
            # попробуем записать и поймать ошибку
            pass

        # где сохранять
        out_dir = self.out_dir or (next(iter(converted_map.keys())).parent if converted_map else Path.cwd())
        out_file = Path(out_dir) / f"{package_key}.pdf"
        try:
            merger.write(str(out_file))
            merger.close()
            msg = f"✅ Пакет {package_key} сформирован: {out_file.name}"
            if missing:
                msg += f"\n⚠️ Отсутствуют документы: {', '.join(missing)}"
            self.status.emit(msg)
            self.finished_signal.emit(msg)
        except Exception as e:
            self.status.emit(f"❌ Ошибка записи {out_file.name}: {e}")
            self.finished_signal.emit(f"Ошибка при формировании {package_key}")

    def _build_all_packages(self, converted_map: Dict[Path, Path], start_progress: int = 50):
        pkg_keys = list(PACKAGE_RULES.keys())
        total_docs = sum(len(PACKAGE_RULES[k]) for k in pkg_keys) or 1
        cumulative = 0
        results = []
        for pkg in pkg_keys:
            if self._stop_requested:
                self.status.emit("Операция прервана")
                return
            self.status.emit(f"Сборка пакета: {pkg} ...")

            # подготовка map
            files_map: Dict[str, List[Path]] = {}
            for orig, pdf in converted_map.items():
                num = extract_doc_number(orig.name)
                if num:
                    files_map.setdefault(num, []).append(pdf)

            merger = PdfMerger()
            missing = []
            for doc in PACKAGE_RULES[pkg]:
                if doc in files_map and files_map[doc]:
                    try:
                        merger.append(str(files_map[doc][0]))
                    except Exception as e:
                        self.status.emit(f"⚠️ Ошибка при добавлении {doc} в {pkg}: {e}")
                        missing.append(doc)
                else:
                    missing.append(doc)
                cumulative += 1
                # прогресс: 50..100 mapped by cumulative/total_docs
                self.progress.emit(start_progress + int(cumulative / total_docs * (100 - start_progress)))

            # сохраняем пакет (если есть страницы)
            try:
                if merger.pages:
                    out_dir = self.out_dir or (next(iter(converted_map.keys())).parent if converted_map else Path.cwd())
                    out_file = Path(out_dir) / f"{pkg}.pdf"
                    merger.write(str(out_file))
                    merger.close()
                    msg = f"✅ {pkg} сформирован: {out_file.name}"
                    if missing:
                        msg += f" (пропущены: {', '.join(missing)})"
                    results.append(msg)
                else:
                    results.append(f"❌ {pkg}: нет документов")
            except Exception as e:
                results.append(f"❌ {pkg}: ошибка записи ({e})")

        final_msg = "📦 Все пакеты сформированы\n" + "\n".join(results)
        self.status.emit(final_msg)
        self.finished_signal.emit(final_msg)

    def request_stop(self):
        self._stop_requested = True

# ---- UI
class DropArea(QTextEdit):
    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)
        self.setPlaceholderText("Перетащите сюда PDF или Word файлы")
        self.setReadOnly(True)

        self.files: List[Path] = []       # оригиналы
        self.temp_files: List[Path] = [] # временные pdf, которые создал поток

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
            paths = [Path(url.toLocalFile()) for url in event.mimeData().urls()]
            self.add_files(paths)
            event.acceptProposedAction()
        else:
            event.ignore()

    def add_files(self, paths: List[Path]):
        for p in paths:
            if p.suffix.lower() in [".pdf", ".docx", ".doc"]:
                self.files.append(p)
        self.refresh()

    def refresh(self):
        # сортировка по номеру (если нет номера — в конец)
        def sort_key(f: Path):
            num = extract_doc_number(f.name)
            return parse_number(num) if num else (999, 999)
        self.files.sort(key=sort_key)
        self.setPlainText("\n".join(f.name for f in self.files))

    def clear_files(self):
        # удаляем временные файлы, если есть
        for tmp in list(self.temp_files):
            try:
                if tmp.exists():
                    tmp.unlink()
            except Exception:
                pass
        self.temp_files.clear()
        self.files.clear()
        self.clear()
        self.setPlaceholderText("Перетащите сюда PDF или Word файлы")

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Формирование пакетов документов")
        self.resize(920, 700)

        root = QVBoxLayout(self)

        lbl = QLabel("Выберите файлы для работы:")
        root.addWidget(lbl)

        # верх: выбор и очистка
        top_h = QHBoxLayout()
        self.btn_choose = QPushButton("Выбрать файлы")
        self.btn_choose.setMinimumHeight(36)
        top_h.addWidget(self.btn_choose)

        self.btn_clear = QPushButton("Очистить список файлов")
        self.btn_clear.setMinimumHeight(36)
        top_h.addWidget(self.btn_clear)

        root.addLayout(top_h)

        # статус и прогресс
        self.status = QLabel("")
        root.addWidget(self.status)

        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        root.addWidget(self.progress)

        # drop area
        self.drop_area = DropArea()
        self.drop_area.setMinimumHeight(360)
        root.addWidget(self.drop_area)

        # кнопки пакетов и "все пакеты"
        btns_h = QHBoxLayout()
        self.btn_customer = QPushButton("Пакет для Заказчика")
        self.btn_sign = QPushButton("Пакет Линвит (подписание)")
        self.btn_ours = QPushButton("Пакет Линвит (наши подписи)")
        for b in (self.btn_customer, self.btn_sign, self.btn_ours):
            b.setMinimumHeight(60)
            btns_h.addWidget(b)

        self.btn_all = QPushButton("📦 Все пакеты")
        self.btn_all.setMinimumSize(60, 60)
        btns_h.addWidget(self.btn_all)

        root.addLayout(btns_h)

        # сигналы
        self.btn_choose.clicked.connect(self.choose_files)
        self.btn_clear.clicked.connect(self.clear_files)
        self.btn_customer.clicked.connect(lambda: self.start_worker(mode="single", package_name="Заказчик"))
        self.btn_sign.clicked.connect(lambda: self.start_worker(mode="single", package_name="Линвит_подписание"))
        self.btn_ours.clicked.connect(lambda: self.start_worker(mode="single", package_name="Линвит_наши"))
        self.btn_all.clicked.connect(lambda: self.start_worker(mode="all"))

        # текущее состояние worker-а
        self.worker: WorkerThread | None = None

    def choose_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "Выберите PDF или Word файлы", "", "Документы (*.pdf *.docx *.doc)"
        )
        if files:
            self.drop_area.add_files([Path(f) for f in files])

    def clear_files(self):
        # если есть запущенный поток — попросим остановиться (без гарантии мгновенного)
        if self.worker and self.worker.isRunning():
            self.worker.request_stop()
            # даём ему время — но UI разморожен в любом случае
        self.drop_area.clear_files()
        self.progress.setValue(0)
        self.status.setText("🗑 Список файлов очищен")

    def set_ui_enabled(self, enabled: bool):
        # включение/отключение кнопок, чтобы не запустить другую задачу
        for w in (self.btn_choose, self.btn_clear, self.btn_customer, self.btn_sign, self.btn_ours, self.btn_all):
            w.setEnabled(enabled)

    def start_worker(self, mode: str, package_name: str | None = None):
        if not self.drop_area.files:
            self.status.setText("❗ Список файлов пуст — добавьте файлы сначала.")
            return
        # блокируем UI элементы
        self.set_ui_enabled(False)
        self.progress.setValue(0)
        self.status.setText("Запуск задачи...")

        # куда сохранять: рядом с первым оригиналом
        out_dir = self.drop_area.files[0].parent if self.drop_area.files else Path.cwd()

        # создаём и запускаем поток
        self.worker = WorkerThread(list(self.drop_area.files), mode=mode, package_name=package_name, out_dir=out_dir)
        self.worker.progress.connect(self.progress.setValue)
        self.worker.status.connect(self.status.setText)
        self.worker.temp_files_signal.connect(self._on_temp_files_created)
        self.worker.finished_signal.connect(self._on_worker_finished)
        self.worker.start()

    def _on_temp_files_created(self, temp_files: list):
        # поток сообщил о созданных временных PDF — сохраним их в drop_area чтобы можно было удалить позднее
        for p in temp_files:
            try:
                path = Path(p)
                if path not in self.drop_area.temp_files:
                    self.drop_area.temp_files.append(path)
            except Exception:
                pass

    def _on_worker_finished(self, msg: str):
        # разблокируем UI
        self.set_ui_enabled(True)
        # прогресс до 100
        self.progress.setValue(100)
        # покажем итоговое сообщение
        self.status.setText(msg)
        self.worker = None

if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())