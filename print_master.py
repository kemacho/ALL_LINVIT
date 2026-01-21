# save as packer_app.py
import sys
import re
import tempfile
import json
from copy import deepcopy
from pathlib import Path
from typing import List, Dict

from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QFileDialog, QTextEdit, QProgressBar, QComboBox, QTabWidget,
    QSpinBox, QListWidget, QListWidgetItem, QAbstractItemView
)
from PySide6.QtCore import Qt, QThread, Signal, QSettings

from PyPDF2 import PdfMerger
from docx2pdf import convert

# -------------------- Документные названия --------------------

SI_DOC_TITLES = {
    "01": "Распоряжение по заявке",
    "02": "Решение по заявке",
    "02.1": "Документы для процедуры СИ",
    "02.2": "Вопросник по АСП_СИ",
    "03": "Заключение ОМТД",
    "04": "Акт выбора ПК",
    "04.1": "Направление-заявка СИ",
    "05": "Программа СИ",
    "06": "Заключение протоколы СИ",
    "07": "Программа АСП",
    "08": "Акт АСП",
    "09": "Распоряжение (анализ результатов)",
    "10": "Решение о выдаче сертификата",
    "13": "Перечень документов по СИ",
    "14": "Бланк сопров. письмо",
}

# если нужны специфичные названия для ИК — можно заполнить
IK_DOC_TITLES = {
    # пример; заполни по аналогии при необходимости
    "01": "Распоряжение ИК",
    "02": "Извещение ИК",
    "03": "Программа ИК",
    "03.1": "Документы для процедуры ИК",
    "03.2": "Вопросник по АСП_ИК",
    "04.1": "Акт по ОМТД ИК",
    "04.2": "Акт по МК ИК",
    "05.1": "Акт выбора ПК ИК",
    "05.2": "Направление-заявка ИК",
    "06": "Программа испыт. ИК",
    "07": "Акт по испыт. ИК",
    "08": "Программа АСП ИК",
    "09": "Акт АСП ИК",
    "10": "Акт по ИК",
    "11.1": "Решение по ИК",
    "14": "Перечень документов по ИК (в папку)",
    "15": "Бланк сопров. письмо (Пакет по ИК)",
}

# -------------------- Правила пакетов по умолчанию --------------------

SI_RULES_DEFAULT = {
    "Заказчик": ["14", "02", "03", "04", "05", "06", "07", "08", "10"],
    "Линвит_подписание": ["02.2", "04", "04", "05", "07", "08"],
    "Линвит_наши": ["13", "15", "01", "02", "03", "04.1", "04.1", "06", "09", "10"],
}

IK_RULES_DEFAULT = {
    "Заказчик": ["15", "02", "03", "04.1", "05.1", "06", "07", "08", "09", "10", "11.1"],
    "Линвит_подписание": ["03", "03.2", "05.1", "05.1", "06", "08", "09", "10"],
    "Линвит_наши": ["14", "01", "02", "04.1", "05.2", "05.2", "07", "11.1"]
}

# рабочие (редактируемые) копии
SI_RULES = deepcopy(SI_RULES_DEFAULT)
IK_RULES = deepcopy(IK_RULES_DEFAULT)

# словарь, с которым работает WorkerThread (обновляется при смене профиля / правок в UI)
PACKAGE_RULES = deepcopy(IK_RULES_DEFAULT)


# -------------------- Утилиты --------------------

def extract_doc_number(filename: str) -> str | None:
    match = re.match(r"^(\d+(?:\.\d+)?)", filename.strip())
    return match.group(1) if match else None


def parse_number(num_str: str) -> tuple[int, int]:
    if "." in num_str:
        main, sub = num_str.split(".")
        return int(main), int(sub)
    return int(num_str), 0


# -------------------- Виджет элемента документа в списке --------------------

class DocItemWidget(QWidget):
    def __init__(self, num: str, title: str, count: int = 0, parent=None):
        super().__init__(parent)
        self.num = num

        layout = QHBoxLayout(self)
        layout.setContentsMargins(2, 2, 2, 2)

        self.lbl = QLabel(f"{num}. {title}")
        self.spin = QSpinBox()
        self.spin.setMinimum(0)
        self.spin.setMaximum(50)
        self.spin.setValue(count)

        layout.addWidget(self.lbl)
        layout.addWidget(self.spin)

    def get_count(self) -> int:
        return self.spin.value()


# -------------------- WorkerThread --------------------

class WorkerThread(QThread):
    progress = Signal(int)
    status = Signal(str)
    finished_signal = Signal(str)
    temp_files_signal = Signal(list)

    def __init__(
        self,
        originals: List[Path],
        mode: str,
        package_name: str | None = None,
        out_dir: Path | None = None
    ):
        super().__init__()
        self.originals = originals
        self.mode = mode
        self.package_name = package_name
        self.out_dir = out_dir
        self._stop_requested = False

    def run(self):
        try:
            self.status.emit("Начинается конвертация (если есть Word-файлы)...")
            total_files = len(self.originals) if self.originals else 1
            converted_map: Dict[Path, Path] = {}
            temp_files: List[Path] = []

            # 1) Конвертация
            for i, orig in enumerate(self.originals, start=1):
                if self._stop_requested:
                    self.status.emit("Операция прервана")
                    self.finished_signal.emit("Операция прервана")
                    return

                if orig.suffix.lower() in [".docx", ".doc"]:
                    self.status.emit(f"⏳ Конвертация {orig.name} ...")
                    tmp_dir = Path(tempfile.gettempdir())
                    out_path = tmp_dir / (orig.stem + "_converted.pdf")
                    try:
                        convert(str(orig), str(out_path))
                    except Exception as e:
                        self.status.emit(f"❌ Ошибка конвертации {orig.name}: {e}")
                        converted_map[orig] = out_path
                    else:
                        converted_map[orig] = out_path
                        temp_files.append(out_path)
                        self.status.emit(f"✅ {orig.name} → PDF")
                else:
                    converted_map[orig] = orig

                self.progress.emit(int(i / total_files * 50))

            if temp_files:
                self.temp_files_signal.emit(temp_files)

            # 2) Сборка пакетов
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

    def _build_single_package(
        self,
        converted_map: Dict[Path, Path],
        package_key: str,
        start_progress: int = 50
    ):
        rules = PACKAGE_RULES.get(package_key, [])
        self.status.emit(f"Сборка пакета: {package_key} ...")

        files_map: Dict[str, List[Path]] = {}
        for orig, pdf in converted_map.items():
            num = extract_doc_number(orig.name)
            if num:
                files_map.setdefault(num, []).append(pdf)

        merger = PdfMerger()
        total = len(rules) if rules else 1
        processed = 0
        missing: List[str] = []
        appended_count = 0

        for doc in rules:
            if self._stop_requested:
                self.status.emit("Операция прервана")
                self.finished_signal.emit("Операция прервана")
                return
            if doc in files_map and files_map[doc]:
                try:
                    merger.append(str(files_map[doc][0]))
                    appended_count += 1
                except Exception as e:
                    self.status.emit(f"⚠️ Ошибка при добавлении {doc}: {e}")
                    missing.append(doc)
            else:
                missing.append(doc)

            processed += 1
            self.progress.emit(start_progress + int(processed / total * (100 - start_progress)))

        if appended_count == 0:
            self.status.emit(f"❌ Нет документов для пакета {package_key}")
            self.finished_signal.emit(f"Пакет {package_key} не сформирован")
            return

        out_dir = self.out_dir or (
            next(iter(converted_map.keys())).parent if converted_map else Path.cwd()
        )
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

    def _build_all_packages(
        self,
        converted_map: Dict[Path, Path],
        start_progress: int = 50
    ):
        pkg_keys = list(PACKAGE_RULES.keys())
        total_docs = sum(len(PACKAGE_RULES[k]) for k in pkg_keys) or 1
        cumulative = 0
        results: List[str] = []

        # общая карта номеров -> pdf
        files_map: Dict[str, List[Path]] = {}
        for orig, pdf in converted_map.items():
            num = extract_doc_number(orig.name)
            if num:
                files_map.setdefault(num, []).append(pdf)

        for pkg in pkg_keys:
            if self._stop_requested:
                self.status.emit("Операция прервана")
                self.finished_signal.emit("Операция прервана")
                return
            self.status.emit(f"Сборка пакета: {pkg} ...")

            merger = PdfMerger()
            missing: List[str] = []
            appended_count = 0

            for doc in PACKAGE_RULES[pkg]:
                if doc in files_map and files_map[doc]:
                    try:
                        merger.append(str(files_map[doc][0]))
                        appended_count += 1
                    except Exception as e:
                        self.status.emit(f"⚠️ Ошибка при добавлении {doc} в {pkg}: {e}")
                        missing.append(doc)
                else:
                    missing.append(doc)
                cumulative += 1
                self.progress.emit(
                    start_progress + int(cumulative / total_docs * (100 - start_progress))
                )

            try:
                if appended_count > 0:
                    out_dir = self.out_dir or (
                        next(iter(converted_map.keys())).parent if converted_map else Path.cwd()
                    )
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


# -------------------- DropArea --------------------

class DropArea(QTextEdit):
    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)
        self.setPlaceholderText("Перетащите сюда PDF или Word файлы")
        self.setReadOnly(True)

        self.files: List[Path] = []
        self.temp_files: List[Path] = []

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
            if p.exists() and p.is_file() and p.suffix.lower() in [".pdf", ".docx", ".doc"]:
                self.files.append(p)
        self.refresh()

    def refresh(self):
        def sort_key(f: Path):
            num = extract_doc_number(f.name)
            return parse_number(num) if num else (999, 999)

        self.files.sort(key=sort_key)
        self.setPlainText("\n".join(f.name for f in self.files))

    def clear_files(self):
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


# -------------------- MainWindow --------------------

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Формирование пакетов документов")
        self.resize(1200, 750)

        self.settings = QSettings("YourCompany", "PackerApp")
        self.current_profile = "ИК"

        root = QHBoxLayout(self)

        # левая часть
        left = QVBoxLayout()
        root.addLayout(left, 3)

        lbl = QLabel("Выберите файлы для работы:")
        left.addWidget(lbl)

        top_h = QHBoxLayout()
        self.btn_choose = QPushButton("Выбрать файлы")
        self.btn_choose.setMinimumHeight(36)
        top_h.addWidget(self.btn_choose)

        self.btn_clear = QPushButton("Очистить список файлов")
        self.btn_clear.setMinimumHeight(36)
        top_h.addWidget(self.btn_clear)

        left.addLayout(top_h)

        self.status = QLabel("")
        left.addWidget(self.status)

        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        left.addWidget(self.progress)

        self.drop_area = DropArea()
        self.drop_area.setMinimumHeight(360)
        left.addWidget(self.drop_area)

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

        left.addLayout(btns_h)

        # правая часть
        right = QVBoxLayout()
        root.addLayout(right, 2)

        profile_h = QHBoxLayout()
        profile_label = QLabel("Профиль:")
        self.profile_combo = QComboBox()
        self.profile_combo.addItems(["ИК", "СИ"])
        profile_h.addWidget(profile_label)
        profile_h.addWidget(self.profile_combo)
        right.addLayout(profile_h)

        self.btn_defaults = QPushButton("Установить правила по умолчанию")
        right.addWidget(self.btn_defaults)

        self.tabs = QTabWidget()
        right.addWidget(self.tabs, 1)

        self.tab_customer = QWidget()
        self.tab_sign = QWidget()
        self.tab_ours = QWidget()

        self.tabs.addTab(self.tab_customer, "Заказчик")
        self.tabs.addTab(self.tab_sign, "Линвит (подп.)")
        self.tabs.addTab(self.tab_ours, "Линвит (наши)")

        self.package_lists: Dict[str, Dict] = {}

        self._setup_package_tab(self.tab_customer, "Заказчик")
        self._setup_package_tab(self.tab_sign, "Линвит_подписание")
        self._setup_package_tab(self.tab_ours, "Линвит_наши")

        # сигналы
        self.btn_choose.clicked.connect(self.choose_files)
        self.btn_clear.clicked.connect(self.clear_files)
        self.btn_customer.clicked.connect(
            lambda: self.start_worker(mode="single", package_name="Заказчик")
        )
        self.btn_sign.clicked.connect(
            lambda: self.start_worker(mode="single", package_name="Линвит_подписание")
        )
        self.btn_ours.clicked.connect(
            lambda: self.start_worker(mode="single", package_name="Линвит_наши")
        )
        self.btn_all.clicked.connect(lambda: self.start_worker(mode="all"))

        self.profile_combo.currentTextChanged.connect(self.on_profile_changed)
        self.btn_defaults.clicked.connect(self.on_reset_defaults)

        self.worker: WorkerThread | None = None

        # загрузка настроек и начальное заполнение
        self.load_settings()
        self.apply_profile_rules()
        self.load_rules_to_ui()

    # ---------- вкладки / списки ----------

    def _setup_package_tab(self, tab_widget: QWidget, package_key: str):
        layout = QVBoxLayout(tab_widget)
        lw = QListWidget()
        lw.setDragDropMode(QAbstractItemView.InternalMove)
        lw.setDefaultDropAction(Qt.MoveAction)
        lw.setSelectionMode(QAbstractItemView.SingleSelection)
        layout.addWidget(lw)
        self.package_lists[package_key] = {"list": lw}

    def load_rules_to_ui(self):
        # очистить
        for pkg_key, info in self.package_lists.items():
            lw: QListWidget = info["list"]
            lw.clear()

        # из PACKAGE_RULES -> элементы
        for pkg_key, info in self.package_lists.items():
            lw: QListWidget = info["list"]
            docs = PACKAGE_RULES.get(pkg_key, [])

            counts: Dict[str, int] = {}
            for num in docs:
                counts[num] = counts.get(num, 0) + 1

            for num, count in counts.items():
                if self.current_profile == "СИ":
                    title = SI_DOC_TITLES.get(num, "")
                else:
                    title = IK_DOC_TITLES.get(num, "")
                item = QListWidgetItem(lw)
                widget = DocItemWidget(num, title, count)
                item.setSizeHint(widget.sizeHint())
                lw.addItem(item)
                lw.setItemWidget(item, widget)

    def save_rules_from_ui(self):
        global PACKAGE_RULES, SI_RULES, IK_RULES

        new_rules: Dict[str, List[str]] = {}

        for pkg_key, info in self.package_lists.items():
            lw: QListWidget = info["list"]
            docs: List[str] = []

            for row in range(lw.count()):
                item = lw.item(row)
                widget = lw.itemWidget(item)
                if not isinstance(widget, DocItemWidget):
                    continue
                num = widget.num
                cnt = widget.get_count()
                if cnt > 0:
                    docs.extend([num] * cnt)

            new_rules[pkg_key] = docs

        PACKAGE_RULES = new_rules

        if self.current_profile == "ИК":
            IK_RULES = deepcopy(PACKAGE_RULES)
        else:
            SI_RULES = deepcopy(PACKAGE_RULES)

    def apply_profile_rules(self):
        global PACKAGE_RULES
        if self.current_profile == "ИК":
            PACKAGE_RULES = deepcopy(IK_RULES)
        else:
            PACKAGE_RULES = deepcopy(SI_RULES)

    # ---------- профиль / дефолты / настройки ----------

    def on_profile_changed(self, profile: str):
        # перед сменой профиля — сохранить в текущий профиль
        self.save_rules_from_ui()
        self.current_profile = profile
        self.apply_profile_rules()
        self.load_rules_to_ui()

    def on_reset_defaults(self):
        global SI_RULES, IK_RULES
        if self.current_profile == "ИК":
            IK_RULES = deepcopy(IK_RULES_DEFAULT)
        else:
            SI_RULES = deepcopy(SI_RULES_DEFAULT)
        self.apply_profile_rules()
        self.load_rules_to_ui()
        self.status.setText("Правила текущего профиля сброшены к значениям по умолчанию.")

    def load_settings(self):
        global SI_RULES, IK_RULES

        data = self.settings.value("rules_json", "")
        profile = self.settings.value("profile", "ИК")
        if isinstance(profile, str):
            self.current_profile = profile
        else:
            self.current_profile = "ИК"

        if data:
            try:
                all_profiles = json.loads(data)
                si = all_profiles.get("СИ")
                ik = all_profiles.get("ИК")
                if isinstance(si, dict):
                    SI_RULES = si
                if isinstance(ik, dict):
                    IK_RULES = ik
            except Exception:
                SI_RULES = deepcopy(SI_RULES_DEFAULT)
                IK_RULES = deepcopy(IK_RULES_DEFAULT)
        else:
            SI_RULES = deepcopy(SI_RULES_DEFAULT)
            IK_RULES = deepcopy(IK_RULES_DEFAULT)

        self.profile_combo.setCurrentText(self.current_profile)

    def save_settings(self):
        # сохранить оба профиля в настройки
        all_profiles = {
            "СИ": SI_RULES,
            "ИК": IK_RULES,
        }
        self.settings.setValue("rules_json", json.dumps(all_profiles, ensure_ascii=False))
        self.settings.setValue("profile", self.current_profile)

    # ---------- файлы / поток ----------

    def choose_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "Выберите PDF или Word файлы", "", "Документы (*.pdf *.docx *.doc)"
        )
        if files:
            self.drop_area.add_files([Path(f) for f in files])

    def clear_files(self):
        if self.worker and self.worker.isRunning():
            self.worker.request_stop()
        self.drop_area.clear_files()
        self.progress.setValue(0)
        self.status.setText("🗑 Список файлов очищен")

    def set_ui_enabled(self, enabled: bool):
        for w in (
            self.btn_choose, self.btn_clear,
            self.btn_customer, self.btn_sign, self.btn_ours, self.btn_all,
            self.profile_combo, self.btn_defaults
        ):
            w.setEnabled(enabled)

    def start_worker(self, mode: str, package_name: str | None = None):
        if not self.drop_area.files:
            self.status.setText("❗ Список файлов пуст — добавьте файлы сначала.")
            return

        # сохранить актуальные правила из UI перед запуском
        self.save_rules_from_ui()
        self.save_settings()

        self.set_ui_enabled(False)
        self.progress.setValue(0)
        self.status.setText("Запуск задачи...")

        out_dir = self.drop_area.files[0].parent if self.drop_area.files else Path.cwd()

        self.worker = WorkerThread(
            list(self.drop_area.files),
            mode=mode,
            package_name=package_name,
            out_dir=out_dir
        )
        self.worker.progress.connect(self.progress.setValue)
        self.worker.status.connect(self.status.setText)
        self.worker.temp_files_signal.connect(self._on_temp_files_created)
        self.worker.finished_signal.connect(self._on_worker_finished)
        self.worker.start()

    def _on_temp_files_created(self, temp_files: list):
        for p in temp_files:
            try:
                path = Path(p)
                if path not in self.drop_area.temp_files:
                    self.drop_area.temp_files.append(path)
            except Exception:
                pass

    def _on_worker_finished(self, msg: str):
        self.set_ui_enabled(True)
        self.progress.setValue(100)
        self.status.setText(msg)
        self.worker = None

    def closeEvent(self, event):
        # при закрытии сохраняем текущие настройки
        self.save_rules_from_ui()
        self.save_settings()
        super().closeEvent(event)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())
