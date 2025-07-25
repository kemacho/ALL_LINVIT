import sys
import tempfile
import shutil
from pathlib import Path
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton,
    QFileDialog, QLabel, QMessageBox, QProgressBar
)
from PySide6.QtCore import Qt, QThread, Signal
import fitz  # PyMuPDF
from PIL import Image, ImageEnhance, ImageFilter, ImageDraw
import img2pdf
import numpy as np


def apply_enhanced_scan_effect(image: Image.Image) -> Image.Image:
    # Контраст и яркость
    image = ImageEnhance.Contrast(image).enhance(1.5)
    image = ImageEnhance.Brightness(image).enhance(1.1)

    # Лёгкая желтизна (имитация бумаги)
    yellow_overlay = Image.new("RGB", image.size, (255, 250, 210))
    image = Image.blend(image, yellow_overlay, alpha=0.1)

    # Слабый шум (зернистость)
    np_img = np.array(image).astype(np.int16)
    noise = np.random.normal(0, 12, np_img.shape)  # шум сильнее
    np_img = np.clip(np_img + noise, 0, 255).astype(np.uint8)
    image = Image.fromarray(np_img)

    # Легкое размытие (дефокус)
    image = image.filter(ImageFilter.GaussianBlur(radius=0.6))

    # Виньетирование (затемнение краёв)
    vignette = create_vignette_mask(image.size)
    image.putalpha(vignette)
    background = Image.new("RGB", image.size, (255, 250, 210))
    background.paste(image, mask=image.split()[3])  # альфа-канал для маски
    image = background

    return image


def create_vignette_mask(size):
    width, height = size
    vignette = Image.new("L", (width, height), 0)
    center_x, center_y = width // 2, height // 2
    max_radius = max(center_x, center_y)

    for y in range(height):
        for x in range(width):
            dx = x - center_x
            dy = y - center_y
            dist = (dx*dx + dy*dy) ** 0.5
            alpha = int(255 * min(dist / max_radius, 1))
            vignette.putpixel((x, y), alpha)

    vignette = vignette.filter(ImageFilter.GaussianBlur(radius=width // 10))
    return vignette

def pdf_to_images_with_fitz(pdf_path, dpi=200):
    doc = fitz.open(pdf_path)
    zoom = dpi / 72
    mat = fitz.Matrix(zoom, zoom)
    images = []

    for page in doc:
        pix = page.get_pixmap(matrix=mat, alpha=False)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        images.append(img)
    return images


class WorkerThread(QThread):
    progress_update = Signal(int)
    max_progress = Signal(int)
    message = Signal(str)

    def __init__(self, pdf_paths, out_dir):
        super().__init__()
        self.pdf_paths = pdf_paths
        self.out_dir = out_dir

    def run(self):
        total_pages = 0
        # Считаем общее количество страниц
        for pdf_path in self.pdf_paths:
            doc = fitz.open(pdf_path)
            total_pages += doc.page_count
            doc.close()
        self.max_progress.emit(total_pages)

        current_progress = 0

        for pdf_path in self.pdf_paths:
            base_name = Path(pdf_path).stem
            temp_dir = tempfile.mkdtemp()

            try:
                images = pdf_to_images_with_fitz(pdf_path, dpi=200)
                processed_paths = []

                for idx, img in enumerate(images):
                    enhanced = apply_enhanced_scan_effect(img)
                    save_path = Path(temp_dir) / f"page_{idx + 1}.jpg"
                    enhanced.save(save_path, "JPEG", quality=95)
                    processed_paths.append(str(save_path))

                    current_progress += 1
                    self.progress_update.emit(current_progress)

                output_pdf_path = Path(self.out_dir) / f"{base_name}_scanned.pdf"
                with open(output_pdf_path, "wb") as f:
                    f.write(img2pdf.convert(processed_paths))

                shutil.rmtree(temp_dir)

            except Exception as e:
                self.message.emit(f"Ошибка при обработке файла:\n{pdf_path}\n\n{str(e)}")

        self.message.emit("Обработка завершена!")


class PDFScannerApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF → Сканы → PDF (с прогрессом и эффектами)")
        self.setFixedSize(450, 250)

        layout = QVBoxLayout()

        self.label = QLabel("Выберите PDF-файлы для обработки")
        self.label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.label)

        self.select_btn = QPushButton("📂 Выбрать PDF-файлы")
        self.select_btn.clicked.connect(self.select_files)
        layout.addWidget(self.select_btn)

        self.convert_btn = QPushButton("🚀 Запустить обработку")
        self.convert_btn.clicked.connect(self.convert_pdfs)
        self.convert_btn.setEnabled(False)
        layout.addWidget(self.convert_btn)

        self.progress = QProgressBar()
        self.progress.setValue(0)
        layout.addWidget(self.progress)

        self.setLayout(layout)
        self.pdf_paths = []
        self.worker = None

    def select_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "Выберите PDF-файлы",
            "",
            "PDF Files (*.pdf)"
        )
        if files:
            self.pdf_paths = files
            self.label.setText(f"Выбрано файлов: {len(files)}")
            self.convert_btn.setEnabled(True)

    def convert_pdfs(self):
        out_dir = QFileDialog.getExistingDirectory(self, "Выберите папку для сохранения")
        if not out_dir:
            return

        self.progress.setValue(0)
        self.convert_btn.setEnabled(False)
        self.select_btn.setEnabled(False)

        self.worker = WorkerThread(self.pdf_paths, out_dir)
        self.worker.progress_update.connect(self.progress.setValue)
        self.worker.max_progress.connect(self.progress.setMaximum)
        self.worker.message.connect(self.show_message)
        self.worker.finished.connect(self.on_finished)
        self.worker.start()

    def show_message(self, text):
        QMessageBox.information(self, "Информация", text)

    def on_finished(self):
        self.convert_btn.setEnabled(True)
        self.select_btn.setEnabled(True)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PDFScannerApp()
    window.show()
    sys.exit(app.exec())