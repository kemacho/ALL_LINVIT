from pypdf import PdfReader, PdfWriter
import os

INPUT_PDF = "Итог1 — копия.pdf"        # исходный файл
OUTPUT_DIR = "output_docs"     # папка для результата

PACKET_PAGES = 9
PACKETS_COUNT = 42

# структура документов внутри пакета
DOC_STRUCTURE = [
    ("doc1", 2),
    ("doc2", 4),
    ("doc3", 2),
    ("doc4", 1),
]

os.makedirs(OUTPUT_DIR, exist_ok=True)

reader = PdfReader(INPUT_PDF)

page_index = 0  # глобальный индекс страницы (0-based)

for packet_num in range(1, PACKETS_COUNT + 1):
    for doc_name, page_count in DOC_STRUCTURE:
        writer = PdfWriter()

        for _ in range(page_count):
            writer.add_page(reader.pages[page_index])
            page_index += 1

        output_filename = f"packet_{packet_num:02d}_{doc_name}.pdf"
        output_path = os.path.join(OUTPUT_DIR, output_filename)

        with open(output_path, "wb") as f:
            writer.write(f)

print("Готово! Все документы успешно разделены.")