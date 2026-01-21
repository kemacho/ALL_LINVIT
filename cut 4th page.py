import os
from tkinter import Tk, filedialog
from pypdf import PdfReader, PdfWriter


def extract_page_4(pdf_path):
    reader = PdfReader(pdf_path)

    if len(reader.pages) < 4:
        print(f"❌ В файле меньше 4 страниц: {pdf_path}")
        return

    writer = PdfWriter()
    writer.add_page(reader.pages[3])  # 4-я страница (нумерация с 0)

    folder = os.path.dirname(pdf_path)
    base_name = os.path.splitext(os.path.basename(pdf_path))[0]
    new_name = f"{base_name} приложение 1.pdf"
    output_path = os.path.join(folder, new_name)

    with open(output_path, "wb") as f:
        writer.write(f)

    print(f"✅ Создан файл: {output_path}")


def main():
    root = Tk()
    root.withdraw()  # скрываем главное окно

    pdf_files = filedialog.askopenfilenames(
        title="Выберите PDF файлы",
        filetypes=[("PDF files", "*.pdf")]
    )

    for pdf in pdf_files:
        extract_page_4(pdf)


if __name__ == "__main__":
    main()