import os
from tkinter import Tk, filedialog
import win32com.client as win32
from docx2pdf import convert

PREFIXES = ("05", "09", "10", "11")


def choose_folder():
    root = Tk()
    root.withdraw()
    return filedialog.askdirectory(title="Выберите папку")


def init_word():
    word = win32.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    word.Options.ConfirmConversions = False
    word.Options.SaveNormalPrompt = False
    return word


def convert_doc_to_docx(word, path):
    try:
        doc = word.Documents.Open(
            FileName=os.path.abspath(path),
            ReadOnly=False,
            AddToRecentFiles=False,
            Visible=False
        )

        new_path = path + "x"  # .docx
        doc.SaveAs2(new_path, FileFormat=16)  # wdFormatXMLDocument
        doc.Close(False)
        return new_path

    except Exception as e:
        print(f"❌ Ошибка конвертации DOC: {path}\n{e}")
        return None


def process_subfolder(word, folder):
    all_docs = []
    filtered_docs = []

    for file in sorted(os.listdir(folder)):
        full_path = os.path.join(folder, file)

        if not file.lower().endswith((".doc", ".docx")):
            continue

        # --- .doc → .docx ---
        if file.lower().endswith(".doc") and not file.lower().endswith(".docx"):
            converted = convert_doc_to_docx(word, full_path)
            if not converted:
                continue
            full_path = converted
            file = os.path.basename(full_path)

        all_docs.append(full_path)

        if file.startswith(PREFIXES):
            filtered_docs.append(full_path)

    # --- PDF: ВСЕ ---
    if all_docs:
        convert(
            all_docs,
            os.path.join(folder, "ALL.pdf")
        )

    # --- PDF: ОТОБРАННЫЕ ---
    if filtered_docs:
        convert(
            filtered_docs,
            os.path.join(folder, "FILTERED.pdf")
        )


def main():
    root_folder = choose_folder()
    if not root_folder:
        return

    word = init_word()

    for name in os.listdir(root_folder):
        subfolder = os.path.join(root_folder, name)
        if os.path.isdir(subfolder):
            print(f"📁 {subfolder}")
            process_subfolder(word, subfolder)

    word.Quit()
    print("✅ Готово")


if __name__ == "__main__":
    main()