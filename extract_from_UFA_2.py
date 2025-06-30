import os
import re
import pandas as pd
import glob
from docx import Document
from docx.document import Document as _Document
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph

def find_third_unique_repeated_value(cells):
    """
    Принимает массив, в котором одинаковые значения идут подряд,
    и возвращает третье уникальное значение из таких повторов.
    """
    unique_values = []
    prev = None

    for val in cells:
        if val != prev:
            unique_values.append(val)
            prev = val
        if len(unique_values) == 3:
            return val
    return None  # если меньше трёх уникальных блоков подряд

def iter_block_items(parent):
    """
    Yield each paragraph and table child within a parent object in document order.
    """
    if isinstance(parent, _Document):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        raise ValueError("Unsupported parent type")

    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)


def extract_data_from_docx(file_path):
    """
    Extracts all specified data points from a single .docx file based on predefined rules.

    Args:
        file_path (str): The path to the .docx file.

    Returns:
        dict: A dictionary containing all the extracted data for one report.
    """
    try:
        doc = Document(file_path)

        # Combine all text from paragraphs and tables for easier searching
        full_text = []
        for block in iter_block_items(doc):
            if isinstance(block, Paragraph):
                full_text.append(block.text)
            elif isinstance(block, Table):
                for row in block.rows:
                    for cell in row.cells:
                        full_text.append(cell.text)

        full_text_content = "\n".join(full_text)


        # --- Initialize dictionary to store results for this file ---
        data = {"Источник файла": os.path.basename(file_path)}

        # --- Extraction Rules Implementation ---

        # Электрические сети
        match = re.search(r'в электрических сетях\s*(.*?)\n', full_text_content)
        data['Электрические сети'] = match.group(1).strip().replace('«', '').replace('»', '') if match else None

        match = re.search(
            r'ПРОТОКОЛ\s*№\s*(.*?)\s*от\s*(\d{1,2}[.\- ]\d{1,2}[.\- ]\d{4}|\d{1,2}\s+[а-я]+\s+\d{4})\s*г?\.?',
            full_text_content,
            re.IGNORECASE
        )
        if match:
            data['Номер протокола'] = match.group(1).strip()
            raw_date = match.group(2).strip()
            # Приводим дату к единому формату (если нужно)
            data['Дата протокола'] = raw_date.replace('-', '.').replace(' ', '.')
        else:
            data['Номер протокола'] = None
            data['Дата протокола'] = None

        # Центр питания
        match = re.search(r'Центр питания:\s*([^\n]+)', full_text_content)
        data['Центр питания'] = match.group(1).strip() if match else None

        # Место в схеме
        match = re.search(r'Место \(обозначение\) в схеме:\s*([^\n]+)', full_text_content)
        data['Место в схеме'] = match.group(1).strip() if match else None

        # Даты испытаний
        pattern = r'Сроки проведения испытаний\s+с\s+(\d{2}\.\d{2}\.\d{4})\s*г?\.?\s*\d{2}:\d{2}:\d{2}\s+по\s+(\d{2}\.\d{2}\.\d{4})\s*г?\.?\s*\d{2}:\d{2}:\d{2}'
        match = re.search(pattern, full_text_content)
        if match:
            data['Дата начала испытаний'] = match.group(1).strip()
            data['Дата окончания испытаний'] = match.group(2).strip()
        else:
            data['Дата начала испытаний'] = None
            data['Дата окончания испытаний'] = None

        # Время наиб
        # Шаблон для извлечения интервалов времени
        pattern = r'Интервал времени наибольших нагрузок:\s*(\d{2}:\d{2}:\d{2})\s*-\s*(\d{2}:\d{2}:\d{2})\s*(\d{2}:\d{2}:\d{2})\s*-\s*(\d{2}:\d{2}:\d{2})'
        match = re.search(pattern, full_text_content)

        if match:
            data['Первый интервал нагрузки (начало)'] = match.group(1).strip()  # 08:00:00
            data['Первый интервал нагрузки (конец)'] = match.group(2).strip()  # 12:30:00
            data['Второй интервал нагрузки (начало)'] = match.group(3).strip()  # 17:00:00
            data['Второй интервал нагрузки (конец)'] = match.group(4).strip()  # 23:30:00
        else:
            # Устанавливаем None для всех полей, если совпадение не найдено
            data.update({
                'Первый интервал нагрузки (начало)': None,
                'Первый интервал нагрузки (конец)': None,
                'Второй интервал нагрузки (начало)': None,
                'Второй интервал нагрузки (конец)': None
            })


        # --- Table-based extractions ---
        in_si_table = False
        in_app1_table = False

        for block in iter_block_items(doc):
            if isinstance(block, Paragraph):
                # Поиск по параграфам
                if 'Перечень средств измерений' in block.text or "7. Перечень средств измерений:" in block.text:
                    in_si_table = True
                elif 'Результаты измерений отклонений напряжения' in block.text:
                    in_app1_table = True
                continue

            if isinstance(block, Table):


                # 7. Перечень средств измерений
                if in_si_table:
                    for row in block.rows:
                        cells = [cell.text.strip() for cell in row.cells]
                        if len(cells) > 4:
                            if '1' in cells[0]:
                                data['Тип СИ ПКЭ'] = cells[2].replace('\n', ' ')
                                data['Заводской № ПКЭ'] = cells[3].replace('\n', ' ')
                                data['Поверка СИ ПКЭ'] = cells[4].replace('\n', ' ')
                            elif '2' in cells[0]:
                                data['Тип СИ (Прибор комбинированный)'] = cells[2].replace('\n', ' ')
                                data['Заводской № СИ (Прибор комбинированный)'] = cells[3].replace('\n', ' ')
                                data['Поверка СИ (Прибор комбинированный)'] = cells[4].replace('\n', ' ')
                    in_si_table = False

                # Приложение 1
                elif in_app1_table:
                    # Получаем все строки таблицы
                    rows = list(block.rows)

                    # Проверяем, что таблица имеет достаточное количество строк
                    if len(rows) >= 16:
                        # Извлекаем данные из конкретных ячеек
                        try:
                            data["δU(-)', %"] = rows[3].cells[2].text.strip().replace('\n', ' ')
                            data["δU(+)', %"] = rows[4].cells[2].text.strip().replace('\n', ' ')
                            data['δU(-)", %'] = rows[14].cells[2].text.strip().replace('\n',' ')
                            data['δU(+)", %'] = rows[15].cells[2].text.strip().replace('\n',' ')
                        except IndexError:
                            # На случай, если структура таблицы не соответствует ожиданиям
                            pass

                    in_app1_table = False

        return data

    except Exception as e:
        print(f"!! ERROR processing file {os.path.basename(file_path)}: {e}")
        return None


def main():
    """
    Main function to find docx files in the current directory,
    process them, and save the aggregated data to an Excel file.
    """
    folder_path = r"U:\Башкирэнерго (готово)\Омск в docx"
    docx_files = glob.glob(os.path.join(folder_path, '*.docx'))

    if not docx_files:
        print("No '.docx' files were found in the current directory.")
        return

    print(f"Found {len(docx_files)} files to process...")

    all_data = []
    for file_path in docx_files:
        print(f"-> Processing: {os.path.basename(file_path)}")
        extracted_data = extract_data_from_docx(file_path)
        if extracted_data:
            all_data.append(extracted_data)

    if not all_data:
        print("\nCould not extract data from any files.")
        return

    df = pd.DataFrame(all_data)

    column_order = [
        'Источник файла',
        'Электрические сети',
        'Номер протокола',
        'Дата протокола',
        'Центр питания',
        'Место в схеме',
        'Дата начала испытаний',
        'Дата окончания испытаний',
        'Первый интервал нагрузки (начало)',
        'Первый интервал нагрузки (конец)',
        'Второй интервал нагрузки (начало)',
        'Второй интервал нагрузки (конец)',
        'Тип СИ ПКЭ',
        'Заводской № ПКЭ',
        'Поверка СИ ПКЭ',
        'Тип СИ (Прибор комбинированный)',
        'Заводской № СИ (Прибор комбинированный)',
        'Поверка СИ (Прибор комбинированный)',
        'dU(-)',
        'dU(+)',
        "δU(-)', %",
        "δU(+)', %",
        'δU(-)", %',
        'δU(+)", %'
    ]

    df = df.reindex(columns=column_order)

    output_filename = 'U:\extracted_data.xlsx'
    try:
        df.to_excel(output_filename, index=False, engine='openpyxl')
        print(f"\n✅ Success! Data has been extracted and saved to '{output_filename}'")
    except Exception as e:
        print(f"\n❌ Error saving Excel file: {e}")

if __name__ == "__main__":
    main()
