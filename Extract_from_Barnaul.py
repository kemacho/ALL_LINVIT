import re
import pandas as pd
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph

def extract_si_table_data(block):
    """
    Ищет таблицу с "Наименование СИ" в первой ячейке первой строки.
    Возвращает словарь с данными из столбцов 2-5 согласно новым правилам.
    """
    si_data = {
        "Строка 2": ["" ,"", "", ""],
        "Строка 3+4": ["", "", "", ""]
    }

    def before_comma(text):
        return text.split(",")[0].strip()

    for el in block:
        if isinstance(el, Table):
            first_cell_text = el.cell(0, 0).text.strip().lower()
            if "наименование си" in first_cell_text:
                rows = el.rows
                if len(rows) < 3:
                    return si_data

                def extract_columns(row):
                    cells = row.cells
                    cols = []
                    for idx in [1, 2, 3, 4]:
                        if idx < len(cells):
                            cols.append(before_comma(cells[idx].text.strip()))
                        else:
                            cols.append("")
                    return cols

                row2 = extract_columns(rows[1])

                if len(rows) == 3:
                    row3 = extract_columns(rows[2])
                    si_data["Строка 2"] = row2
                    si_data["Строка 3+4"] = row3
                elif len(rows) == 4:
                    row3 = extract_columns(rows[2])
                    row4 = extract_columns(rows[3])
                    combined = []
                    for c3, c4 in zip(row3, row4):
                        parts = []
                        if c3:
                            parts.append(c3)
                        if c4:
                            parts.append(c4)
                        combined.append(", ".join(parts))
                    si_data["Строка 2"] = row2
                    si_data["Строка 3+4"] = combined
                else:
                    # если больше 4 строк, просто берем 2 и 3,4 как выше
                    si_data["Строка 2"] = row2
                    si_data["Строка 3+4"] = ["", "", "", ""]

                return si_data
    return si_data


# === 1. Загружаем документ ===
doc = Document(r"U:\Протоколы лето 2024.docx")  # <-- замени путь

# === 2. Собираем все элементы в теле документа ===
elements = []
for block in doc.element.body:
    if block.tag.endswith("tbl"):
        elements.append(Table(block, doc))
    elif block.tag.endswith("p"):
        elements.append(Paragraph(block, doc))

# === 3. Разбиваем по блокам (начало — таблица с "Заказчик:")
blocks = []
current_block = []
is_first = True

for el in elements:
    text = ""
    if isinstance(el, Table):
        text = "\n".join(cell.text.strip() for row in el.rows for cell in row.cells)
    elif isinstance(el, Paragraph):
        text = el.text.strip()

    if "Заказчик:" in text:
        if not is_first and current_block:
            blocks.append(current_block)
            current_block = []
        is_first = False
    if not is_first:
        current_block.append(el)

if current_block:
    blocks.append(current_block)

# === 4. Обрабатываем каждый блок ===
results = []

month_map = {
    "января": "01", "февраля": "02", "марта": "03", "апреля": "04",
    "мая": "05", "июня": "06", "июля": "07", "августа": "08",
    "сентября": "09", "октября": "10", "ноября": "11", "декабря": "12"
}

for i, block in enumerate(blocks):
    print(f"🔄 Обработка блока {i + 1} из {len(blocks)}")

    protocol_value = ""
    place_value = ""
    power_value = ""
    date_value = ""
    protocol_date = ""

    block_text = ""

    for el in block:
        if isinstance(el, Table):
            table_text = "\n".join(cell.text.strip() for row in el.rows for cell in row.cells)

            # Таблица с "Заказчик:" — берём дату из последней строки, правой ячейки
            if "Заказчик:" in table_text:
                last_row = el.rows[-1]
                if len(last_row.cells) > 0:
                    raw_date = last_row.cells[-1].text.strip().lower()
                    match = re.search(r"(\d{1,2})\s+([а-яё]+)\s+(\d{4})", raw_date)
                    if match:
                        day, month_str, year = match.groups()
                        month = month_map.get(month_str)
                        if month:
                            protocol_date = f"{int(day):02d}.{month}.{year}"

            for row in el.rows:
                cells = row.cells
                # Номер протокола (в таблице 1×2)
                if len(cells) == 2:
                    left = cells[0].text.strip().upper()
                    right = cells[1].text.strip()
                    if "ПРОТОКОЛ" in left and not protocol_value:
                        protocol_value = right
                # Добавляем весь текст таблицы
                for cell in row.cells:
                    block_text += " " + cell.text.strip()

        elif isinstance(el, Paragraph):
            block_text += " " + el.text.strip()

    # Место в схеме
    place_match = re.search(r"Место \(обозначение\) в схеме:\s*(.*?)\s*Uн", block_text)
    if place_match:
        place_value = place_match.group(1).strip()

    # Центр питания
    power_match = re.search(r"Центр питания:\s*(.*?с\.ш\.\))", block_text)
    if power_match:
        power_value = power_match.group(1).strip()

    # Сроки проведения испытаний (только даты)
    date_match = re.search(
        r"Сроки проведения испытаний:.*?с\s*(\d{2}\.\d{2}\.\d{4})\s+\d{2}:\d{2}\s+по\s+(\d{2}\.\d{2}\.\d{4})",
        block_text
    )
    if date_match:
        date_value = f"{date_match.group(1)} – {date_match.group(2)}"

    # --- Вытягиваем данные из таблицы "Наименование СИ" ---
    si_data = extract_si_table_data(block)

    results.append({
        "Протокол": protocol_value,
        "Место в схеме": place_value,
        "Центр питания": power_value,
        "Сроки испытаний": date_value,
        "Дата протокола": protocol_date,

        "SI Строка 2 - 2-й столбец": si_data["Строка 2"][0],
        "SI Строка 2 - 3-й столбец": si_data["Строка 2"][1],
        "SI Строка 2 - 4-й столбец": si_data["Строка 2"][2],
        "SI Строка 2 - 5-й столбец": si_data["Строка 2"][3],

        "SI Строка 3+4 - 2-й столбец": si_data["Строка 3+4"][0],
        "SI Строка 3+4 - 3-й столбец": si_data["Строка 3+4"][1],
        "SI Строка 3+4 - 4-й столбец": si_data["Строка 3+4"][2],
        "SI Строка 3+4 - 5-й столбец": si_data["Строка 3+4"][3],
    })

# === 5. Сохраняем в Excel ===
df = pd.DataFrame(results)
df.to_excel(r"U:\test\результаты_протоколов.xlsx", index=False)

print("✅ Готово! Результаты сохранены в 'результаты_протоколов.xlsx'")