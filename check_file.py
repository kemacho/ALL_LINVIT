import os
import glob

import tkinter as tk
from tkinter import filedialog, ttk
from tkinter import messagebox
import threading

# Названия подпапок СИ
nameSI = ['0 Заявка и приложение', '1 Распоряжение по заявке', '2 Решение по заявке', '3 Заключения по ОМД и ТД',
          '4 Акт выбора ПК', '5 Протоколы СИ', '6 Заключение СИ', '7 Программа проверки произ', '8 Акт ПП',
          '9 Распоряжение на анализ', '10 Решение о выдаче', '11 Сертификат', '12 Доп.материалы']

# Названия подпапок ИК
nameIK = ['0 Распоряжение', '1 Письмо-уведомление', '2 Программа ИК', '3 Программа проверки произ', '4 Акт выбора ПК',
          '5 Акт проверки производства', '6 Протоколы ИК', '7 Акт по результатам ИК', '8 Распоряжение на анализ',
          '9 Решение по ИК', '10 Доп. материалы']

# Шаблоны
SI_TEMPLATES = {
    "Шаблон РЖД": ['2', '1', '1', '4', '1', 'Any', '1', '1', '1', '1', '2', '2', 'Any'],
    "Шаблон 2024": ['1', '1', '1', '4', '1', 'Any', '1', '1', '1', '1', '1', '1', 'Any']
}

IK_TEMPLATES = {
    "Шаблон РЖД": ['1', '1', '1', '1', '1', '1', 'Any', '1', '1', '1', 'Any'],
    "Шаблон 2024": ['1', '1', '1', '1', '1', '1', 'Any', '1', '1', '1', 'Any']
}

def create_gui():
    """Создает графический интерфейс для выбора папки и заполнения FileQNT."""
    root = tk.Tk()
    root.title("Обработчик папок")

    # --- Выбор корневой папки ---
    frame_path = ttk.Frame(root, padding=10)
    frame_path.pack(fill=tk.X)

    path_label = ttk.Label(frame_path, text="Корневая папка:")
    path_label.pack(side=tk.LEFT)

    path_entry = ttk.Entry(frame_path, width=50)
    path_entry.pack(side=tk.LEFT, expand=True, fill=tk.X)

    def browse_folder():
        """Открывает диалоговое окно выбора папки."""
        folder_selected = filedialog.askdirectory()
        path_entry.delete(0, tk.END)
        path_entry.insert(0, folder_selected)

    browse_button = ttk.Button(frame_path, text="Обзор", command=browse_folder)
    browse_button.pack(side=tk.LEFT)

    # --- Выбор шаблона FileQNT (SI) ---
    frame_template_si = ttk.Frame(root, padding=10)
    frame_template_si.pack(fill=tk.X)

    template_label_si = ttk.Label(frame_template_si, text="Шаблон СИ:")
    template_label_si.pack(side=tk.LEFT)

    template_names_si = list(SI_TEMPLATES.keys()) + ["Пользовательский"]  # Добавьте свои шаблоны
    template_var_si = tk.StringVar(value="Пользовательский")  # Установите значение по умолчанию

    template_combobox_si = ttk.Combobox(frame_template_si, textvariable=template_var_si, values=template_names_si, state="readonly")
    template_combobox_si.pack(side=tk.LEFT)

    def apply_template_si():
        """Применяет выбранный шаблон FileQNT (SI)."""
        selected_template = template_var_si.get()
        if selected_template != "Пользовательский":
            template_values = SI_TEMPLATES.get(selected_template)
            if template_values:
                for i, value in enumerate(template_values):
                    fileqnt_vars[i].set(value)

    apply_template_si_button = ttk.Button(frame_template_si, text="Применить шаблон СИ", command=apply_template_si)
    apply_template_si_button.pack(side=tk.LEFT)

    # --- Выбор шаблона FileQNT (IK) ---
    frame_template_ik = ttk.Frame(root, padding=10)
    frame_template_ik.pack(fill=tk.X)

    template_label_ik = ttk.Label(frame_template_ik, text="Шаблон ИК:")
    template_label_ik.pack(side=tk.LEFT)

    template_names_ik = list(IK_TEMPLATES.keys()) + ["Пользовательский"]  # Добавьте свои шаблоны
    template_var_ik = tk.StringVar(value="Пользовательский")  # Установите значение по умолчанию

    template_combobox_ik = ttk.Combobox(frame_template_ik, textvariable=template_var_ik, values=template_names_ik, state="readonly")
    template_combobox_ik.pack(side=tk.LEFT)

    def apply_template_ik():
        """Применяет выбранный шаблон FileQNT (IK)."""
        selected_template = template_var_ik.get()
        if selected_template != "Пользовательский":
            template_values = IK_TEMPLATES.get(selected_template)
            if template_values:
                for i, value in enumerate(template_values):
                    fileqnt_ik_vars[i].set(value)

    apply_template_ik_button = ttk.Button(frame_template_ik, text="Применить шаблон ИК", command=apply_template_ik)
    apply_template_ik_button.pack(side=tk.LEFT)

    # --- FileQNT и FileQNT_IK ---
    frame_qnt = ttk.Frame(root, padding=10)
    frame_qnt.pack(fill=tk.BOTH, expand=True)

    # FileQNT (nameSI)
    fileqnt_label = ttk.Label(frame_qnt, text="СИ:")
    fileqnt_label.grid(row=0, column=0, sticky=tk.W)

    fileqnt_vars = []  # Для хранения переменных выбора (StringVar)
    fileqnt_entries = []

    fileqnt_ik_label = ttk.Label(frame_qnt, text="ИК:")
    fileqnt_ik_label.grid(row=0, column=2, sticky=tk.W)

    fileqnt_ik_vars = []  # Для хранения переменных выбора (StringVar)
    fileqnt_ik_entries = []

    options = ["1", "2", "4", "Any"]  # Варианты значений для FileQNT и FileQNT_IK

    for i, si_name in enumerate(nameSI):
        label = ttk.Label(frame_qnt, text=f"{si_name}:")
        label.grid(row=i + 1, column=0, sticky=tk.E)

        var = tk.StringVar()  # Создаем StringVar для хранения выбранного значения
        var.set("1")
        fileqnt_vars.append(var)

        entry = ttk.Combobox(frame_qnt, textvariable=var, values=options, state="readonly") #state="readonly" что бы нельзя было ввести значение с клавиатуры
        entry.grid(row=i + 1, column=1, sticky=tk.W)
        fileqnt_entries.append(entry)

    for i in range(len(nameIK)):
        var = tk.StringVar()  # Создаем StringVar для хранения выбранного значения
        var.set("1")
        fileqnt_ik_vars.append(var)

        label = ttk.Label(frame_qnt, text=f"{nameIK[i]}:")
        label.grid(row=i + 1, column=2, sticky=tk.E)

        entry = ttk.Combobox(frame_qnt, textvariable=var, values=options, state="readonly")
        entry.grid(row=i + 1, column=3, sticky=tk.W)
        fileqnt_ik_entries.append(entry)


    # --- Прогрессбар и сообщения ---
    progress_bar = ttk.Progressbar(root, orient="horizontal", length=300, mode="determinate")
    progress_bar.pack(pady=10)

    message_label = ttk.Label(root, text="")
    message_label.pack(pady=5)

    # --- Кнопка запуска ---
    def start_processing():
        """Запускает обработку папок с выбранными параметрами."""
        in_path = path_entry.get()
        if not in_path:
            messagebox.showerror("Ошибка", "Необходимо выбрать корневую папку.")
            return

        # Получаем значения FileQNT из ComboBox'ов
        fileqnt = [var.get() for var in fileqnt_vars]
        fileqnt_ik = [var.get() for var in fileqnt_ik_vars]

        # Запускаем обработку в отдельном потоке
        threading.Thread(target=process_folders_wrapper, args=(in_path, fileqnt, fileqnt_ik, progress_bar, message_label, root)).start()

    start_button = ttk.Button(root, text="Начать обработку", command=start_processing)
    start_button.pack(pady=10)

    root.mainloop()

def process_folders_wrapper(inpath, fileqnt, fileqnt_ik, progress_bar, message_label, root):
    """Обёртка для process_folders, чтобы можно было вызывать из потока."""
    try:
        process_folders(inpath, fileqnt, fileqnt_ik, progress_bar, message_label, root)
    except Exception as e:
        messagebox.showerror("Ошибка", f"Произошла общая ошибка: {e}")
    finally:
        root.after(0, lambda: messagebox.showinfo("Готово", "Обработка завершена!")) # Вывод сообщения по завершении


def process_folders(inpath, fileqnt, fileqnt_ik, progress_bar, message_label, root):
    """Обрабатывает папки на основе заданных параметров."""

    # Названия папок
    try:
        folder_names = os.listdir(inpath)
    except FileNotFoundError:
        messagebox.showerror("Ошибка", "Указанная папка не найдена.")
        return
    except Exception as e:
        messagebox.showerror("Ошибка", f"Произошла ошибка при чтении папки: {e}")
        return

    worry = []  # непонятки

    def worrymessage(pt, nm1, nm2):
        worry.append(pt)
        worry.append(nm1)
        worry.append(nm2)

    def process_folder_with_one_file(old_folder, contents, Pos_dest, Neg_dest):
        """Обрабатывает папки, ожидающие 1 файл."""
        if len(contents) == 1 and old_folder != Pos_dest:
            os.rename(old_folder, Pos_dest)
        elif len(contents) == 0 and old_folder != Neg_dest:
            os.rename(old_folder, Neg_dest)
        elif len(contents) > 1:
            os.rename(old_folder, Pos_dest)
            worrymessage(old_folder, len(contents), '1')

    def process_folder_with_four_files(old_folder, contents, Pos_dest, Neg_dest):
        """Обрабатывает папки, ожидающие 4 файла."""
        if len(contents) == 4 and old_folder != Pos_dest:
            os.rename(old_folder, Pos_dest)
        elif len(contents) < 4 and old_folder != Neg_dest:
            os.rename(old_folder, Neg_dest)
        elif len(contents) > 4:
            os.rename(old_folder, Pos_dest)
            worrymessage(old_folder, len(contents), '4')

    def process_folder_with_any_files(old_folder, contents, Pos_dest, Neg_dest):
        """Обрабатывает папки, где ожидается любое количество файлов (больше 1)."""
        if len(contents) > 0 and old_folder != Pos_dest:
            os.rename(old_folder, Pos_dest)
        elif len(contents) == 0 and old_folder != Neg_dest:
            os.rename(old_folder, Neg_dest)
        elif len(contents) > 20:
            os.rename(old_folder, Pos_dest)
            worrymessage(old_folder, len(contents), 'возможно не так много')

    def process_folder_with_two_files(old_folder, contents, Pos_dest, Norm_dest, Neg_dest):
        """Обрабатывает папки, ожидающие 2 файла."""
        if len(contents) == 2 and old_folder != Pos_dest:
            os.rename(old_folder, Pos_dest)
        elif len(contents) == 1 and old_folder != Norm_dest:
            os.rename(old_folder, Norm_dest)
        elif len(contents) == 0 and old_folder != Neg_dest:
            os.rename(old_folder, Neg_dest)
        elif len(contents) > 2:
            os.rename(old_folder, Pos_dest)
            worrymessage(old_folder, len(contents), '2')

    def check(FileQNT, old_folder, contents, Pos_dest, Norm_dest, Neg_dest):
        # Фильтруем список contents, исключая "Thumbs.db"
        filtered_contents = [f for f in contents if f != "Thumbs.db"]

        if FileQNT == "1":
            process_folder_with_one_file(old_folder, filtered_contents, Pos_dest, Neg_dest)
        if FileQNT == "2":
            process_folder_with_two_files(old_folder, filtered_contents, Pos_dest, Norm_dest, Neg_dest)
        if FileQNT == "4":
            process_folder_with_four_files(old_folder, filtered_contents, Pos_dest, Neg_dest)
        elif FileQNT == 'Any':
            process_folder_with_any_files(old_folder, filtered_contents, Pos_dest, Neg_dest)


    # Названия подпапок ИК
    nameIK = ['0 Распоряжение', '1 Письмо-уведомление', '2 Программа ИК', '3 Программа проверки произ', '4 Акт выбора ПК',
              '5 Акт проверки производства', '6 Протоколы ИК', '7 Акт по результатам ИК', '8 Распоряжение на анализ',
              '9 Решение по ИК', '10 Доп. материалы']

    total_folders = len(folder_names)
    for i in range(0, total_folders):

        # Путь к подпапкам первого уровня
        name = folder_names[i]

        # Путь к подпапкам второго уровня
        pathSI = rf'{inpath}\{name}\0. СИ'
        pathIK1 = rf'{inpath}\{name}\1. ИК-1'
        pathIK2 = rf'{inpath}\{name}\2. ИК-2'

        # Проверка для папки СИ
        for j in range(0, len(fileqnt)):

            nameSIzv = nameSI[j] + '*'
            try:
                old_folder = glob.glob(os.path.join(pathSI, nameSIzv))
                if old_folder: # Добавляем проверку, что список не пустой
                    old_folder = old_folder[0]
                    contents = os.listdir(old_folder)

                    folder = os.path.join(pathSI, nameSI[j])
                    Pos_dest = str(folder) + ' (+)'
                    Norm_dest = str(folder) + ' (+—)'
                    Neg_dest = str(folder) + ' (—)'

                    check(fileqnt[j], old_folder, contents, Pos_dest, Norm_dest, Neg_dest)
                else:
                    print(f"Папка {nameSIzv} не найдена в {pathSI}")

            except Exception as e:
                print(f"Ошибка при обработке {nameSI[j]} в {name}: {e}")

        # Проверка для папки ИК1
        for j in range(0, len(fileqnt_ik)):

            nameIK1zv = nameIK[j] + '*'
            try:
                old_folder = glob.glob(os.path.join(pathIK1, nameIK1zv))
                if old_folder:  # Добавляем проверку, что список не пустой
                    old_folder = old_folder[0]
                    contents = os.listdir(old_folder)

                    folder = os.path.join(pathIK1, nameIK[j])
                    Pos_dest = str(folder) + ' (+)'
                    Norm_dest = str(folder) + ' (+—)'
                    Neg_dest = str(folder) + ' (—)'

                    check(fileqnt_ik[j], old_folder, contents, Pos_dest, Norm_dest, Neg_dest)
                else:
                    print(f"Папка {nameIK1zv} не найдена в {pathIK1}")

            except Exception as e:
                print(f"Ошибка при обработке {nameIK[j]} в {name} (ИК1): {e}")
        # Проверка для папки ИК2
        for j in range(0, len(fileqnt_ik)):
            nameIK2zv = nameIK[j] + '*'
            try:
                old_folder = glob.glob(os.path.join(pathIK2, nameIK2zv))
                if old_folder: # Добавляем проверку, что список не пустой
                    old_folder = old_folder[0]
                    contents = os.listdir(old_folder)

                    folder = os.path.join(pathIK2, nameIK[j])
                    Pos_dest = str(folder) + ' (+)'
                    Norm_dest = str(folder) + ' (+—)'
                    Neg_dest = str(folder) + ' (—)'

                    check(fileqnt_ik[j], old_folder, contents, Pos_dest, Norm_dest, Neg_dest)
                else:
                    print(f"Папка {nameIK2zv} не найдена в {pathIK2}")

            except Exception as e:
                print(f"Ошибка при обработке {nameIK[j]} в {name} (ИК2): {e}")

        # Обновляем прогрессбар и сообщение
        progress = int((i + 1) / total_folders * 100)
        root.after(0, lambda p=progress: progress_bar.config(value=p))
        root.after(0, lambda m=f"Обработано папок: {i + 1} из {total_folders}": message_label.config(text=m))
        root.update() #  принудительное обновление интерфейса, чтобы прогресс был виден

    #Выводим отлавливаем ошибку что бы i не вышел за размеры worry
    worry_messages = []
    for i in range(0, len(worry), 3):
        msg = f'Пожалуйста проверьте папку: {worry[i]}, там находится {worry[i + 1]} файла, вместо {worry[i + 2]}'
        worry_messages.append(msg)

    # Выводим все сообщения о беспокойствах в messagebox
    if worry_messages:
        all_worry_messages = "\n".join(worry_messages)
        root.after(0, lambda: messagebox.showinfo("Предупреждения", all_worry_messages))

if __name__ == "__main__":
    create_gui()