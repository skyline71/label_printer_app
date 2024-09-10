import tkinter as tk
import ttkbootstrap as ttk
import pandas as pd
import subprocess
import time
from tkinter import messagebox
from datetime import datetime


class LabelApp:
    def __init__(self, root):
        self.root = root
        self.df = pd.DataFrame()

        # Инициализация размеров окна
        width = 520
        height = 480
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 3

        # Инициализация главного окна
        self.root.title("Редактор этикеток Альфа-технолоджи")
        self.root.geometry(f'{width}x{height}')
        self.root.resizable(False, False)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
        self.style = ttk.Style(theme='superhero')

        # Установка интерфейса
        self.setup_ui()

        # Создание .bat файлов для закрытия и открытия Excel
        self.create_bat_files()

    def setup_ui(self):
        # Установка переменных
        self.sv = tk.StringVar()
        self.sv1 = tk.StringVar()
        self.party_entries = []
        self.QUANTITIES = []

        # Создание элементов интерфейса
        self.label7 = ttk.Label(self.root, text="Наименование\nэтикетки:")
        self.label7.place(x=10, y=10)

        # Загрузка данных из Excel файла с номенклатурой и заполнение Combobox
        self.book = pd.read_excel('workbook.xlsx')
        labels = self.book['label'].tolist()
        self.combobox = ttk.Combobox(self.root, values=labels, width=40, state="readonly")
        self.combobox.place(x=125, y=10)

        self.label2 = ttk.Label(self.root, text="Дата производства:")
        self.label2.place(x=10, y=60)

        self.current_date = datetime.today()
        self.de = ttk.DateEntry(self.root, dateformat='%d.%m.%Y', firstweekday=0, startdate=self.current_date, width=10)
        self.de.place(x=125, y=60)

        self.nowdate_btn = ttk.Button(self.root, text="Текущая дата", command=self.set_now_date,
                                      bootstyle=(ttk.SUCCESS, ttk.OUTLINE), width=14)
        self.nowdate_btn.place(x=235, y=60)

        self.label6 = ttk.Label(self.root, text="Номер поддона:")
        self.label6.place(x=10, y=110)

        self.pallet = ttk.Entry(self.root, width=15, textvariable=self.sv1)
        self.pallet.place(x=125, y=110)
        self.pallet.config(validate="key", validatecommand=(self.root.register(self.validate_pallet), '%P'))

        self.clearpallet_btn = ttk.Button(self.root, text="Очистить поле", command=self.clear_pallet,
                                          bootstyle=(ttk.WARNING, ttk.OUTLINE))
        self.clearpallet_btn.place(x=235, y=110)

        self.label3 = ttk.Label(self.root, text="Количество партий:\n(Не больше 5)")
        self.label3.place(x=10, y=150)

        self.num_parties = ttk.Spinbox(self.root, from_=1, to=5, width=5, command=self.update_party_entries)
        self.num_parties.set(1)
        self.num_parties.place(x=125, y=155)
        self.num_parties.config(validate="key", validatecommand=(self.root.register(self.validate_num_party), '%P'))

        self.parties_frame = ttk.Frame(self.root)
        self.parties_frame.place(x=10, y=200)

        self.print_btn = ttk.Button(self.root, text="Печать", command=self.all_confirm,
                                    bootstyle=(ttk.SUCCESS, ttk.OUTLINE))
        self.print_btn.place(x=10, y=410)

        self.clearall_btn = ttk.Button(self.root, text="Очистить все", command=self.clear_all,
                                       bootstyle=(ttk.DANGER, ttk.OUTLINE))
        self.clearall_btn.place(x=80, y=410)

        self.restore_btn = ttk.Button(self.root, text="Восстановить предыдущие значения", command=self.last_change,
                                      bootstyle=(ttk.WARNING, ttk.OUTLINE))
        self.restore_btn.place(x=185, y=410)

        self.progressbar = ttk.Progressbar(self.root, orient="horizontal", mode='indeterminate')

        # Обновление полей ввода для номеров партий и количества штук
        self.update_party_entries()

    def show_progressbar(self):
        self.progressbar.place(x=10, y=450, width=500)
        self.progressbar.start()

    def hide_progressbar(self):
        self.progressbar.stop()
        self.progressbar.place_forget()

    def set_now_date(self):
        # Установка текущей даты в DateEntry
        self.de.entry.delete(0, tk.END)
        self.de.entry.insert(0, self.current_date.strftime('%d.%m.%Y'))

    def validate_pallet(self, new_value):
        # Валидация для номера поддона: не более 3 знаков и только положительные числа
        if new_value.isdigit() and len(new_value) <= 3:
            return True
        elif new_value == "":  # позволяет очищать поле
            return True
        else:
            return False

    def validate_num_party(self, new_value):
        # Валидация для количества партий: не более 2 знаков и только положительные числа
        if new_value.isdigit() and len(new_value) <= 2 and int(new_value) > 0:
            return True
        elif new_value == "":  # позволяет очищать поле
            return True
        else:
            return False

    def validate_party(self, new_value):
        # Валидация для номера партии: не более 8 знаков и только положительные числа
        if new_value.isdigit() and len(new_value) <= 8:
            return True
        elif new_value == "":  # позволяет очищать поле
            return True
        else:
            return False

    def validate_quantity(self, new_value):
        # Валидация для количества штук: только положительные числа и не более 3 знаков
        if new_value.isdigit() and int(new_value) > 0 and len(new_value) <= 3:
            return True
        elif new_value == "":  # позволяет очищать поле
            return True
        else:
            return False

    def clear_label(self):
        # Очистка поля для наименования
        self.combobox.set('')

    def clear_pallet(self):
        # Очистка поля для номера поддона
        self.pallet.delete(0, tk.END)

    def clear_lot(self):
        # Очистка поля для номера партии и количества штук
        for entry1, entry2 in self.party_entries:
            entry1.delete(0, tk.END)
            entry2.delete(0, tk.END)

    def clear_all(self):
        # Очистка всех полей
        self.clear_label()
        self.clear_pallet()
        self.clear_lot()

    def last_change(self):
        # Восстановление предыдущих значений из DataFrame
        if not self.df.empty:
            # Восстановление номера поддона
            self.pallet.delete(0, tk.END)
            self.pallet.insert(0, self.df.at[0, 'pallet'])

            # Восстановление наименования этикетки
            self.combobox.set(self.df.at[0, 'labelname'])

            # Восстановление даты производства
            self.de.entry.delete(0, tk.END)
            self.de.entry.insert(0, self.df.at[0, 'date'])

            # Восстановление количества партий
            num_parties = len(self.df['party'].unique())
            self.num_parties.set(num_parties)
            self.update_party_entries()

            # Восстановление номеров партий и количества штук
            for i, (entry1, entry2) in enumerate(self.party_entries):
                entry1.delete(0, tk.END)
                entry1.insert(0, self.df.at[i, 'party'])
                entry2.delete(0, tk.END)
                entry2.insert(0, self.df.at[i, 'lotofpallet'])

    def update_party_entries(self):
        # Обновление полей ввода для номеров партий и количества штук
        try:
            num_parties = int(self.num_parties.get())
        except ValueError:
            num_parties = 0  # Значение по умолчанию, если Spinbox пустой или содержит неверное значение

        current_values = [(entry1.get(), entry2.get()) for entry1, entry2 in self.party_entries]

        for widget in self.parties_frame.winfo_children():
            widget.destroy()

        self.party_entries = []
        self.QUANTITIES = []

        for i in range(num_parties):
            row = i
            label1 = ttk.Label(self.parties_frame, text=f"Номер партии {i + 1}:")
            label1.grid(row=row, column=1, padx=5, pady=5, sticky="w")
            label2 = ttk.Label(self.parties_frame, text=f"Количество штук:")
            label2.grid(row=row, column=3, padx=5, pady=5, sticky="w")

            entry1 = ttk.Entry(self.parties_frame, validate="key",
                               validatecommand=(self.root.register(self.validate_party), '%P'))
            entry1.grid(row=row, column=2, padx=5, pady=5)
            entry2 = ttk.Entry(self.parties_frame, validate="key",
                               validatecommand=(self.root.register(self.validate_quantity), '%P'))
            entry2.grid(row=row, column=4, padx=5, pady=5)

            if i < len(current_values):
                entry1.insert(0, current_values[i][0])
                entry2.insert(0, current_values[i][1])

            self.party_entries.append((entry1, entry2))

    def all_confirm(self):
        # Проверка корректности данных и создание DataFrame
        if not self.combobox.get():
            tk.messagebox.showerror("Ошибка", "Пожалуйста, выберите наименование этикетки")
            return

        if not self.de.entry.get():
            tk.messagebox.showerror("Ошибка", "Пожалуйста, выберите дату производства")
            return

        if not self.pallet.get():
            tk.messagebox.showerror("Ошибка", "Пожалуйста, введите номер поддона")
            return

        if self.num_parties.get() == '' or int(self.num_parties.get()) > 5:
            tk.messagebox.showerror("Ошибка", "Пожалуйста, введите количество партий")
            return

        self.PARTIES = []
        self.QUANTITIES = []

        for entry1, entry2 in self.party_entries:
            if not entry1.get():
                tk.messagebox.showerror("Ошибка", "Пожалуйста, введите номер партии")
                return
            elif len(entry1.get()) != 8:
                tk.messagebox.showerror("Ошибка", "Номер партии состоит из 8 цифр")
                return
            if not entry2.get():
                tk.messagebox.showerror("Ошибка", "Пожалуйста, введите количество штук")
                return

            self.PARTIES.append(entry1.get())
            self.QUANTITIES.append(int(entry2.get()))

        time.sleep(1.5)
        self.LABELNAME = self.combobox.get()
        self.PALLET = self.pallet.get()
        self.DATE = self.de.entry.get()
        self.curr_row = self.book.loc[self.book['label'] == self.LABELNAME]
        self.show_progressbar()
        self.root.after(1500, self.create_dataframe)

    def close_file(self):
        # Закрытие файлов Excel
        subprocess.Popen(r"CloseExcel.bat")

    def open_file(self):
        # Открытие файлов Excel
        subprocess.Popen(r"StartExcel.bat")

    def create_bat_files(self):
        # Создание .bat файлов для закрытия и открытия Excel
        with open('CloseExcel.bat', 'w') as file:
            file.write("taskkill /F /IM excel.exe\ntaskkill /F /IM nicelabelprint.exe\nTIMEOUT /T 2 /NOBREAK")

    def create_dataframe(self):
        if not self.curr_row.empty:
            date_obj = datetime.strptime(self.DATE, '%d.%m.%Y')
            datecode = date_obj.strftime('%d%m%y')
            lastdate = date_obj.strftime('%y')[-1]
            base_pallet = int(self.PALLET)

            # Создаем списки для каждого столбца
            labels = []
            weights = []
            productcodes = []
            regions = []
            datecodes = []
            lastdates = []
            dates = []
            sidecodes = []
            lowercodes = []
            pallets = []
            parties = []
            quantities = []
            descriptions = []
            groupweights = []

            # Проводим вычисления для каждого столбца
            for i, (party, quantity) in enumerate(zip(self.PARTIES, self.QUANTITIES)):
                labels.append(self.curr_row['labelname'].values[0])
                weights.append(self.curr_row['weight'].values[0])
                productcodes.append(str(self.curr_row['productcode'].values[0]).zfill(12))
                regions.append('O')
                datecodes.append(datecode)
                lastdates.append(lastdate)
                dates.append(self.DATE)
                sidecodes.append(f"{datecode}{quantity}O{party}P{str(base_pallet).zfill(3)}")
                lowercodes.append(f"P9{lastdate}{party}{str(base_pallet).zfill(3)}")
                pallets.append(str(base_pallet).zfill(3))
                parties.append(party)
                quantities.append(quantity)
                descriptions.append(self.curr_row['description'].values[0])
                groupweights.append(quantity * weights[0])

            # Создаем DataFrame напрямую из списков
            data = {
                'labelname': labels,
                'weight': weights,
                'productcode': productcodes,
                'region': regions,
                'groupweight': groupweights,
                'lotofpallet': quantities,
                'datecode': datecodes,
                'lastdate': lastdates,
                'sidecode': sidecodes,
                'lowercode': lowercodes,
                'pallet': pallets,
                'date': dates,
                'party': parties,
                'description': descriptions
            }

            self.df = pd.DataFrame(data)
            self.df.to_excel('output_table.xlsx', index=False)
            self.hide_progressbar()
            self.open_file()


if __name__ == "__main__":
    root = tk.Tk()
    app = LabelApp(root)
    root.mainloop()
