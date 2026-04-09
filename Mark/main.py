import customtkinter as ctk
from tkinter import ttk, messagebox, StringVar, IntVar, filedialog, simpledialog
from collections import defaultdict
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

MONTHS = [
    "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
    "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"
]

EXPENSE_TYPES = ["Ремонт", "Электроэнергия", "Непредвиденные"]


class RentApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Учёт аренды помещений")
        self.geometry("1450x900")
        self.minsize(1450, 900)
        self.resizable(False, False)

        self.rooms = []
        self.tenants = []
        self.room_vars = []
        self.records = []
        self.current_month = StringVar(value=MONTHS[0])

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        self.top_frame = ctk.CTkFrame(self)
        self.top_frame.grid(row=0, column=0, padx=15, pady=(15, 10), sticky="ew")
        self.top_frame.grid_columnconfigure(1, weight=1)

        self.title_label = ctk.CTkLabel(
            self.top_frame,
            text="Учёт аренды помещений",
            font=ctk.CTkFont(size=22, weight="bold")
        )
        self.title_label.grid(row=0, column=0, padx=15, pady=15, sticky="w")

        self.new_file_button = ctk.CTkButton(
            self.top_frame,
            text="Создать файл",
            command=self.create_new_file
        )
        self.new_file_button.grid(row=0, column=1, padx=10, pady=15, sticky="e")

        self.save_excel_button = ctk.CTkButton(
            self.top_frame,
            text="Сохранить в Excel",
            command=self.save_to_excel
        )
        self.save_excel_button.grid(row=0, column=2, padx=10, pady=15, sticky="e")

        self.main_frame = ctk.CTkFrame(self)
        self.main_frame.grid(row=1, column=0, padx=15, pady=(0, 15), sticky="nsew")
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(2, weight=1)

        self.input_frame = ctk.CTkFrame(self.main_frame)
        self.input_frame.grid(row=0, column=0, padx=15, pady=15, sticky="ew")
        self.input_frame.grid_columnconfigure(1, weight=1)
        self.input_frame.grid_columnconfigure(3, weight=1)
        self.input_frame.grid_columnconfigure(5, weight=1)

        self.form_title = ctk.CTkLabel(
            self.input_frame,
            text="Сначала создайте файл и укажите помещения и арендаторов",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        self.form_title.grid(row=0, column=0, columnspan=6, padx=10, pady=(10, 15), sticky="w")

        self.rooms_value_label = ctk.CTkLabel(self.input_frame, text="Помещения: не заданы")
        self.rooms_value_label.grid(row=1, column=0, columnspan=6, padx=10, pady=5, sticky="w")

        self.tenants_value_label = ctk.CTkLabel(self.input_frame, text="Арендаторы: не заданы")
        self.tenants_value_label.grid(row=2, column=0, columnspan=6, padx=10, pady=(0, 10), sticky="w")

        self.month_label = ctk.CTkLabel(self.input_frame, text="Месяц:")
        self.month_label.grid(row=3, column=0, padx=10, pady=10, sticky="w")
        self.month_menu = ctk.CTkOptionMenu(self.input_frame, values=MONTHS, variable=self.current_month)
        self.month_menu.grid(row=3, column=1, padx=10, pady=10, sticky="ew")

        self.tenant_label = ctk.CTkLabel(self.input_frame, text="Арендатор:")
        self.tenant_label.grid(row=3, column=2, padx=10, pady=10, sticky="w")
        self.tenant_menu = ctk.CTkOptionMenu(self.input_frame, values=["-"])
        self.tenant_menu.grid(row=3, column=3, padx=10, pady=10, sticky="ew")

        self.add_tenant_to_list_button = ctk.CTkButton(
            self.input_frame,
            text="Добавить арендатора в список",
            command=self.add_tenant_to_list
        )
        self.add_tenant_to_list_button.grid(row=3, column=4, columnspan=2, padx=10, pady=10, sticky="ew")

        self.rent_label = ctk.CTkLabel(self.input_frame, text="Сумма аренды:")
        self.rent_label.grid(row=4, column=0, padx=10, pady=10, sticky="w")
        self.rent_entry = ctk.CTkEntry(self.input_frame, placeholder_text="например 30000")
        self.rent_entry.grid(row=4, column=1, padx=10, pady=10, sticky="ew")

        self.rooms_pick_label = ctk.CTkLabel(self.input_frame, text="Помещения арендатора:")
        self.rooms_pick_label.grid(row=4, column=2, padx=10, pady=10, sticky="w")

        self.rooms_check_frame = ctk.CTkFrame(self.input_frame)
        self.rooms_check_frame.grid(row=5, column=0, columnspan=6, padx=10, pady=(0, 10), sticky="ew")

        self.add_button = ctk.CTkButton(self.input_frame, text="Добавить запись аренды", command=self.add_tenant)
        self.add_button.grid(row=6, column=0, columnspan=6, padx=10, pady=(10, 10), sticky="ew")

        self.expenses_frame = ctk.CTkFrame(self.main_frame)
        self.expenses_frame.grid(row=1, column=0, padx=15, pady=(0, 10), sticky="ew")
        self.expenses_frame.grid_columnconfigure(1, weight=1)
        self.expenses_frame.grid_columnconfigure(3, weight=1)

        self.expenses_title = ctk.CTkLabel(
            self.expenses_frame,
            text="Расходы по помещениям",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        self.expenses_title.grid(row=0, column=0, columnspan=4, padx=10, pady=(10, 15), sticky="w")

        self.expense_room_label = ctk.CTkLabel(self.expenses_frame, text="Помещение:")
        self.expense_room_label.grid(row=1, column=0, padx=10, pady=10, sticky="w")
        self.expense_room_menu = ctk.CTkOptionMenu(self.expenses_frame, values=["-"])
        self.expense_room_menu.grid(row=1, column=1, padx=10, pady=10, sticky="ew")

        self.expense_type_label = ctk.CTkLabel(self.expenses_frame, text="Тип расхода:")
        self.expense_type_label.grid(row=1, column=2, padx=10, pady=10, sticky="w")
        self.expense_type_menu = ctk.CTkOptionMenu(self.expenses_frame, values=EXPENSE_TYPES)
        self.expense_type_menu.grid(row=1, column=3, padx=10, pady=10, sticky="ew")

        self.expense_value_label = ctk.CTkLabel(self.expenses_frame, text="Сумма:")
        self.expense_value_label.grid(row=2, column=0, padx=10, pady=10, sticky="w")
        self.expense_value_entry = ctk.CTkEntry(self.expenses_frame, placeholder_text="0.00")
        self.expense_value_entry.grid(row=2, column=1, padx=10, pady=10, sticky="ew")

        self.add_expense_button = ctk.CTkButton(self.expenses_frame, text="Добавить расход", command=self.add_expense)
        self.add_expense_button.grid(row=2, column=2, columnspan=2, padx=10, pady=10, sticky="ew")

        self.table_frame = ctk.CTkFrame(self.main_frame)
        self.table_frame.grid(row=2, column=0, padx=15, pady=(0, 10), sticky="nsew")
        self.table_frame.grid_columnconfigure(0, weight=1)
        self.table_frame.grid_rowconfigure(0, weight=1)

        columns = (
            "month", "tenant", "rooms", "rent", "share",
            "repair_room", "electricity_room", "unexpected_room",
            "repair_total", "electricity_total", "unexpected_total",
            "expenses", "net"
        )
        self.table = ttk.Treeview(self.table_frame, columns=columns, show="headings", height=14)

        headings = {
            "month": "Месяц",
            "tenant": "Арендатор",
            "rooms": "Помещения",
            "rent": "Аренда",
            "share": "Доля за помещение",
            "repair_room": "Ремонт по помещениям",
            "electricity_room": "Электроэнергия по помещениям",
            "unexpected_room": "Непредвиденные по помещениям",
            "repair_total": "Ремонт всего",
            "electricity_total": "Электроэнергия всего",
            "unexpected_total": "Непредвиденные всего",
            "expenses": "Всего расходов",
            "net": "Чистый доход",
        }

        for col in columns:
            self.table.heading(col, text=headings[col])
            self.table.column(col, anchor="center", width=150, stretch=False)

        widths = {
            "month": 100,
            "tenant": 170,
            "rooms": 220,
            "rent": 100,
            "share": 120,
            "repair_room": 170,
            "electricity_room": 190,
            "unexpected_room": 190,
            "repair_total": 110,
            "electricity_total": 130,
            "unexpected_total": 130,
            "expenses": 120,
            "net": 120,
        }
        for col, width in widths.items():
            self.table.column(col, width=width, stretch=False)

        vsb = ttk.Scrollbar(self.table_frame, orient="vertical", command=self.table.yview)
        hsb = ttk.Scrollbar(self.table_frame, orient="horizontal", command=self.table.xview)
        self.table.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.table.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        self.summary_frame = ctk.CTkFrame(self.main_frame)
        self.summary_frame.grid(row=3, column=0, padx=15, pady=(0, 15), sticky="ew")
        self.summary_frame.grid_columnconfigure(1, weight=1)

        self.summary_title = ctk.CTkLabel(self.summary_frame, text="Итоги", font=ctk.CTkFont(size=16, weight="bold"))
        self.summary_title.grid(row=0, column=0, columnspan=2, padx=10, pady=(10, 10), sticky="w")

        self.summary_label = ctk.CTkLabel(self.summary_frame, text="Данные пока не рассчитаны", justify="left")
        self.summary_label.grid(row=1, column=0, columnspan=2, padx=10, pady=(0, 12), sticky="w")

    def update_tenant_menu(self):
        self.tenant_menu.configure(values=self.tenants if self.tenants else ["-"])
        self.tenant_menu.set(self.tenants[0] if self.tenants else "-")
        self.tenants_value_label.configure(
            text="Арендаторы: " + ", ".join(self.tenants) if self.tenants else "Арендаторы: не заданы"
        )

    def add_tenant_to_list(self):
        name = simpledialog.askstring("Новый арендатор", "Введите имя арендатора:")
        if not name or not name.strip():
            return
        name = name.strip()
        if name not in self.tenants:
            self.tenants.append(name)
            self.update_tenant_menu()
        self.tenant_menu.set(name)

    def create_new_file(self):
        count_text = simpledialog.askstring("Новый файл", "Сколько помещений в объекте?")
        if not count_text:
            return

        try:
            count = int(count_text)
            if count <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Ошибка", "Введите корректное число помещений.")
            return

        rooms = []
        for i in range(count):
            room_name = simpledialog.askstring("Названия помещений", f"Введите название помещения №{i + 1}:")
            if room_name is None or not room_name.strip():
                messagebox.showwarning("Внимание", "Название помещения не введено.")
                return
            rooms.append(room_name.strip())

        tenants_text = simpledialog.askstring("Арендаторы", "Введите имена арендаторов через запятую:")
        tenants = []
        if tenants_text:
            tenants = [t.strip() for t in tenants_text.split(",") if t.strip()]

        self.rooms = rooms
        self.tenants = tenants

        self.rooms_value_label.configure(text="Помещения: " + ", ".join(self.rooms))
        self.update_tenant_menu()

        self.form_title.configure(text="Файл создан. Можно добавлять арендаторов и расходы.")

        for widget in self.rooms_check_frame.winfo_children():
            widget.destroy()

        self.room_vars = []
        for i, room in enumerate(self.rooms):
            var = IntVar(value=0)
            cb = ctk.CTkCheckBox(self.rooms_check_frame, text=room, variable=var)
            cb.grid(row=i // 4, column=i % 4, padx=10, pady=8, sticky="w")
            self.room_vars.append(var)

        self.expense_room_menu.configure(values=self.rooms if self.rooms else ["-"])
        self.expense_room_menu.set(self.rooms[0] if self.rooms else "-")

    def add_tenant(self):
        if not self.rooms:
            messagebox.showwarning("Внимание", "Сначала создайте файл и укажите помещения.")
            return

        tenant = self.tenant_menu.get().strip()
        rent_text = self.rent_entry.get().strip()
        selected_rooms = [room for room, var in zip(self.rooms, self.room_vars) if var.get() == 1]

        if not tenant or tenant == "-" or not rent_text or not selected_rooms:
            messagebox.showerror("Ошибка", "Выберите арендатора, сумму и помещения.")
            return

        try:
            rent = float(rent_text)
        except ValueError:
            messagebox.showerror("Ошибка", "Введите корректную сумму аренды.")
            return

        share = rent / len(selected_rooms)

        record = {
            "month": self.current_month.get(),
            "tenant": tenant,
            "rooms": selected_rooms,
            "rent": rent,
            "share": share,
            "expenses": defaultdict(lambda: {t: 0.0 for t in EXPENSE_TYPES})
        }
        self.records.append(record)
        self.refresh_table()
        self.clear_tenant_form()

    def add_expense(self):
        room = self.expense_room_menu.get()
        expense_type = self.expense_type_menu.get()
        value_text = self.expense_value_entry.get().strip()

        if not room or room == "-" or not value_text:
            messagebox.showerror("Ошибка", "Выберите помещение и введите сумму.")
            return

        try:
            value = float(value_text)
        except ValueError:
            messagebox.showerror("Ошибка", "Введите корректную сумму расхода.")
            return

        for record in self.records:
            if room in record["rooms"]:
                record["expenses"][room][expense_type] += value

        self.expense_value_entry.delete(0, "end")
        self.refresh_table()

    def refresh_table(self):
        for item in self.table.get_children():
            self.table.delete(item)

        total_rent = 0
        total_expenses = 0
        total_rooms = len(self.rooms)

        for record in self.records:
            repair_room = []
            electricity_room = []
            unexpected_room = []

            for room in record["rooms"]:
                r = record["expenses"][room]["Ремонт"]
                e = record["expenses"][room]["Электроэнергия"]
                u = record["expenses"][room]["Непредвиденные"]

                repair_room.append(f"{room}: {r:.2f}")
                electricity_room.append(f"{room}: {e:.2f}")
                unexpected_room.append(f"{room}: {u:.2f}")

            repair_total = sum(record["expenses"][room]["Ремонт"] for room in record["rooms"])
            electricity_total = sum(record["expenses"][room]["Электроэнергия"] for room in record["rooms"])
            unexpected_total = sum(record["expenses"][room]["Непредвиденные"] for room in record["rooms"])

            room_expense_sum = repair_total + electricity_total + unexpected_total
            net = record["rent"] - room_expense_sum
            total_rent += record["rent"]
            total_expenses += room_expense_sum

            self.table.insert(
                "",
                "end",
                values=(
                    record["month"],
                    record["tenant"],
                    ", ".join(record["rooms"]),
                    f"{record['rent']:.2f}",
                    f"{record['share']:.2f}",
                    "\n".join(repair_room) if repair_room else "-",
                    "\n".join(electricity_room) if electricity_room else "-",
                    "\n".join(unexpected_room) if unexpected_room else "-",
                    f"{repair_total:.2f}",
                    f"{electricity_total:.2f}",
                    f"{unexpected_total:.2f}",
                    f"{room_expense_sum:.2f}",
                    f"{net:.2f}"
                )
            )

        total_income_without_expenses = total_rent
        total_income_with_expenses = total_rent - total_expenses
        avg_without_expenses = total_income_without_expenses / total_rooms if total_rooms else 0
        avg_with_expenses = total_income_with_expenses / total_rooms if total_rooms else 0

        self.summary_label.configure(
            text=(
                f"Общий доход без расходов: {total_income_without_expenses:.2f}\n"
                f"Общий доход с расходами: {total_income_with_expenses:.2f}\n"
                f"Средний доход без расходов: {avg_without_expenses:.2f}\n"
                f"Средний доход с расходами: {avg_with_expenses:.2f}\n"
                f"Общее количество помещений: {total_rooms}"
            ),
            justify="left"
        )

    def clear_tenant_form(self):
        self.rent_entry.delete(0, "end")
        for var in self.room_vars:
            var.set(0)

    def save_to_excel(self):
        if not self.records:
            messagebox.showwarning("Внимание", "Нет данных для сохранения.")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="Сохранить как"
        )
        if not file_path:
            return

        wb = Workbook()
        ws = wb.active
        ws.title = "Аренда"

        header_fill = PatternFill("solid", fgColor="1F4E78")
        header_font = Font(color="FFFFFF", bold=True)
        bold_font = Font(bold=True)
        center = Alignment(horizontal="center", vertical="center")
        left = Alignment(horizontal="left", vertical="center")
        thin = Side(style="thin", color="000000")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        headers = ["Месяц", "Арендатор", "Помещения", "Аренда", "Доля за помещение", "Ремонт по помещениям", "Электроэнергия по помещениям", "Непредвиденные по помещениям", "Ремонт всего", "Электроэнергия всего", "Непредвиденные всего", "Всего расходов", "Чистый доход"]
        row = 1

        months_order = []
        for rec in self.records:
            if rec["month"] not in months_order:
                months_order.append(rec["month"])

        total_rent = 0
        total_expenses = 0

        for month in months_order:
            ws.cell(row=row, column=1, value=month)
            ws.cell(row=row, column=1).font = Font(bold=True, size=14)
            row += 1

            for col, head in enumerate(headers, start=1):
                cell = ws.cell(row=row, column=col, value=head)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = center
                cell.border = border
            row += 1

            for rec in self.records:
                if rec["month"] != month:
                    continue

                repair_room = []
                electricity_room = []
                unexpected_room = []

                for room in rec["rooms"]:
                    repair_room.append(f"{room}: {rec['expenses'][room]['Ремонт']:.2f}")
                    electricity_room.append(f"{room}: {rec['expenses'][room]['Электроэнергия']:.2f}")
                    unexpected_room.append(f"{room}: {rec['expenses'][room]['Непредвиденные']:.2f}")

                repair_total = sum(rec["expenses"][room]["Ремонт"] for room in rec["rooms"])
                electricity_total = sum(rec["expenses"][room]["Электроэнергия"] for room in rec["rooms"])
                unexpected_total = sum(rec["expenses"][room]["Непредвиденные"] for room in rec["rooms"])

                room_expense_sum = repair_total + electricity_total + unexpected_total
                net = rec["rent"] - room_expense_sum

                total_rent += rec["rent"]
                total_expenses += room_expense_sum

                values = [
                    rec["month"],
                    rec["tenant"],
                    ", ".join(rec["rooms"]),
                    rec["rent"],
                    rec["share"],
                    "\n".join(repair_room) if repair_room else "-",
                    "\n".join(electricity_room) if electricity_room else "-",
                    "\n".join(unexpected_room) if unexpected_room else "-",
                    repair_total,
                    electricity_total,
                    unexpected_total,
                    room_expense_sum,
                    net
                ]

                for col, value in enumerate(values, start=1):
                    cell = ws.cell(row=row, column=col, value=value)
                    cell.border = border
                    cell.alignment = left if col in (2, 3, 6, 7, 8) else center
                row += 1

            row += 1

        total_rooms = len(self.rooms)
        total_income_without_expenses = total_rent
        total_income_with_expenses = total_rent - total_expenses
        avg_without_expenses = total_income_without_expenses / total_rooms if total_rooms else 0
        avg_with_expenses = total_income_with_expenses / total_rooms if total_rooms else 0

        summary_items = [
            ("Общий доход без расходов", total_income_without_expenses),
            ("Общий доход с расходами", total_income_with_expenses),
            ("Средний доход без расходов", avg_without_expenses),
            ("Средний доход с расходами", avg_with_expenses),
            ("Общее количество помещений", total_rooms),
        ]

        ws.cell(row=row, column=1, value="Итоги")
        ws.cell(row=row, column=1).font = Font(bold=True, size=14)
        row += 1

        for name, value in summary_items:
            ws.cell(row=row, column=1, value=name)
            ws.cell(row=row, column=2, value=value)
            ws.cell(row=row, column=1).font = bold_font
            row += 1

        widths = {
            1: 18, 2: 28, 3: 35, 4: 15, 5: 18,
            6: 26, 7: 30, 8: 30,
            9: 15, 10: 18, 11: 18, 12: 15, 13: 18
        }
        for col_idx, width in widths.items():
            ws.column_dimensions[chr(64 + col_idx)].width = width

        wb.save(file_path)
        messagebox.showinfo("Готово", f"Файл сохранён:\n{file_path}")


if __name__ == "__main__":
    app = RentApp()
    app.mainloop()