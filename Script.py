import tkinter as tk
import pandas as pd
from tkinter import ttk, simpledialog, messagebox, filedialog
import sqlite3
import csv

class Main(tk.Frame):
    def __init__(self, root):
        super().__init__(root)
        self.init_main()
        self.db = Database()
        self.view_records()

    def init_main(self):
        self.style = ttk.Style()
        self.style.configure('TButton', background='#FFC141')
        self.style.configure('TFrame', background='#0f0f10')
        self.style.configure('TLabel', background='#0f0f10', foreground='#FFC141')
        self.style.configure('Treeview', background='#0f0f10', fieldbackground='#0f0f10', foreground='#FFC141')

        self.grid(row=0, column=0, sticky='nsew')
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self.tree = ttk.Treeview(self, columns=['ID', 'name', 'tel', 'email', 'salary'], height=15, show='headings')
        self.tree.column('ID', width=30, anchor=tk.CENTER)
        self.tree.column('name', width=150, anchor=tk.CENTER)
        self.tree.column('tel', width=150, anchor=tk.CENTER)
        self.tree.column('email', width=150, anchor=tk.CENTER)
        self.tree.column('salary', width=100, anchor=tk.CENTER)
        self.tree.heading('ID', text='ID')
        self.tree.heading('name', text='ФИО')
        self.tree.heading('tel', text='Телефон')
        self.tree.heading('email', text='E-mail')
        self.tree.heading('salary', text='Заработная плата')

        self.tree.bind("<Double-1>", self.on_item_double_click)
        self.tree.bind("<Button-3>", self.show_context_menu)
        self.tree.bind("<Control-a>", self.select_all_records)
        self.tree.bind("<Delete>", self.delete_record)
        self.tree.bind("<Control-f>", self.search_by_name)
        self.tree.bind("<F5>", self.refresh_page)


        self.tree.grid(row=0, column=0, padx=5, pady=5, sticky='nsew')

        menu = tk.Menu(root)
        root.config(menu=menu)
        file_menu = tk.Menu(menu, tearoff=False)
        menu.add_cascade(label="Меню", menu=file_menu)
        file_menu.add_command(label="Экспорт в CSV", command=self.export_to_csv)
        file_menu.add_command(label="Импорт из CSV", command=self.import_from_csv)
        file_menu.add_separator()
        file_menu.add_command(label="Выход", command=root.destroy)

        menu.add_cascade(label="Поиск", command=self.search_by_name)

        add_contact_frame = ttk.Frame(self, style='TFrame')
        add_contact_frame.grid(row=0, column=1, padx=5, pady=5, sticky='nsew')

        label_name = ttk.Label(add_contact_frame, text='ФИО')
        label_name.grid(row=0, column=0, padx=5, pady=5, sticky='w')
        label_tel = ttk.Label(add_contact_frame, text='Телефон')
        label_tel.grid(row=1, column=0, padx=5, pady=5, sticky='w')
        label_email = ttk.Label(add_contact_frame, text='E-mail')
        label_email.grid(row=2, column=0, padx=5, pady=5, sticky='w')
        label_salary = ttk.Label(add_contact_frame, text='Заработная плата')
        label_salary.grid(row=3, column=0, padx=5, pady=5, sticky='w')

        self.entry_name = ttk.Entry(add_contact_frame)
        self.entry_name.grid(row=0, column=1, padx=5, pady=5, sticky='w')
        self.entry_tel = ttk.Entry(add_contact_frame)
        self.entry_tel.grid(row=1, column=1, padx=5, pady=5, sticky='w')
        self.entry_email = ttk.Entry(add_contact_frame)
        self.entry_email.grid(row=2, column=1, padx=5, pady=5, sticky='w')
        self.entry_salary = ttk.Entry(add_contact_frame)
        self.entry_salary.grid(row=3, column=1, padx=5, pady=5, sticky='w')

        self.button_add = ttk.Button(add_contact_frame, text='Добавить', command=self.add_record)
        self.button_add.grid(row=4, column=0, columnspan=2, padx=5, pady=5, sticky='w')

    def on_item_double_click(self, event):
        item = self.tree.selection()[0]
        data = self.tree.item(item, 'values')
        if data:
            EditDialog(self, self.db, data)

    def show_context_menu(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            context_menu = tk.Menu(self, tearoff=0)
            context_menu.add_command(label="Позвонить")
            context_menu.add_command(label="Изменить", command=self.edit_record)
            context_menu.add_command(label="Удалить", command=self.delete_record)
            context_menu.post(event.x_root, event.y_root)

    def view_records(self):
        self.tree.delete(*self.tree.get_children())
        data = self.db.get_all_data()
        for row in data:
            self.tree.insert('', 'end', values=row)

    def refresh_page(self, event):
        self.view_records()

    def export_to_csv(self):
        self.db.export_to_csv()

    def import_from_csv(self):
        file_path = tk.filedialog.askopenfilename(defaultextension=".csv", filetypes=[("CSV Files", "*.csv")])

        if file_path:
            db = Database()
            db.import_data_from_csv(file_path)
            self.view_records()

    def edit_record(self):
        selected_item = self.tree.selection()[0]
        data = self.tree.item(selected_item, 'values')
        if data:
            EditDialog(self, self.db, data)

    def delete_record(self, event=None):
        selected_item = self.tree.selection()[0]
        data = self.tree.item(selected_item, 'values')
        if data:
            self.db.delete_data(data[0])
            self.view_records()

    def select_all_records(self, event):
        self.tree.selection_set(self.tree.get_children())

    def add_record(self):
        name = self.entry_name.get()
        tel = self.entry_tel.get()
        email = self.entry_email.get()
        salary = self.entry_salary.get()
        if not name or not tel or not email:
            error_message = "Пожалуйста, заполните все поля."
            tk.messagebox.showerror("Ошибка", error_message)
            return

        db = Database()
        db.insert_data(name, tel, email, salary)
        self.view_records()
        self.entry_name.delete(0, 'end')
        self.entry_tel.delete(0, 'end')
        self.entry_email.delete(0, 'end')
        self.entry_salary.delete(0, 'end')

    def search_by_name(self, event=None):
        search_query = simpledialog.askstring("Поиск по ФИО", "Введите ФИО:")
        if search_query:
            results = self.db.search_by_name(search_query)
            self.tree.delete(*self.tree.get_children())
            for row in results:
                self.tree.insert('', 'end', values=row)



class EditDialog(tk.Toplevel):
    def __init__(self, parent, db, data):
        super().__init__(parent)
        self.parent = parent
        self.db = db
        self.data = data
        self.init_edit_dialog()

    def init_edit_dialog(self):
        self.title('Редактировать')
        self.geometry('400x250')
        self.resizable(False, False)
        self.grab_set()
        self.focus_set()

        label_name = ttk.Label(self, text='ФИО')
        label_name.grid(row=0, column=0, padx=5, pady=5, sticky='w')
        label_tel = ttk.Label(self, text='Телефон')
        label_tel.grid(row=1, column=0, padx=5, pady=5, sticky='w')
        label_email = ttk.Label(self, text='E-mail')
        label_email.grid(row=2, column=0, padx=5, pady=5, sticky='w')
        label_salary = ttk.Label(self, text='Заработная плата')
        label_salary.grid(row=3, column=0, padx=5, pady=5, sticky='w')

        self.entry_name = ttk.Entry(self)
        self.entry_name.grid(row=0, column=1, padx=5, pady=5, sticky='w')
        self.entry_tel = ttk.Entry(self)
        self.entry_tel.grid(row=1, column=1, padx=5, pady=5, sticky='w')
        self.entry_email = ttk.Entry(self)
        self.entry_email.grid(row=2, column=1, padx=5, pady=5, sticky='w')
        self.entry_salary = ttk.Entry(self)
        self.entry_salary.grid(row=3, column=1, padx=5, pady=5, sticky='w')

        self.entry_name.insert(0, self.data[1])
        self.entry_tel.insert(0, self.data[2])
        self.entry_email.insert(0, self.data[3])
        self.entry_salary.insert(0, self.data[4])

        self.button_cancel = ttk.Button(self, text='Закрыть', command=self.destroy)
        self.button_cancel.grid(row=4, column=1, padx=5, pady=5, sticky='e')
        self.button_edit = ttk.Button(self, text='Сохранить', command=self.edit_record)
        self.button_edit.grid(row=4, column=0, padx=5, pady=5, sticky='w')

    def edit_record(self):
        name = self.entry_name.get()
        tel = self.entry_tel.get()
        email = self.entry_email.get()
        salary = self.entry_salary.get()
        if name and tel and email:
            self.db.edit_data(self.data[0], name, tel, email, salary)
            self.parent.view_records()
            self.destroy()


class Database:
    def __init__(self):
        self.conn = sqlite3.connect('db.db')
        self.c = self.conn.cursor()
        self.c.execute('''
            CREATE TABLE IF NOT EXISTS db(id INTEGER PRIMARY KEY, name TEXT, tel TEXT, email TEXT, salary TEXT)''')
        self.conn.commit()

    def insert_data(self, name, tel, email, salary):
        self.c.execute('''
            INSERT INTO db (name, tel, email, salary) VALUES (?, ?, ?, ?)''', (name, tel, email, salary))
        self.conn.commit()

    def get_all_data(self):
        self.c.execute('''SELECT * FROM db''')
        return self.c.fetchall()

    def edit_data(self, id, name, tel, email, salary):
        self.c.execute('''
            UPDATE db SET name = ?, tel = ?, email = ?, salary = ? WHERE id = ?''', (name, tel, email, salary, id))
        self.conn.commit()

    def delete_data(self, id):
        self.c.execute('DELETE FROM db WHERE id = ?', (id,))
        self.conn.commit()

    def export_to_csv(self):
        try:
            data = self.get_all_data()
            df = pd.DataFrame(data, columns=['ID', 'ФИО', 'Телефон', 'E-mail', 'Заработная плата'])
            file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV Files", "*.csv")])
            if file_path:
                df.to_csv(file_path, index=False, encoding='utf-8-sig', sep=';')
                messagebox.showinfo("Успешно", "Экспорт в CSV успешен.")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при экспорте в CSV: {e}")

    def import_data_from_csv(self, file_path):
        try:
            with open(file_path, newline='', encoding='utf-8') as csvfile:
                csvreader = csv.reader(csvfile, delimiter=';')
                next(csvreader)
                for row in csvreader:
                    name, tel, email, salary = row[1], row[2], row[3], row[4]
                    self.insert_data(name, tel, email, salary)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Произошла ошибка при импорте данных из CSV: {str(e)}")

    def search_by_name(self, query):
        query = f"%{query}%"
        self.c.execute('''
            SELECT * FROM db WHERE name LIKE ?''', (query,))
        return self.c.fetchall()

if __name__ == '__main__':
    root = tk.Tk()
    root.resizable(False, False)
    app = Main(root)
    app.pack()
    root.title('Список сотрудников компании')
    root.geometry('800x450')
    root.configure(bg='#0f0f10')
    root.mainloop()
