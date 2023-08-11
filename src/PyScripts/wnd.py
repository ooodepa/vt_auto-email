import tkinter as tk
import tkinter.messagebox
import pandas as pd
import datetime
from tkinter import ttk, simpledialog
from PyScripts.msg import *


def parse_treeview(treeview):
    data = []
    for item in treeview.get_children():
        values = treeview.item(item, 'values')
        data.append(values)

    return pd.DataFrame(data, columns=["Почта", "Дата", "Клиент"])


def value_exists(tree, value):
    for item in tree.get_children():
        if tree.item(item, 'values')[0] == value:
            return True
    return False


class Window:
    def __init__(self):
        self.root = tk.Tk()
        self.root.geometry('500x600')
        self.root.title("ooodepa@mail.ru")

        self.tree = ttk.Treeview(self.root, columns=("Mail", "Data", "Name"), show="headings")

        self.tree.heading("Mail", text="Почта")
        self.tree.heading("Data", text="Дата")
        self.tree.heading("Name", text="Клиент")

        self.tree.pack(fill=tk.BOTH, expand=True)

        # Create a vertical scrollbar and connect it to the Treeview
        scrollbar = ttk.Scrollbar(self.tree, orient="vertical", command=self.tree.yview)
        scrollbar.pack(side="right", fill="y")

        # Configure the Treeview to use the scrollbar
        self.tree.configure(yscrollcommand=scrollbar.set)

        # Bind the <Configure> event of the root window to update the column widths
        self.root.bind("<Configure>", self.update_column_widths)
        self.tree.bind("<Double-1>", self.delete_row)

        self.tree.pack(fill=tk.BOTH, expand=True)

        self.password = ""
        self.file_path = ""

        with open("pw.dat", "r") as f:
            self.password = f.read().__str__()

        self.create_menu()

    def delete_row(self, event):
        selected_item = self.tree.selection()
        if selected_item:
            self.tree.delete(selected_item[0])

    def create_menu(self):
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        file_menu = tk.Menu(menubar, tearoff=False)
        file_menu.add_command(label="Open Excel", command=self.open_excel_file)
        file_menu.add_command(label="Add Client", command=self.add_client)
        file_menu.add_command(label="Sort", command=self.sort)
        file_menu.add_separator()
        file_menu.add_command(label="Send Mail", command=self.send_mail)
        file_menu.add_command(label="Change Sender Email", command=self.change_email)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)

        menubar.add_cascade(label="File", menu=file_menu)

    def add_client(self):
        mail = simpledialog.askstring("Add client", "Mail:")
        if not mail:
            return
        if value_exists(self.tree, mail):
            tk.messagebox.showinfo("Duplicate", "This client already exist.")
            return
        name = simpledialog.askstring("Add client", "Name:")
        if not name:
            return
        if tk.messagebox.askokcancel("Add client", f"Add {name} ({mail})?"):
            self.tree.insert("", tk.END, values=(mail, "new", name))

    def sort(self):
        df = parse_treeview(self.tree)

        df['Date'] = pd.to_datetime(df['Дата'])
        df['Date'] = df['Date'].dt.strftime('%Y-%m-%d')
        df = df.sort_values(by='Date', ascending=False)
        df = df.drop('Date', axis=1)

        self.update_table(df)

    def change_email(self):
        self.root.title(simpledialog.askstring("Change Email", "New Email:"))
        self.password = simpledialog.askstring("Password Entry", "Enter your password:", show='*')

    def send_mail(self):
        receivers = self.get_emails_list()
        sender = self.root.title()
        message = get_mail(sender, simpledialog.askstring("Select Subject", "Subject:"))
        send_email(sender, receivers, self.password, message)

        self.update_dates()
        tkinter.messagebox.showinfo("Info", "Complete!")
        df_t = parse_treeview(self.tree)

        if self.file_path == "":
            self.file_path = filedialog.asksaveasfilename(
                title="Save Clients",
                initialdir=os.path.join(os.getcwd(), "Tables"),
                defaultextension=".xlsx",
                filetypes=[("Excel Files", "*.xlsx")]
            )

        if self.file_path == "":
            return

        if not os.path.exists(self.file_path):
            df_t.to_excel(self.file_path)
            return

        df_f = pd.read_excel(self.file_path)
        df_merged = pd.concat([df_f, df_t], ignore_index=True)
        df_merged.drop_duplicates(subset=["Почта"], keep="last", inplace=True)
        df_merged.dropna(inplace=True, how="all")
        with pd.ExcelWriter(self.file_path, engine="openpyxl", mode="w") as writer:
            df_merged.to_excel(writer, index=False, sheet_name="Sheet1")

    def update_dates(self):
        today = datetime.datetime.now().strftime("%Y-%m-%d")
        for item in self.tree.get_children():
            self.tree.item(item, values=(self.tree.item(item, "values")[0], today, self.tree.item(item, "values")[2]))

    def get_emails_list(self):
        emails_list = []
        for item in self.tree.get_children():
            emails_list.append(self.tree.item(item, "values")[0])
        return emails_list

    def open_excel_file(self):
        self.file_path = filedialog.askopenfilename(
            initialdir=os.path.join(os.getcwd(), "Tables"),
            filetypes=[("Excel Files", "*.xlsx;*.xls")]
        )
        if self.file_path:
            try:
                df = pd.read_excel(self.file_path)
                df_selected = df.iloc[:, :3]  # Select first 3 columns (Mail, Data, Name)
                df_selected.dropna(inplace=True, how="all")
                self.update_table(df_selected)
                self.sort()
            except Exception as e:
                print("Error reading Excel file:", e)

    def update_table(self, df):
        # Clear existing data in the Treeview
        self.tree.delete(*self.tree.get_children())

        # Insert data from the DataFrame into the Treeview
        for _, row in df.iterrows():
            self.tree.insert("", tk.END, values=(row["Почта"], row["Дата"], row["Клиент"]))

    def update_column_widths(self, event):
        window_width = event.width
        self.tree.column("Mail", width=int(window_width * 0.5))
        self.tree.column("Data", width=int(window_width * 0.2))
        self.tree.column("Name", width=int(window_width * 0.3))

    def run(self):
        self.root.mainloop()
