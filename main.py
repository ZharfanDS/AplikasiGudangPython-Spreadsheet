import gspread
from oauth2client.service_account import ServiceAccountCredentials
import tkinter as tk
from tkinter import ttk, messagebox #, simpledialog #SIMPLEDIALOG DIGUNAKAN DI FILE FUNCTION.
from delete_data import delete_data
from update_data import update_data
from add_data import add_data
from display_data import display_data
from refresh_data import refresh_data
from update_time import update_time
from export_all_sheets import export_all_sheets
from create_widgets import create_widgets
from search_text import search_text

# Tentukan scope untuk Google Sheets dan Google Drive API
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

# Membaca path dari file creditpath.txt
with open('credspath.txt', 'r') as file:
    path_line = file.readline().strip()
    path = path_line.split('=')[1].strip()  # Mengambil bagian setelah '='

# Autentikasi menggunakan file credentials.json
creds = ServiceAccountCredentials.from_json_keyfile_name(path, scope)
client = gspread.authorize(creds)

with open('spreadsheet_url.txt', 'r') as file:
    url_line = file.readline().strip()
    url = url_line.split('=')[1].strip()
    
# Buka Google Spreadsheet berdasarkan URL
spreadsheet = client.open_by_url(url)

class SpreadsheetApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Google Sheets GUI - Project")
        self.geometry("1000x600")
        
        # Create widgets terlebih dahulu sebelum melakukan pengecekan sheet
        self.create_widgets()

        # Setelah widget diinisialisasi, baru lakukan pengecekan sheet
        self.check_all_sheets()

    def check_all_sheets(self):
        # Cek apakah semua sheet kosong
        sheets = ["Stock Barang", "MR", "Barang Keluar", "PR", "Barang Masuk", "Stock Report"]
        all_empty = True
        
        for sheet_name in sheets:
            sheet = spreadsheet.worksheet(sheet_name)
            data = sheet.get_all_records()
            if data:
                all_empty = False
                break

        if all_empty:
            # Jika semua sheet kosong, arahkan ke sheet "Stock Report" dan minta input
            messagebox.showinfo("Info", "Semua sheet kosong. Harap masukkan data di sheet 'Stock Report' terlebih dahulu.")
            # self.sheet = spreadsheet.worksheet("Stock Report")
            # self.add_data()
    
    def create_dropdown_dialog(self, column_name, options):
        """ Menampilkan dialog dropdown dengan opsi yang diberikan. """
        dialog = tk.Toplevel(self)
        dialog.title(f"Pilih {column_name}")

        tk.Label(dialog, text=f"Pilih {column_name}:").pack(padx=10, pady=10)
        
        var = tk.StringVar(dialog)
        var.set(options[0] if options else "")

        dropdown = ttk.Combobox(dialog, textvariable=var, values=options, state="readonly")
        dropdown.pack(padx=10, pady=10)

        def on_select():
            dialog.result = var.get()
            dialog.destroy()

        tk.Button(dialog, text="OK", command=on_select).pack(padx=10, pady=10)

        dialog.result = None
        dialog.wait_window()

        return dialog.result
    
    def on_sheet_select(self, sheet_name):
        self.sheet = spreadsheet.worksheet(sheet_name)
        # Tampilkan nama sheet yang dipilih di label
        self.selected_sheet_label.config(text=f"Sheet saat ini : {sheet_name}")
        self.display_data()    
    
    # Memanggil function yang ada di dalam file create_widgets.py
    def create_widgets(self):
        create_widgets(self)    # Memanggil function
    
    # Memanggil function yang ada di dalam file update_time.py
    def update_time(self):
        update_time(self)   # Memanggil function
    
    # Memanggil function yang ada di dalam file export_all_sheets.py    
    def export_all_sheets(self):
        export_all_sheets(self) # Memanggil function
    
    # Memanggil function yang ada di dalam file refresh_data.py
    def refresh_data(self):
        refresh_data(self)  # Memanggil function

    # Memanggil function yang ada di dalam file display_data.py
    def display_data(self):
        display_data(self)  # Memanggil function
    
    # Memanggil function yang ada di dalam file add_data.py
    def add_data(self):
        add_data(self)  # Memanggil function

    # Memanggil function yang ada di dalam file update_data.py
    def update_data(self):
        update_data(self)   # Memanggil function

    # Memanggil function yang ada di dalam file delete_data.py
    def delete_data(self):
        delete_data(self)   # Memanggil function

    def search_text(self):
        search_text(self)
        
if __name__ == "__main__":
    app = SpreadsheetApp()
    app.mainloop()