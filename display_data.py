import tkinter as tk

def display_data(self):
    data = self.sheet.get_all_records()
    self.text_area.config(state='normal')  # Izinkan perubahan teks
    self.text_area.delete(1.0, tk.END)

    if data:
        # Tampilkan data jika tersedia
        columns = data[0].keys()
        column_widths = {col: len(col) for col in columns}

        for row in data:
            for col in columns:
                column_widths[col] = max(column_widths[col], len(str(row[col])))

        # Menampilkan header
        header = "No | " + " | ".join(f"{col:<{column_widths[col]}}" for col in columns)

        # Membuat separator yang tepat
        no_separator = "---+-"  # Separator untuk kolom No
        column_separators = "-+-".join("-" * column_widths[col] for col in columns)
        separator = no_separator + column_separators  # Menggabungkan separator

        self.text_area.insert(tk.END, header + "\n" + separator + "\n")

        # Menampilkan data dengan nomor baris
        for idx, row in enumerate(data, start=1):
            row_values = [f"{row[col]:<{column_widths[col]}}" for col in columns]
            self.text_area.insert(tk.END, f"{idx:<2} | " + " | ".join(row_values) + "\n")
    else:
        # Jika tidak ada data, tampilkan pesan
        self.text_area.insert(tk.END, "Data Not Found.")
    self.text_area.config(state='disabled')