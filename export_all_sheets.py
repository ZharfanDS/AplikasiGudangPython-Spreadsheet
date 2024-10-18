import pandas as pd
from tkinter import filedialog, messagebox

def export_all_sheets(self):
    from main import spreadsheet
    # Tanyakan kepada pengguna tempat untuk menyimpan file
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")])

    if not file_path:
        return  # Jika pengguna membatalkan pemilihan file, akhiri fungsi

    # Daftar semua nama sheet
    sheets = ["Stock Barang", "MR", "Barang Keluar", "PR", "Barang Masuk", "Stock Report"]

    # Dictionary untuk menyimpan DataFrame dari setiap sheet
    data_frames = {}

    try:
        # Looping melalui setiap sheet dan ambil datanya
        for sheet_name in sheets:
            sheet = spreadsheet.worksheet(sheet_name)
            data = sheet.get_all_records()

            if not data:
                continue  # Jika sheet kosong, lewati sheet ini

            # Convert data menjadi DataFrame pandas dan simpan ke dictionary
            df = pd.DataFrame(data)
            data_frames[sheet_name] = df

        if not data_frames:
            messagebox.showinfo("Info", "Semua sheet kosong, tidak ada data untuk diekspor.")
            return

        # Cek tipe file yang dipilih pengguna
        if file_path.endswith(".xlsx"):
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                # Simpan setiap DataFrame ke sheet yang berbeda di Excel
                for sheet_name, df in data_frames.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
        elif file_path.endswith(".csv"):
            # Jika pengguna memilih CSV, ekspor satu per satu sebagai file CSV terpisah
            for sheet_name, df in data_frames.items():
                csv_file_path = file_path.replace(".csv", f"_{sheet_name}.csv")
                df.to_csv(csv_file_path, index=False)

        messagebox.showinfo("Sukses", f"Semua sheet berhasil diekspor ke {file_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Gagal mengekspor sheet: {e}")