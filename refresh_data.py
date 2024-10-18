import tkinter as tk
from tkinter import messagebox

def refresh_data(self):
    if hasattr(self, 'sheet') and self.sheet.title == "Stock Barang":       # Fungsi untuk refresh data dari Google Sheets dan update "Stock Akhir"
        # Dapatkan data terbaru dari Google Sheets
        existing_codes = [row[0] for row in self.sheet.get_all_values()[1:]]  # Misalnya kode ada di kolom pertama
        if not existing_codes:
            messagebox.showwarning("Warning", "Tidak ada kode yang tersedia untuk dihapus.")
            return
        
        sheet_data = self.sheet.get_all_records()

        # Bersihkan teks area sebelum menampilkan data baru
        self.text_area.delete(1.0, tk.END)

        # Iterasi melalui setiap baris data di "Stock Barang"
        for i, row in enumerate(sheet_data, start=2):  # Mulai dari baris kedua (baris pertama adalah header)
            # Ambil nilai dari setiap kolom, jika kosong atau tidak valid, set nilai default 0
            stock_awal = row.get("Stock Awal", 0)
            masuk = row.get("Masuk", 0)
            keluar = row.get("Keluar", 0)
            min_stock_barang = row.get("Min Stock", 0)

            try:
                # Pengecekan apakah nilai Stock Awal kosong atau tidak valid
                if not stock_awal or not isinstance(stock_awal, (int, float)):
                    stock_awal = 0

                # Konversi nilai-nilai yang diambil ke integer jika diperlukan
                masuk = int(masuk) if masuk else 0
                keluar = int(keluar) if keluar else 0

                # Hitung Stock Akhir dengan rumus Stock Akhir = Stock Awal + Masuk - Keluar
                stock_akhir = stock_awal + masuk - keluar
                print("Stock awal: ", stock_awal)
                print("Masuk: ", masuk)
                print("Keluar: ", keluar)
                print("Stock akhir: ", stock_akhir)
                
                status_barang = ""
                if stock_akhir < min_stock_barang:
                    status_barang = "Kurang"
                else:
                    status_barang = "Cukup"
                
                stock_akhir_column = self.sheet.find("Stock Akhir")
                if stock_akhir_column is not None:
                    # Update nilai "Stock Akhir" di sheet
                    self.sheet.update_cell(i, self.sheet.find("Stock Akhir").col, stock_akhir)
                    self.sheet.update_cell(i, self.sheet.find("Keterangan").col, status_barang)
                else:
                    messagebox.showwarning("Warning", "Kolom Stock Akhir tidak ditemukan di sheet")
                    #print("Kolom 'Stock Akhir' tidak ditemukan di sheet.")

                # Tampilkan baris data yang sudah diperbarui di text area
                self.text_area.insert(tk.END, f"{row}, Stock Akhir: {stock_akhir}\n")
                
            except ValueError:
                # Jika ada nilai yang tidak valid, lewati dan tampilkan pesan di text area
                self.text_area.insert(tk.END, f"Data invalid pada baris {i}: {row}\n")

        messagebox.showinfo("Info", "Data 'Stock Akhir' telah di-refresh dan diperbarui.")
    else:
        messagebox.showwarning("Warning", "Refresh button hanya bisa digunakan di sheet 'Stock Barang'.")

    #print("TEST DISPLAY")
    self.display_data()