from tkinter import messagebox

def search_text(self):
    if not hasattr(self, 'sheet') or self.sheet is None:
        messagebox.showwarning("Peringatan", "Pilih sheet terlebih dahulu.")
        return
    
    search_query = self.search_entry.get().lower()  # Ambil input pencarian dan ubah jadi huruf kecil
    if not search_query:
        messagebox.showwarning("Peringatan", "Masukkan teks yang ingin dicari.")
        return
    
    sheet_data = self.sheet.get_all_records()
    search_results = []
    
    # Cari teks di semua kolom
    for row in sheet_data:
        for key, value in row.items():
            if search_query in str(value).lower():
                search_results.append(row)
                break
    
    if search_results:
        result_text = "\n".join([str(result) for result in search_results])
        messagebox.showinfo("Hasil Pencarian", result_text)
    else:
        messagebox.showinfo("Hasil Pencarian", "Tidak ada data yang ditemukan.")