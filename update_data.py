from tkinter import simpledialog, messagebox

def update_data(self):
    if hasattr(self, 'sheet'):
        if self.sheet.title == "Stock Barang":
            # Ambil data kode yang sudah ada di sheet
            available_codes = [row[0] for row in self.sheet.get_all_values()[1:]]  # Asumsi kode ada di kolom pertama
            if not available_codes:
                messagebox.showwarning("Warning", "Tidak ada kode yang tersedia untuk diubah.")
                return
            
            # Tampilkan dropdown untuk memilih kode
            kode = self.create_dropdown_dialog("Pilih Kode untuk Diubah", available_codes)
            if kode is None:
                return
            
            # Temukan nomor entri yang sesuai dengan kode yang dipilih
            entry_number = None
            for i, row in enumerate(self.sheet.get_all_values()[1:], start=1):  # Mulai dari 1 untuk melewati header
                if row[0] == kode:  # Asumsi kode ada di kolom pertama
                    entry_number = i
                    break
            
            if entry_number is None:
                messagebox.showwarning("Warning", "Kode tidak ditemukan.")
                return
            
            # Ambil semua kolom dari header (baris pertama)
            columns = self.sheet.row_values(1)
            
            # Filter kolom yang tidak ingin ditampilkan di dropdown
            excluded_columns = ["Kode", "Kategori", "Nama Barang", "Satuan", "Masuk", "Keluar", "Stock Akhir", "Keterangan"]
            filtered_columns = [col for col in columns if col not in excluded_columns]

            # Tampilkan dialog untuk memilih kolom yang ingin diubah
            kolom = self.create_dropdown_dialog("Pilih Kolom yang Ingin Diubah", filtered_columns)
            if kolom is None or kolom not in filtered_columns:
                messagebox.showwarning("Warning", "Nama kolom tidak valid.")
                return
            
            # Minta pengguna untuk memasukkan nilai baru
            nilai_baru = simpledialog.askstring("Input", f"Masukkan nilai baru untuk {kolom}:")
            if nilai_baru is None:
                return
            
            # Update data di sheet
            col_index = columns.index(kolom) + 1
            row_index = entry_number + 1
            self.sheet.update_cell(row_index, col_index, nilai_baru)
            messagebox.showinfo("Info", "Data sedang diperbarui, harap tunggu.")
            self.refresh_data()
        
        elif self.sheet.title == "Barang Masuk":
            # Ambil data kode yang sudah ada di sheet
            available_pr = [row[3] for row in self.sheet.get_all_values()[1:]]  # Asumsi kode ada di kolom pertama
            if not available_pr:
                messagebox.showwarning("Warning", "Tidak ada Purchasing Request yang tersedia untuk diubah.")
                return
            
            # Tampilkan dropdown untuk memilih kode
            kode_purchasing_req = self.create_dropdown_dialog("No. Purchasing Requests untuk Diubah", available_pr)
            if kode_purchasing_req is None:
                return
            
            # Temukan nomor entri yang sesuai dengan kode yang dipilih
            entry_number = None
            for i, row in enumerate(self.sheet.get_all_values()[1:], start=1):  # Mulai dari 1 untuk melewati header
                if row[3] == kode_purchasing_req:  # Asumsi kode ada di kolom pertama
                    entry_number = i
                    break
            
            if entry_number is None:
                messagebox.showwarning("Warning", "No. Purchasing Requests tidak ditemukan.")
                return
            
            # Ambil semua kolom dari header (baris pertama)
            columns = self.sheet.row_values(1)
            
            # Filter kolom yang tidak ingin ditampilkan di dropdown
            excluded_columns = ["Tanggal Permintaan", "Tanggal Masuk", "TAT", "No. Purchasing Request", "Kode", "Kategori", "Nama Barang", "Satuan", "Jumlah Masuk", "Grand Total"]
            filtered_columns = [col for col in columns if col not in excluded_columns]

            # Tampilkan dialog untuk memilih kolom yang ingin diubah
            kolom = self.create_dropdown_dialog("Pilih Kolom yang Ingin Diubah", filtered_columns)
            if kolom is None or kolom not in filtered_columns:
                messagebox.showwarning("Warning", "Nama kolom tidak valid.")
                return
            
            # Minta pengguna untuk memasukkan nilai baru
            nilai_baru = simpledialog.askstring("Input", f"Masukkan nilai baru untuk {kolom}:")
            if nilai_baru is None:
                return
            
            # Update data di sheet
            col_index = columns.index(kolom) + 1
            row_index = entry_number + 1
            self.sheet.update_cell(row_index, col_index, nilai_baru)
            messagebox.showinfo("Info", "Data telah diperbarui.")
            self.display_data()
        
        elif self.sheet.title == "MR":
            # Ambil data kode yang sudah ada di sheet
            available_mr = [row[2] for row in self.sheet.get_all_values()[1:]]  # Asumsi kode ada di kolom pertama
            if not available_mr:
                messagebox.showwarning("Warning", "Tidak ada Material Request yang tersedia untuk diubah.")
                return
            
            # Tampilkan dropdown untuk memilih kode
            kode_material_req_mr = self.create_dropdown_dialog("No. Material Requests untuk Diubah", available_mr)
            if kode_material_req_mr is None:
                return
            
            # Temukan nomor entri yang sesuai dengan kode yang dipilih
            entry_number = None
            for i, row in enumerate(self.sheet.get_all_values()[1:], start=1):  # Mulai dari 1 untuk melewati header
                if row[2] == kode_material_req_mr:  # Asumsi kode ada di kolom pertama
                    entry_number = i
                    break
            
            if entry_number is None:
                messagebox.showwarning("Warning", "No. Material Requests tidak ditemukan.")
                return
            
            # Ambil semua kolom dari header (baris pertama)
            columns = self.sheet.row_values(1)
            
            # Filter kolom yang tidak ingin ditampilkan di dropdown
            excluded_columns = ["Tanggal Permintaan", "No. Material Request", "Kode", "Kategori", "Nama Barang", "Pesanan", "Diserahkan", "Kekurangan", "Status"]
            filtered_columns = [col for col in columns if col not in excluded_columns]

            # Tampilkan dialog untuk memilih kolom yang ingin diubah
            kolom = self.create_dropdown_dialog("Pilih Kolom yang Ingin Diubah", filtered_columns)
            if kolom is None or kolom not in filtered_columns:
                messagebox.showwarning("Warning", "Nama kolom tidak valid.")
                return
            
            # Minta pengguna untuk memasukkan nilai baru
            nilai_baru = simpledialog.askstring("Input", f"Masukkan nilai baru untuk {kolom}:")
            if nilai_baru is None:
                return
            
            # Update data di sheet
            col_index = columns.index(kolom) + 1
            row_index = entry_number + 1
            self.sheet.update_cell(row_index, col_index, nilai_baru)
            messagebox.showinfo("Info", "Data telah diperbarui.")
            self.display_data()    
            
        elif self.sheet.title == "Barang Keluar":
            # Ambil data kode yang sudah ada di sheet
            available_mr = [row[2] for row in self.sheet.get_all_values()[1:]]  # Asumsi kode ada di kolom pertama
            if not available_mr:
                messagebox.showwarning("Warning", "Tidak ada No. Material Request yang tersedia untuk diubah.")
                return
            
            # Tampilkan dropdown untuk memilih kode
            kode_material_req = self.create_dropdown_dialog("No. Material Request untuk Diubah", available_mr)
            if kode_material_req is None:
                return
            
            # Temukan nomor entri yang sesuai dengan kode yang dipilih
            entry_number = None
            for i, row in enumerate(self.sheet.get_all_values()[1:], start=1):  # Mulai dari 1 untuk melewati header
                if row[2] == kode_material_req:  # Asumsi kode ada di kolom pertama
                    entry_number = i
                    break
            
            if entry_number is None:
                messagebox.showwarning("Warning", "No. Material Request tidak ditemukan.")
                return
            
            # Ambil semua kolom dari header (baris pertama)
            columns = self.sheet.row_values(1)
            
            # Tampilkan dialog untuk memilih kolom yang ingin diubah
            kolom = self.create_dropdown_dialog("Pilih Kolom yang Ingin Diubah", columns)
            if kolom is None or kolom not in columns:
                messagebox.showwarning("Warning", "Nama kolom tidak valid.")
                return
            
            # Minta pengguna untuk memasukkan nilai baru
            nilai_baru = simpledialog.askstring("Input", f"Masukkan nilai baru untuk {kolom}:")
            if nilai_baru is None:
                return
            
            # Update data di sheet
            col_index = columns.index(kolom) + 1
            row_index = entry_number + 1
            self.sheet.update_cell(row_index, col_index, nilai_baru)
            messagebox.showinfo("Info", "Data telah diperbarui.")
            self.display_data()
        
        elif self.sheet.title == "PR":
            # Ambil data kode yang sudah ada di sheet
            available_mr = [row[2] for row in self.sheet.get_all_values()[1:]]  # Asumsi kode ada di kolom pertama
            if not available_mr:
                messagebox.showwarning("Warning", "Tidak ada Material Request yang tersedia untuk diubah.")
                return
            
            # Tampilkan dropdown untuk memilih kode
            kode_material_req_mr = self.create_dropdown_dialog("No. Material Requests untuk Diubah", available_mr)
            if kode_material_req_mr is None:
                return
            
            # Temukan nomor entri yang sesuai dengan kode yang dipilih
            entry_number = None
            for i, row in enumerate(self.sheet.get_all_values()[1:], start=1):  # Mulai dari 1 untuk melewati header
                if row[2] == kode_material_req_mr:  # Asumsi kode ada di kolom pertama
                    entry_number = i
                    break
            
            if entry_number is None:
                messagebox.showwarning("Warning", "No. Material Requests tidak ditemukan.")
                return
            
            # Ambil semua kolom dari header (baris pertama)
            columns = self.sheet.row_values(1)
            
            # Filter kolom yang tidak ingin ditampilkan di dropdown
            excluded_columns = ["Tanggal Permintaan", "No. Material Request", "Kode", "Kategori", "Nama Barang", "Pesanan", "Diserahkan", "Kekurangan", "Status"]
            filtered_columns = [col for col in columns if col not in excluded_columns]

            # Tampilkan dialog untuk memilih kolom yang ingin diubah
            kolom = self.create_dropdown_dialog("Pilih Kolom yang Ingin Diubah", filtered_columns)
            if kolom is None or kolom not in filtered_columns:
                messagebox.showwarning("Warning", "Nama kolom tidak valid.")
                return
            
            # Minta pengguna untuk memasukkan nilai baru
            nilai_baru = simpledialog.askstring("Input", f"Masukkan nilai baru untuk {kolom}:")
            if nilai_baru is None:
                return
            
            # Update data di sheet
            col_index = columns.index(kolom) + 1
            row_index = entry_number + 1
            self.sheet.update_cell(row_index, col_index, nilai_baru)
            messagebox.showinfo("Info", "Data telah diperbarui.")
            self.display_data()
            
    else:
        messagebox.showwarning("Warning", "Pilih sheet terlebih dahulu.")