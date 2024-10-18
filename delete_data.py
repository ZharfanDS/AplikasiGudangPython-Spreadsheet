from tkinter import messagebox

def delete_data(self):
    from main import spreadsheet
    if hasattr(self, 'sheet'):
        if self.sheet.title == "Stock Barang":
            # Ambil data kode yang sudah ada di sheet
            existing_codes = [row[0] for row in self.sheet.get_all_values()[1:]]  # Misalnya kode ada di kolom pertama
            if not existing_codes:
                messagebox.showwarning("Warning", "Tidak ada kode yang tersedia untuk dihapus.")
                return
            
            # Tampilkan dropdown untuk memilih kode
            kode = self.create_dropdown_dialog("Kode untuk Dihapus datanya", existing_codes)
            if kode is None:
                return
            
            # Temukan nomor entri yang sesuai dengan kode yang dipilih
            entry_number = None
            for i, row in enumerate(self.sheet.get_all_values()[1:], start=1):  # Mulai dari 1 untuk melewati header
                if row[0] == kode:  # Misalnya kode ada di kolom pertama
                    entry_number = i
                    break

            if entry_number is None:
                messagebox.showwarning("Warning", "Kode tidak ditemukan.")
                return

            # Hapus baris berdasarkan nomor entri yang ditemukan
            self.sheet.delete_rows(entry_number + 1)
            messagebox.showinfo("Info", "Data telah dihapus.")
            self.display_data()
            
            
        elif self.sheet.title == "Barang Masuk":
            # Ambil semua data dari sheet, kecuali header
            all_data = self.sheet.get_all_values()[1:]  

            if not all_data:
                messagebox.showwarning("Warning", "Tidak ada data yang tersedia untuk dihapus.")
                return

            # Buat list deskripsi entri dengan format yang lebih jelas
            entry_options = [
                f"No: {i}\n " + "\n".join([f"{col}: {val}" for col, val in zip(self.sheet.get_all_values()[0], row)])
                for i, row in enumerate(all_data, start=1)
            ]

            # Tampilkan dialog dropdown untuk memilih barisan yang ingin dihapus
            selected_entry_desc = self.create_dropdown_dialog("Pilih barisan yang ingin dihapus", entry_options)

            if selected_entry_desc is None:
                return

            # Ambil nomor urutan entri yang dipilih
            entry_number = int(selected_entry_desc.split("\n")[0].replace("No: ", ""))

            # Dapatkan data dari entri yang dipilih
            selected_row = all_data[entry_number - 1]  # -1 untuk menyesuaikan dengan indeks list (karena entry_number dimulai dari 1)
            
            # Cari indeks kolom yang sesuai dengan "No. Purchasing Request", "Kode", dan "Jumlah Masuk"
            header = self.sheet.get_all_values()[0]
            no_purchasing_request_index = header.index("No. Purchasing Request")
            kode_masuk_index = header.index("Kode")
            jumlah_masuk_index = header.index("Jumlah Masuk")

            # Ambil nilai dari kolom yang diminta
            no_purchasing_request_value = selected_row[no_purchasing_request_index]
            kode_masuk_value = selected_row[kode_masuk_index]
            jumlah_masuk_value = selected_row[jumlah_masuk_index]

            # Print nilai kolom yang dipilih
            print(f"No. Purchasing Request: {no_purchasing_request_value}")
            print(f"Kode: {kode_masuk_value}")
            print(f"Jumlah Masuk: {jumlah_masuk_value}") 

            # Hapus baris berdasarkan nomor entri yang ditemukan
            self.sheet.delete_rows(entry_number + 1)  # +1 untuk menghindari header
            messagebox.showinfo("Info", "Data telah dihapus.")
            self.display_data()
            
            # Ambil data kode yang sudah ada di sheet
            # all_data = self.sheet.get_all_values()[1:]  # Ambil semua data, kecuali header
            # existing_codes = [row[3] for row in all_data]  # Misalnya kode ada di kolom ke-4 (indeks 3)
            
            # if not existing_codes:
            #     messagebox.showwarning("Warning", "Tidak ada data yang tersedia untuk dihapus.")
            #     return
            
            # # Tampilkan dropdown untuk memilih kode
            # kode_purchase_req = self.create_dropdown_dialog("No. Purchasing Requests untuk Dihapus datanya", existing_codes)
            # if kode_purchase_req is None:
            #     return
            
            # # Temukan semua entri yang sesuai dengan kode yang dipilih
            # matching_entries = [(i, row) for i, row in enumerate(all_data, start=1) if row[3] == kode_purchase_req]
            
            # if not matching_entries:
            #     messagebox.showwarning("Warning", "Kode tidak ditemukan.")
            #     return
            
            # # Jika ada lebih dari satu entri yang cocok, tampilkan pilihan urutan mana yang ingin dihapus
            # if len(matching_entries) > 1:
            #     # Buat list deskripsi entri (misalnya tampilkan kolom lain agar lebih informatif)
            #     entry_options = [
            #         f"No: {i}\n " + "\n".join([f"{col}: {val}" for col, val in zip(self.sheet.get_all_values()[0], row)])
            #         for i, row in matching_entries
            #     ]
                
            #     # Tampilkan dialog dropdown untuk memilih entri yang tepat
            #     selected_entry_desc = self.create_dropdown_dialog("Pilih entri yang ingin dihapus", entry_options)
                
            #     if selected_entry_desc is None:
            #         return
                
            #     # Ambil nomor urutan entri yang dipilih
            #     entry_number = int(selected_entry_desc.split("\n")[0].replace("No: ", " "))
                
            # else:
            #     # Jika hanya ada satu entri yang cocok, pilih yang pertama
            #     entry_number = matching_entries[0][0]
            
            # # Hapus baris berdasarkan nomor entri yang ditemukan
            # self.sheet.delete_rows(entry_number + 1)
            # messagebox.showinfo("Info", "Data telah dihapus.")
            # self.display_data()
        
        # NEW FIX
        elif self.sheet.title == "Barang Keluar":
            # Ambil semua data dari sheet, kecuali header
            all_data = self.sheet.get_all_values()[1:]  

            if not all_data:
                messagebox.showwarning("Warning", "Tidak ada data yang tersedia untuk dihapus.")
                return

            # Buat list deskripsi entri dengan format yang lebih jelas
            entry_options = [
                f"No: {i}\n " + "\n".join([f"{col}: {val}" for col, val in zip(self.sheet.get_all_values()[0], row)])
                for i, row in enumerate(all_data, start=1)
            ]

            # Tampilkan dialog dropdown untuk memilih barisan yang ingin dihapus
            selected_entry_desc = self.create_dropdown_dialog("Pilih barisan yang ingin dihapus", entry_options)

            if selected_entry_desc is None:
                return

            # Ambil nomor urutan entri yang dipilih
            entry_number = int(selected_entry_desc.split("\n")[0].replace("No: ", ""))

            # Dapatkan data dari entri yang dipilih
            selected_row = all_data[entry_number - 1]  # -1 untuk menyesuaikan dengan indeks list (karena entry_number dimulai dari 1)

            # Cari indeks kolom yang sesuai dengan "No. Material Request", "Kode", dan "Jumlah Keluar"
            header = self.sheet.get_all_values()[0]
            no_material_request_index = header.index("No. Material Request")
            kode_index = header.index("Kode")
            jumlah_keluar_index = header.index("Jumlah Keluar")

            # Ambil nilai dari kolom yang diminta
            no_material_request_value = selected_row[no_material_request_index]
            kode_value = selected_row[kode_index]
            jumlah_keluar_value = selected_row[jumlah_keluar_index]

            # Print nilai kolom yang dipilih
            print(f"No. Material Request: {no_material_request_value}")
            print(f"Kode: {kode_value}")
            print(f"Jumlah Keluar: {jumlah_keluar_value}") 
            
            # Dapatkan semua data dari sheet "MR"
            mr_sheet = spreadsheet.worksheet("MR")  # Pastikan mengganti 'Spreadsheet_Name' dengan nama spreadsheet Anda
            mr_data = mr_sheet.get_all_values()
            stockbarang_sheet = spreadsheet.worksheet("Stock Barang")
            stockbarang_data = stockbarang_sheet.get_all_values()
            
            # Cari indeks kolom di sheet "MR"
            mr_header = mr_data[0]
            no_material_request_index_mr = mr_header.index("No. Material Request")
            diserahkan_index = mr_header.index("Diserahkan")
            kekurangan_index = mr_header.index("Kekurangan")
            status_index = mr_header.index("Status")
            
            # Cari indeks kolom di sheet "Stock Barang"
            stockbarang_header = stockbarang_data[0]
            stockbarang_kode_index = stockbarang_header.index("Kode")
            stockbarang_minstock_index = stockbarang_header.index("Min Stock")
            stockbarang_stockawal_index = stockbarang_header.index("Stock Awal")
            stockbarang_keluar_index = stockbarang_header.index("Keluar")
            stockbarang_masuk_index = stockbarang_header.index("Masuk")
            stockbarang_stockakhir_index = stockbarang_header.index("Stock Akhir")
            stockbarang_keterangan_index = stockbarang_header.index("Keterangan")
            
            
            # Cari baris yang cocok dengan No. Material Request di MR
            for i, row in enumerate(mr_data[1:], start=2):  # Mulai dari 2 karena header adalah baris pertama
                if row[no_material_request_index_mr] == no_material_request_value:
                    # Dapatkan nilai saat ini dari kolom "Diserahkan", "Kekurangan", "Status"
                    diserahkan_value = int(row[diserahkan_index])
                    kekurangan_value = int(row[kekurangan_index])
                    
                    # Update nilai kolom "Diserahkan" dan "Kekurangan"
                    new_diserahkan_value = diserahkan_value - int(jumlah_keluar_value)
                    new_kekurangan_value = kekurangan_value + int(jumlah_keluar_value)
                    new_status_value = ""
                    if new_kekurangan_value == 0:
                        new_status_value = "Terpenuhi"
                    elif new_kekurangan_value != 0:
                        new_status_value = "Belum Terpenuhi"

                    # Update data di sheet "MR"
                    mr_sheet.update_cell(i, diserahkan_index + 1, new_diserahkan_value)  # i adalah baris, diserahkan_index + 1 untuk kolom (Google Sheets menggunakan 1-based index)
                    mr_sheet.update_cell(i, kekurangan_index + 1, new_kekurangan_value)
                    mr_sheet.update_cell(i, status_index + 1, new_status_value)

                    print(f"Data di sheet MR berhasil diperbarui:\nDiserahkan: {new_diserahkan_value}\nKekurangan: {new_kekurangan_value}\nStatus: {new_status_value}")
                    break
            else:
                print("No. Material Request tidak ditemukan di sheet MR.")
            
            # Cari baris yang cocok dengan Kode di Stock Barang
            for i, row in enumerate(stockbarang_data[1:], start=2):
                if row[stockbarang_kode_index] == kode_value:
                    keluar_value = int(row[stockbarang_keluar_index])
                    masuk_value = int(row[stockbarang_masuk_index])
                    stockawal_value = int(row[stockbarang_stockawal_index])
                    minstock_value = int(row[stockbarang_minstock_index])
                    
                    # Update nilai kolom "Keluar", "Stock Akhir", dan "Keterangan"
                    new_keluar_value = keluar_value - int(jumlah_keluar_value)
                    new_stockakhir_value = stockawal_value + masuk_value - new_keluar_value
                    new_keterangan_value = ""
                    if new_stockakhir_value < minstock_value:
                        new_keterangan_value = "Kurang"
                    else:
                        new_keterangan_value = "Cukup"
                        
                    stockbarang_sheet.update_cell(i, stockbarang_keluar_index + 1, new_keluar_value)
                    stockbarang_sheet.update_cell(i, stockbarang_stockakhir_index + 1, new_stockakhir_value)
                    stockbarang_sheet.update_cell(i, stockbarang_keterangan_index + 1, new_keterangan_value)

                    print(f"Data di sheet Stock Baranng berhasil diperbarui:\nKeluar: {new_keluar_value}\nStock Akhir: {new_stockakhir_value}\nKeterangan: {new_keterangan_value}")
            else:
                print("Kode tidak ditemukan di sheet Stock Barang.")
                
            # Hapus baris berdasarkan nomor entri yang ditemukan
            self.sheet.delete_rows(entry_number + 1)  # +1 untuk menghindari header
            messagebox.showinfo("Info", "Data telah dihapus.")
            self.display_data()
    else:
        messagebox.showwarning("Warning", "Pilih sheet terlebih dahulu.")