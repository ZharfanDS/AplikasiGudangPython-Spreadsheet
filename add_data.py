from tkinter import simpledialog, messagebox

def add_data(self):
    from main import spreadsheet
    if hasattr(self, 'sheet'):
        if self.sheet.title == "Stock Barang":
            # Ambil data dari "Stock Report" untuk opsi dropdown
            stock_report = spreadsheet.worksheet("Stock Report")
            report_data = stock_report.get_all_records()

            if not report_data:
                messagebox.showwarning("Warning", "Sheet 'Stock Report' kosong. Tambahkan data di 'Stock Report' terlebih dahulu.")
                return

            # Ambil data dari Stock Barang untuk pengecekan kode yang sudah ada
            stock_barang_data = self.sheet.get_all_records()
            existing_codes = [row["Kode"] for row in stock_barang_data]

            # Ambil data untuk dropdown
            stock_report_options = {
                "Kode": [row["Kode Barang / Jasa"] for row in report_data],
            }

            # Menampilkan dropdown untuk memilih kode
            kode = self.create_dropdown_dialog("Kode", stock_report_options["Kode"])
            if kode is None:
                return

            # Cek apakah kode sudah ada di Stock Barang
            if kode in existing_codes:
                messagebox.showwarning("Warning", f"Kode {kode} sudah ada di Stock Barang.")
                return

            # Temukan baris yang sesuai di Stock Report
            selected_row = next((row for row in report_data if row["Kode Barang / Jasa"] == kode), None)
            if selected_row:
                # Isi otomatis kolom lainnya
                new_row = [
                    selected_row["Kode Barang / Jasa"],  # Kode
                    selected_row["Nama Barang / Jasa"],  # Nama Barang
                    selected_row["Kategori Barang / Jasa"],  # Kategori
                    selected_row["Satuan"]  # Satuan
                ]
            else:
                messagebox.showwarning("Warning", "Data tidak ditemukan di 'Stock Report'.")
                return
            
            # Input untuk nilai kolom terkait stok
            min_stock = simpledialog.askinteger("Input", "Masukkan Min Stock:")
            stock_awal = simpledialog.askinteger("Input", "Masukkan Stock Awal:")

            # Cek apakah ada input yang kosong
            if not all([min_stock, stock_awal]):
                messagebox.showwarning("Warning", "Semua field harus diisi.")
                return

            # Cek apakah nilai yang dimasukkan adalah numerik dan konversi ke int
            try:
                min_stock = int(min_stock)
                stock_awal = int(stock_awal)
            except ValueError:
                messagebox.showwarning("Warning", "Nilai yang dimasukkan harus berupa angka bulat.")
                return
            
            stock_akhir_data = stock_awal
            stock_masuk = 0
            stock_keluar = 0
            keterangan = ""
            if stock_akhir_data < min_stock:
                keterangan = "Kurang"
            else:
                keterangan = "Cukup"
            new_row.extend([min_stock, stock_awal, stock_masuk, stock_keluar, stock_akhir_data, keterangan])
            # Tambahkan baris baru ke Stock Barang
            self.sheet.append_row(new_row)
            messagebox.showinfo("Info", "Data telah ditambahkan.")
            self.display_data()
        
        elif self.sheet.title == "Barang Masuk":
            # Ambil data dari "Stock Barang" untuk pengecekan kode yang sudah ada
            stock_barang = spreadsheet.worksheet("Stock Barang")
            stock_barang_data = stock_barang.get_all_records()

            # Ambil data dari "Barang Masuk" untuk pengecekan kode yang sudah ada
            barang_masuk_data = self.sheet.get_all_records()

            data_pr = spreadsheet.worksheet("PR")
            data_on_pr = data_pr.get_all_records()
            
            # Input untuk Tanggal Masuk dan Tanggal Keluar
            tanggal_permintaan_masuk = simpledialog.askstring("Input", "Masukkan Tanggal Permintaan (format: DD-MM-YYYY):")
            tanggal_masuk = simpledialog.askstring("Input", "Masukkan Tanggal Masuk (format: DD-MM-YYYY):")

            # Cek apakah semua input ada
            if not tanggal_permintaan_masuk or not tanggal_masuk:
                messagebox.showwarning("Warning", "Tanggal Permintaan dan Tanggal Masuk harus diisi.")
                return

            # Konversi string tanggal ke objek tanggal
            try:
                from datetime import datetime
                format_tanggal = "%d-%m-%Y"
                tanggal_permintaan_obj = datetime.strptime(tanggal_permintaan_masuk, format_tanggal)
                tanggal_masuk_obj = datetime.strptime(tanggal_masuk, format_tanggal)
            except ValueError:
                messagebox.showwarning("Warning", "Format tanggal tidak valid. Gunakan format DD-MM-YYYY.")
                return

            # Hitung TAT sebagai selisih dalam hari
            tat = (tanggal_masuk_obj - tanggal_permintaan_obj).days
            
            available_pr = [row["No. Purchasing Request"] for row in data_on_pr]
            
            if not available_pr:
                messagebox.showwarning("Warning", "Tidak ada No. Purchasing Requests apapun di sheet PR untuk ditambahkan ke sheet Barang Keluar")
                return
            
            no_purchasingreq = self.create_dropdown_dialog("No. Purchasing Request", available_pr)
            if no_purchasingreq is None:
                return
            
            no_penyerahan_masuk = simpledialog.askstring("Input", "Masukkan data untuk No. Penyerahan:")
            pembelian = simpledialog.askstring("Input", "Masukkan data untuk Pembelian:")

            # Temukan baris yang sesuai di Stock Report
            selected_row = next((row for row in data_on_pr if row["No. Purchasing Request"] == no_purchasingreq), None)
            if not selected_row:
                messagebox.showwarning("Warning", "Data tidak ditemukan di 'Stock Report'.")
                return


            # Tambahkan baris baru di "Barang Masuk"
            new_row = [
                tanggal_permintaan_masuk,        # Tanggal Permintaan
                tanggal_masuk,             # Tanggal Masuk
                tat,                       # TAT
                no_purchasingreq,          # No. Purchasing Request
                no_penyerahan_masuk,       # No. Penyerahan
                pembelian,                 # Pembelian
                selected_row["Kode"],  # Kode
                selected_row["Kategori"],  # Kategori
                selected_row["Nama Barang"],  # Nama Barang
                selected_row["Satuan"],    # SatuaN
            ]
            
            stock_kode_list = [row["Kode"] for row in stock_barang_data]
            if selected_row["Kode"] not in stock_kode_list:
                messagebox.showwarning("Warning", "Masukkan terlebih dahulu Stock Barang untuk kode ini!")
                return
            
            # Hitung total "Jumlah Masuk" untuk kode yang sama
            total_jumlah_masuk = 0
            for row in barang_masuk_data:
                if row["Kode"] == selected_row["Kode"]:
                    try:
                        total_jumlah_masuk += int(row["Jumlah Masuk"])
                    except (ValueError, TypeError):
                        continue
            
            for row in data_on_pr:
                if row["No. Purchasing Request"] == no_purchasingreq:
                    permintaan_pesanan_pr = int(row["Pesanan"])
                    permintaan_diterima_pr = int(row["Diterima"])
                    #print(permintaan_pesanan_pr)
                    
            pesanan_sisa_pr = permintaan_pesanan_pr - permintaan_diterima_pr
            
            if pesanan_sisa_pr == 0:
                messagebox.showwarning("Warning", "PESANAN INI SUDAH TERSELESAIKAN!")
                return
            
            pesanan = pesanan_sisa_pr 
            # simpledialog.askinteger("Input", "Masukkan data untuk Pesanan:")
            jumlah_masuk = simpledialog.askinteger("Input", "Masukkan data untuk Jumlah Masuk:")
            
            if jumlah_masuk > permintaan_pesanan_pr:
                messagebox.showwarning("Warning", "Jumlah keluar tidak boleh melebihi permintaan pesanan!")
                return
            
            if jumlah_masuk > pesanan_sisa_pr:
                messagebox.showwarning("Warning", "Jumlah keluar tidak boleh melebihi permintaan pesanan! pesanan ini hanya kekurangan : " + str(pesanan_sisa_pr) +".")
                return
            
            # Cek apakah `jumlah_masuk` kosong, beri default nilai 0
            if not jumlah_masuk:
                jumlah_masuk = 0
            
            total_jumlah_masuk += jumlah_masuk  # Tambahkan jumlah masuk baru
            
            harga_satuan_masuk = simpledialog.askinteger("Input", "Masukkan data untuk Harga Satuan:")
            ppn = simpledialog.askstring("Input", "Masukkan data untuk PPN:")
            grand_total_price_masuk = jumlah_masuk * harga_satuan_masuk
            supplier = simpledialog.askstring("Input", "Masukkan data untuk Supplier:")
            
            # Tambahkan input lainnya ke new_row
            new_row.extend([pesanan, jumlah_masuk, harga_satuan_masuk, ppn, grand_total_price_masuk, supplier])
            # Tambahkan baris baru ke "Barang Masuk"
            self.sheet.append_row(new_row)
            
            # Update total "Jumlah Masuk" di "Stock Barang"
            row_in_stock_barang = next((row for row in stock_barang_data if row["Kode"] == selected_row["Kode"]), None)
            if row_in_stock_barang:
                stock_awal = row_in_stock_barang.get("Stock Awal", 0)
                keluar = row_in_stock_barang.get("Keluar", 0)

                try:
                    stock_awal = int(stock_awal) if str(stock_awal).strip() else 0
                except (ValueError, TypeError):
                    stock_awal = 0

                try:
                    keluar = int(keluar) if str(keluar).strip() else 0
                except (ValueError, TypeError):
                    keluar = 0

                # Hitung Stock Akhir
                stock_akhir = stock_awal + total_jumlah_masuk - keluar

                # Update jumlah masuk dan stock akhir di sheet "Stock Barang"
                for i, row in enumerate(stock_barang_data, start=2):  # Mulai dari baris ke-2 karena baris pertama adalah header
                    if row["Kode"] == selected_row["Kode"]:
                        ambil_jumlah_minimal_barang_masuk = int(row["Min Stock"]) if isinstance(row["Min Stock"], str) and row["Min Stock"].isdigit() else row["Min Stock"]
                        if ambil_jumlah_minimal_barang_masuk == 0:
                            messagebox.showwarning("Warning", "Masukkan terlebih dahulu minimum stock pada barang untuk kode ini!")
                            return
                        
                        keterangan_recent_barang_masuk = str(row["Keterangan"]) 
                        # # Cek apakah 'Keterangan' kosong atau hanya berisi whitespace
                        # if not keterangan_recent_barang_masuk.strip():
                        #     print("Keterangan kosong")
                        # else:
                        #     print("Keterangan diisi:", keterangan_recent_barang_masuk)
                            
                        stock_barang.update_cell(i, stock_barang.find("Masuk").col, total_jumlah_masuk)
                        stock_barang.update_cell(i, stock_barang.find("Stock Akhir").col, stock_akhir)
                        
                        if stock_akhir < ambil_jumlah_minimal_barang_masuk:
                            keterangan = "Kurang"
                            keterangan_recent_barang_masuk = keterangan
                        elif stock_akhir > ambil_jumlah_minimal_barang_masuk :
                            keterangan = "Cukup"
                            keterangan_recent_barang_masuk = keterangan
                        stock_barang.update_cell(i, stock_barang.find("Keterangan").col, keterangan_recent_barang_masuk)
                        break
                
                for i, row in enumerate(data_on_pr, start=2):
                    if row ["No. Purchasing Request"] == no_purchasingreq:
                        ambil_jumlah_diterima = int(row["Diterima"])
                        ambil_jumlah_diterima += jumlah_masuk
                        
                        data_pr.update_cell(i, data_pr.find("Diterima").col, ambil_jumlah_diterima)
                        
                        pr_ambil_pesanan_mr = int(row["Pesanan"])
                        
                        kekurangan_pr = pr_ambil_pesanan_mr - ambil_jumlah_diterima
                        
                        data_pr.update_cell(i, data_pr.find("Kekurangan").col, kekurangan_pr)
                        
                        status_data_pr = str(row["Status"])
                        
                        if kekurangan_pr == 0:
                            status = "Terpenuhi"
                            status_data_pr = status
                        elif ambil_jumlah_diterima == 0 & kekurangan_pr == pr_ambil_pesanan_mr:
                            status = "Belum Terpenuhi"
                            status_data_pr = status
                        elif ambil_jumlah_diterima < pr_ambil_pesanan_mr:
                            status = "Terpenuhi Sebagian"
                            status_data_pr = status
                            
                        data_pr.update_cell(i, data_pr.find("Status").col, status_data_pr)
                        break #test
                        
            messagebox.showinfo("Info", "Data telah ditambahkan di Barang Masuk dan Stock Barang telah diperbarui.")
            self.display_data()
        
        elif self.sheet.title == "Barang Keluar":
            # Ambil data dari "Stock Barang" untuk pengecekan kode yang sudah ada
            stock_barang = spreadsheet.worksheet("Stock Barang")
            stock_barang_data = stock_barang.get_all_records()
            
            data_mr = spreadsheet.worksheet("MR")
            data_on_mr = data_mr.get_all_records()
            
            # Ambil data dari "Barang Keluar" untuk pengecekan kode yang sudah ada
            barang_keluar_data = self.sheet.get_all_records()
            
            tanggal_permintaan_keluar = simpledialog.askstring("Input", "Masukkan Tanggal Permintaan (format: DD-MM-YYYY):")
            # Konversi string tanggal ke objek tanggal
            try:
                from datetime import datetime
                format_tanggal = "%d-%m-%Y"
                tanggal_permintaan_keluar_obj = datetime.strptime(tanggal_permintaan_keluar, format_tanggal)
            except ValueError:
                messagebox.showwarning("Warning", "Format tanggal tidak valid. Gunakan format DD-MM-YYYY.")
                return
            
            user = simpledialog.askstring("Input", "Masukkan User:")
            # no_material_request = simpledialog.askstring("Input", "Masukkan No. Material Request:")
            
            available_mr = [row["No. Material Request"] for row in data_on_mr]
            
            if not available_mr:
                messagebox.showwarning("Warning", "Tidak ada No. Material Requests apapun di sheet MR untuk ditambahkan ke sheet Barang Keluar")
                return
            
            no_material_request = self.create_dropdown_dialog("No. Material Request", available_mr)
            if no_material_request is None:
                return
            
            no_penyerahan_keluar = simpledialog.askstring("Input", "Masukkan No. Penyerahan:")
                
            # Ambil data dari "Stock Report"
            stock_report = spreadsheet.worksheet("Stock Report")
            report_data = stock_report.get_all_records()

            if not report_data:
                messagebox.showwarning("Warning", "Sheet 'Stock Report' kosong atau tidak dapat diambil.")
                return
            
            # Temukan baris yang sesuai di Stock Report berdasarkan kode yang dipilih
            selected_row = next((row for row in data_on_mr if row["No. Material Request"] == no_material_request), None)
            if not selected_row:
                messagebox.showwarning("Warning", "Data tidak ditemukan di 'MR'.")
                return

            # Isi otomatis kolom lainnya dari Stock Report
            new_row = [
                tanggal_permintaan_keluar,        # Tanggal Permintaan
                user,                             # User
                no_material_request,              # No. Material Request
                no_penyerahan_keluar,             # No. Penyerahan Keluar
                selected_row["Kode"],  # Kode
                selected_row["Kategori"],  # Kategori
                selected_row["Nama Barang"],  # Nama Barang
                selected_row["Satuan"],  # Satuan
            ]
            
            stock_kode_list = [row["Kode"] for row in stock_barang_data]
            if selected_row["Kode"] not in stock_kode_list:
                messagebox.showwarning("Warning", "Masukkan terlebih dahulu Stock Barang untuk kode ini!")
                return
            
            for row in stock_barang_data:
                if row["Kode"] == selected_row["Kode"]:
                    stock_akhir_barang_keluar = row.get("Stock Akhir")  # Mengambil nilai dari kolom 'Stock Akhir'
                    # Cek apakah nilai kosong (None) atau string kosong
                    if not stock_akhir_barang_keluar or stock_akhir_barang_keluar == 0:
                        messagebox.showwarning("Warning", "Stock Kosong! Masukkan terlebih dahulu stock masuk pada barang untuk kode ini!")
                        return
            
            total_jumlah_keluar = 0
            for row in barang_keluar_data:
                if row["Kode"] == selected_row["Kode"]:
                    try:
                        total_jumlah_keluar += int(row["Jumlah Keluar"])
                    except (ValueError, TypeError):
                        continue
            
            for row in data_on_mr:
                if row["No. Material Request"] == no_material_request:
                    permintaan_pesanan_mr = int(row["Pesanan"])
                    permintaan_diserahkan_mr = int(row["Diserahkan"])
                    # print(permintaan_pesanan_mr)
                    
            pesanan_sisa = permintaan_pesanan_mr - permintaan_diserahkan_mr
            
            if pesanan_sisa == 0:
                messagebox.showwarning("Warning", "PESANAN INI SUDAH TERSELESAIKAN!")
                return
            
            jumlah_keluar = simpledialog.askinteger("Input", "Masukkan data untuk Jumlah Keluar:")
            
            if jumlah_keluar > permintaan_pesanan_mr:
                messagebox.showwarning("Warning", "Jumlah keluar tidak boleh melebihi permintaan pesanan!")
                return
            
            if jumlah_keluar > pesanan_sisa:
                messagebox.showwarning("Warning", "Jumlah keluar tidak boleh melebihi permintaan pesanan! pesanan ini hanya kekurangan : " + str(pesanan_sisa) +".")
                return
            
            if jumlah_keluar > stock_akhir_barang_keluar:
                messagebox.showwarning("Warning", "Jumlah yang dikeluarkan lebih banyak dari stock akhir barang!")
                messagebox.showwarning("Warning", "Masukkan barang terlebih dahulu karena kekurangan stock!")
                return
            
            total_jumlah_keluar += jumlah_keluar  # Tambahkan jumlah keluar baru

            harga_satuan_keluar = simpledialog.askstring("Input", "Masukkan data untuk Harga Satuan:")
            # Menghitung grand total price
            if harga_satuan_keluar == "":
                grand_total_price_keluar = ""
            else:
                grand_total_price_keluar = "Rp. " + str(jumlah_keluar * int(harga_satuan_keluar))  # Mengonversi hasil ke string

            tanggal_keluar = simpledialog.askstring("Input", "Masukkan Tanggal Keluar (format: DD-MM-YYYY):")
            # Konversi string tanggal ke objek tanggal
            try:
                tanggal_keluar_obj = datetime.strptime(tanggal_keluar, format_tanggal)
            except ValueError:
                messagebox.showwarning("Warning", "Format tanggal tidak valid. Gunakan format DD-MM-YYYY.")
                return

            # Tambahkan data ke new_row
            new_row.extend([jumlah_keluar, harga_satuan_keluar, grand_total_price_keluar, tanggal_keluar])

            # Tambahkan baris baru ke "Barang Keluar"
            self.sheet.append_row(new_row)

            # Update total "Jumlah Masuk" di "Stock Barang"
            row_in_stock_barang = next((row for row in stock_barang_data if row["Kode"] == selected_row["Kode"]), None)
            if row_in_stock_barang:
                stock_awal = row_in_stock_barang.get("Stock Awal", 0)
                masuk = row_in_stock_barang.get("Masuk", 0)

                try:
                    stock_awal = int(stock_awal) if str(stock_awal).strip() else 0
                except (ValueError, TypeError):
                    stock_awal = 0
                    
                try:
                    masuk = int(masuk) if str(masuk).strip() else 0
                except (ValueError, TypeError):
                    masuk = 0

                # Hitung Stock Akhir
                # print(stock_awal)
                # print(masuk)
                # print(total_jumlah_keluar)
                
                stock_akhir = stock_awal + masuk - total_jumlah_keluar

                # Update jumlah masuk dan stock akhir di sheet "Stock Barang"
                for i, row in enumerate(stock_barang_data, start=2):  # Mulai dari baris ke-2 karena baris pertama adalah header
                    if row["Kode"] == selected_row["Kode"]:
                        ambil_jumlah_minimal_barang_keluar = int(row["Min Stock"])
                        keterangan_recent = str(row["Keterangan"])
                            
                        stock_barang.update_cell(i, stock_barang.find("Keluar").col, total_jumlah_keluar)
                        stock_barang.update_cell(i, stock_barang.find("Stock Akhir").col, stock_akhir)
                        
                        if stock_akhir < ambil_jumlah_minimal_barang_keluar:
                            keterangan = "Kurang"
                            keterangan_recent = keterangan
                        elif stock_akhir > ambil_jumlah_minimal_barang_keluar :
                            keterangan = "Cukup"
                            keterangan_recent = keterangan
                        stock_barang.update_cell(i, stock_barang.find("Keterangan").col, keterangan_recent)
                        break
                    
                for i, row in enumerate(data_on_mr, start=2):
                    if row ["No. Material Request"] == no_material_request:
                        ambil_jumlah_keluar = int(row["Diserahkan"])
                        ambil_jumlah_keluar += jumlah_keluar
                        
                        data_mr.update_cell(i, data_mr.find("Diserahkan").col, ambil_jumlah_keluar)
                        
                        ambil_pesanan_mr = int(row["Pesanan"])
                        
                        kekurangan_mr = ambil_pesanan_mr - ambil_jumlah_keluar
                        
                        data_mr.update_cell(i, data_mr.find("Kekurangan").col, kekurangan_mr)
                        
                        if kekurangan_mr == 0:
                            status = "Terpenuhi"
                        elif ambil_jumlah_keluar == 0 & kekurangan_mr == ambil_pesanan_mr:
                            status = "Belum Terpenuhi"
                        elif ambil_jumlah_keluar < ambil_pesanan_mr:
                            status = "Terpenuhi Sebagian"
                            
                        data_mr.update_cell(i, data_mr.find("Status").col, status)
                        
            messagebox.showinfo("Info", "Data telah ditambahkan di Barang Masuk dan Stock Barang telah diperbarui.")
            self.display_data()
        
        elif self.sheet.title == "MR":
            stock_report = spreadsheet.worksheet("Stock Report")
            mr_data = spreadsheet.worksheet("MR")
            get_mr_data = mr_data.get_all_records()
            tanggal_permintaan_mr = simpledialog.askstring("Input", "Masukkan Tanggal Permintaan (format: DD-MM-YYYY):")
            # Konversi string tanggal ke objek tanggal
            try:
                from datetime import datetime
                format_tanggal = "%d-%m-%Y"
                tanggal_permintaan_mr_obj = datetime.strptime(tanggal_permintaan_mr, format_tanggal)
            except ValueError:
                messagebox.showwarning("Warning", "Format tanggal tidak valid. Gunakan format DD-MM-YYYY.")
                return
            
            user_mr = simpledialog.askstring("Input", "Masukkan User:")
            no_material_request_mr = simpledialog.askstring("Input", "Masukkan No. Material Request:")
            
            # Cek apakah No. Material Request sudah ada di sheet MR
            no_material_request_exists = False
            for row in get_mr_data:
                if row.get("No. Material Request") == no_material_request_mr:
                    no_material_request_exists = True
                    break

            if no_material_request_exists:
                # print("No. Material Request sudah ada!")
                messagebox.showinfo("Info", "No. Material Request sudah ada!")
                return
                
            report_data = stock_report.get_all_records()
            # Ambil data kode yang belum ada di "Barang Masuk"
            available_codes = [row["Kode Barang / Jasa"] for row in report_data]

            if not available_codes:
                messagebox.showwarning("Warning", "Tidak ada kode baru yang tersedia untuk ditambahkan.")
                return

            # Tampilkan dropdown untuk memilih kode
            kode = self.create_dropdown_dialog("Kode", available_codes)
            if kode is None:
                return
            
            # Temukan baris yang sesuai di Stock Report berdasarkan kode yang dipilih
            selected_row = next((row for row in report_data if row["Kode Barang / Jasa"] == kode), None)
            if not selected_row:
                messagebox.showwarning("Warning", "Data tidak ditemukan di 'Stock Report'.")
                return
            
            pesanan_mr = simpledialog.askstring("Input", "Masukkan jumlah pesanan:")
            
            try:
                pesanan_mr = int(pesanan_mr)  # Pastikan pesanan adalah angka
            except ValueError:
                messagebox.showwarning("Warning", "Jumlah pesanan harus berupa angka.")
                return
            
            # Isi otomatis kolom lainnya dari Stock Report
            new_row = [
                tanggal_permintaan_mr,        # Tanggal Permintaan
                user_mr,                      # User
                no_material_request_mr,       # No. Material Request
                selected_row["Kode Barang / Jasa"],  # Kode
                selected_row["Kategori Barang / Jasa"],  # Kategori
                selected_row["Nama Barang / Jasa"],  # Nama Barang
                selected_row["Satuan"],       # Satuan
                pesanan_mr                    # Pesanan
            ]
            
            # Ambil data dari sheet "Barang Keluar" untuk pengecekan jumlah diserahkan
            barang_keluar = spreadsheet.worksheet("Barang Keluar")
            barang_keluar_data = barang_keluar.get_all_records()

            # Cari jumlah diserahkan berdasarkan No. Material Request dan Kode
            diserahkan_mr = 0  # Default jika tidak ada data
            for row in barang_keluar_data:
                if row["No. Material Request"] == no_material_request_mr and row["Kode"] == kode:
                    diserahkan_mr += int(row["Jumlah Keluar"])
            
            # Hitung kekurangan
            kekurangan = pesanan_mr - diserahkan_mr
            
            # Tentukan status berdasarkan kekurangan
            if kekurangan == 0:
                status_mr = "Terpenuhi"
            elif diserahkan_mr == 0:
                status_mr = "Belum Terpenuhi"
            elif diserahkan_mr < pesanan_mr:
                status_mr = "Terpenuhi Sebagian"
            
            # Tambahkan kolom Diserahkan, Kekurangan, dan Status ke new_row
            new_row.extend([diserahkan_mr, kekurangan, status_mr])
            
            # Tambahkan baris baru ke sheet "MR"
            self.sheet.append_row(new_row)
            
            messagebox.showinfo("Info", "Data telah ditambahkan di MR.")
            self.display_data()

        elif self.sheet.title == "PR":
            mr_report = spreadsheet.worksheet("MR")
            pr_data_on_mr = mr_report.get_all_records()
            
            stock_barang = spreadsheet.worksheet("Stock Barang")
            stock_barang_data = stock_barang.get_all_records()
            
            pr_report = spreadsheet.worksheet("PR")
            get_pr_data = pr_report.get_all_records()
            
            tanggal_pembelian = simpledialog.askstring("Input", "Masukkan Tanggal Pembelian (format: DD-MM-YYYY):")
            # Konversi string tanggal ke objek tanggal
            try:
                from datetime import datetime
                format_tanggal = "%d-%m-%Y"
                tanggal_pembelian_mr_obj = datetime.strptime(tanggal_pembelian, format_tanggal)
            except ValueError:
                messagebox.showwarning("Warning", "Format tanggal tidak valid. Gunakan format DD-MM-YYYY.")
                return
            
            no_purchasing_request_pr = simpledialog.askstring("Input", "Masukkan No. Purchasing Request:")
            
            # Cek apakah No. Material Request sudah ada di sheet MR
            no_purchasing_request_exists = False
            for row in get_pr_data:
                if row.get("No. Purchasing Request") == no_purchasing_request_pr:
                    no_purchasing_request_exists = True
                    break

            if no_purchasing_request_exists:
                # print("No. Purchasing Request sudah ada!")
                messagebox.showinfo("Info", "No. Purchasing Request sudah ada!")
                return
            
            available_mr_in_pr = [row["No. Material Request"] for row in pr_data_on_mr]
            
            if not available_mr_in_pr:
                messagebox.showwarning("Warning", "Tidak ada No. Material Requests apapun di sheet MR untuk ditambahkan ke sheet Barang Keluar")
                return
            
            no_material_request_pr = self.create_dropdown_dialog("No. Material Request", available_mr_in_pr)
            if no_material_request_pr is None:
                return
            
            # Cek apakah ada No Material Request yang sama di sheet "PR"
            no_material_request_exist = False
            for row in get_pr_data:
                if row.get("No. Material Request") == no_material_request_pr:
                    no_material_request_exist = True
                    break
            
            if no_material_request_exist:
                messagebox.showinfo("Info", "No. Material Request sudah ada!")
                return
            
            for row in pr_data_on_mr:
                # test = row.get("No. Material Request")
                if row.get("No. Material Request") == no_material_request_pr:
                    status = str(row["Status"])
                    # print(test, no_material_request_pr, status)
                    if status == "Terpenuhi":
                        messagebox.showinfo("Info", "No. Material Request sudah terpenuhi!")
                        return
            
            pembelian_pr = simpledialog.askstring("Input", "Masukkan data Pembelian:")
            
            selected_row = next((row for row in pr_data_on_mr if row["No. Material Request"] == no_material_request_pr), None)
            if not selected_row:
                messagebox.showwarning("Warning", "Data tidak ditemukan di 'MR'.")
                return
            
            # Isi otomatis kolom lainnya dari Stock Report
            new_row = [
                tanggal_pembelian,        # Tanggal Permintaan
                no_purchasing_request_pr,    
                no_material_request_pr,     # No. Material Request
                pembelian_pr, 
                selected_row["Kode"],  # Kode
                selected_row["Kategori"],  # Kategori
                selected_row["Nama Barang"],  # Nama Barang
                selected_row["Satuan"],       # Satuan
            ]
            
            stock_kode_list = [row["Kode"] for row in stock_barang_data]
            if selected_row["Kode"] not in stock_kode_list:
                messagebox.showwarning("Warning", "Masukkan terlebih dahulu Stock Barang untuk kode ini!")
                return
            
            for row in pr_data_on_mr:
                if row["No. Material Request"] == no_material_request_pr:
                    check_kekurangan = row.get("Kekurangan")
                    pesanan_pr_temp = check_kekurangan
                    # print("Pesanan : " + str(pesanan_pr_temp))
            
            pesanan_pr = pesanan_pr_temp
            
            # Ambil data dari sheet "Barang Keluar" untuk pengecekan jumlah diserahkan
            barang_masuk = spreadsheet.worksheet("Barang Masuk")
            barang_masuk_data = barang_masuk.get_all_records()

            # Cari jumlah diserahkan berdasarkan No. Material Request dan Kode
            diterima_pr = 0  # Default jika tidak ada data
            for row in barang_masuk_data:
                if row["No. Purchasing Request"] == no_material_request_pr and row["Kode"] == kode:
                    diterima_pr += int(row["Jumlah Masuk"])
        
            # Hitung kekurangan
            kekurangan_pr = pesanan_pr - diterima_pr
            
            # Tentukan status berdasarkan kekurangan
            if kekurangan_pr == 0:
                status_pr = "Terpenuhi"
            elif diterima_pr == 0:
                status_pr = "Belum Terpenuhi"
            elif diterima_pr < pesanan_mr:
                status_pr = "Terpenuhi Sebagian"
                
            user_pr = simpledialog.askstring("Input", "Masukkan data User:")
            
            remark_pr = simpledialog.askstring("Input", "Masukkan Remark:")
            
            # Tambahkan kolom Diserahkan, Kekurangan, dan Status ke new_row
            new_row.extend([pesanan_pr, diterima_pr, kekurangan_pr, status_pr, user_pr, remark_pr])
            
            # Tambahkan baris baru ke sheet "PR"
            self.sheet.append_row(new_row)
            
            messagebox.showinfo("Info", "Data telah ditambahkan di PR.")
            self.display_data()
    else:
        messagebox.showwarning("Warning", "Pilih sheet terlebih dahulu.")