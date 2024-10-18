import tkinter as tk

def create_widgets(self):
    # Frame untuk sheet selection
    frame_selection = tk.Frame(self)
    frame_selection.pack(fill='x', padx=10, pady=10)
    
    # Tambahkan frame baru untuk teks dan button, dan tempatkan di tengah
    frame_center = tk.Frame(self)
    frame_center.pack(pady=1)  # Beri jarak vertikal agar tidak terlalu ke atas

    # Tambahkan label "Select Sheet :" di frame_center
    label = tk.Label(frame_center, text="Select Sheet :")
    label.pack(side='left', padx=1)

    # Membuat buttons untuk setiap sheet di frame_center
    sheets = ["Stock Barang", "MR", "Barang Keluar", "PR", "Barang Masuk", "Stock Report"]
    for sheet_name in sheets:
        button = tk.Button(frame_center, text=sheet_name, command=lambda name=sheet_name: self.on_sheet_select(name))
        button.pack(side='left', padx=1)

    # Tambahkan label di bawah button untuk menampilkan nama sheet yang dipilih
    self.selected_sheet_label = tk.Label(self, text="Sheet yang dipilih: -")
    self.selected_sheet_label.pack(pady=5)

    # Frame untuk area teks dan tombol
    frame_main = tk.Frame(self)
    frame_main.pack(fill='both', expand=True, padx=10, pady=10)

    # Text area to display sheet data
    self.text_area = tk.Text(frame_main, wrap='none', height=20, state='disabled')
    self.text_area.pack(side='left', fill='both', expand=True, padx=5, pady=5)

    # Scrollbar untuk area teks
    self.scroll_y = tk.Scrollbar(frame_main, orient='vertical', command=self.text_area.yview)
    self.scroll_y.pack(side='right', fill='y')
    self.text_area.config(yscrollcommand=self.scroll_y.set)

    # Frame untuk scrollbar horizontal
    frame_scroll_x = tk.Frame(self)
    frame_scroll_x.pack(fill='x')
    # Scrollbar horizontal untuk area teks
    self.scroll_x = tk.Scrollbar(frame_scroll_x, orient='horizontal', command=self.text_area.xview)
    self.scroll_x.pack(side='bottom', fill='x')
    self.text_area.config(xscrollcommand=self.scroll_x.set)

    # Buttons for actions
    frame_actions = tk.Frame(self)
    frame_actions.pack(fill='x', padx=10, pady=10)

    self.add_button = tk.Button(frame_actions, text="Add Data", command=self.add_data)
    self.add_button.pack(side='left', padx=5, pady=5)

    self.update_button = tk.Button(frame_actions, text="Update Data", command=self.update_data)
    self.update_button.pack(side='left', padx=5, pady=5)

    self.delete_button = tk.Button(frame_actions, text="Delete Data", command=self.delete_data)
    self.delete_button.pack(side='left', padx=5, pady=5)
    
    # Button untuk "Refresh"
    self.refresh_button = tk.Button(frame_actions, text="Refresh Stock Barang", command=self.refresh_data)
    self.refresh_button.pack(side='left', padx=5, pady=5)
    
    # Button untuk "Export"
    self.export_button = tk.Button(frame_actions, text="Export (Save) File", command=self.export_all_sheets)
    self.export_button.pack(side='left', padx=5, pady=5)
    
    # Tambahkan entry untuk pencarian
    self.search_entry = tk.Entry(frame_actions)
    self.search_entry.pack(side='left', padx=5, pady=5)
    
    # Button untuk "Search"
    self.search_button = tk.Button(frame_actions, text="Cari", command=self.search_text)
    self.search_button.pack(side='left', padx=5, pady=5)
    
    # Tambahkan label waktu di kanan bawah
    self.time_label = tk.Label(self, text="", anchor="e")
    self.time_label.pack(side='bottom', fill='x')
    
    # Mulai perbarui waktu
    self.update_time()
    