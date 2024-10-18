from datetime import datetime

def update_time(self):
    # Dapatkan waktu saat ini dan format
    current_time = datetime.now().strftime("%d-%m-%Y %H:%M:%S")
    # Update label dengan waktu saat ini
    self.time_label.config(text=f"Waktu saat ini: {current_time}")
    # Panggil lagi fungsi ini setelah 1 detik
    self.after(1000, self.update_time)