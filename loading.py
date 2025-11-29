# loading_form.py
import tkinter as tk
from tkinter import ttk
import threading
import time

class LoadingForm:
    def __init__(self, parent=None, title="Đang xử lý...", message="Vui lòng chờ..."):
        self.parent = parent
        self.title = title
        self.message = message
        self.root = None
        self.is_running = False

    def start(self):
        """Bắt đầu hiển thị form loading (trong thread riêng)"""
        if self.is_running:
            return
        self.is_running = True
        threading.Thread(target=self._show, daemon=True).start()

    def stop(self):
        """Đóng form loading"""
        if self.root and self.is_running:
            self.is_running = False
            self.root.quit()  # Dừng mainloop
            self.root.destroy()

    def _show(self):
        """Tạo và hiển thị cửa sổ loading"""
        self.root = tk.Toplevel() if self.parent else tk.Tk()
        self.root.title(self.title)
        self.root.geometry("320x120")
        self.root.resizable(False, False)
        self.root.configure(bg="#f0f0f0")

        # Căn giữa màn hình
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (self.root.winfo_width() // 2)
        y = (self.root.winfo_screenheight() // 2) - (self.root.winfo_height() // 2)
        self.root.geometry(f"+{x}+{y}")

        # Vô hiệu hóa cửa sổ cha (nếu có)
        if self.parent:
            self.parent.attributes('-disabled', True)
            self.root.transient(self.parent)
            self.root.grab_set()

        # Icon loading (Progressbar dạng vòng)
        frame = tk.Frame(self.root, bg="#f0f0f0")
        frame.pack(expand=True)

        self.progress = ttk.Progressbar(frame, mode='indeterminate', length=220)
        self.progress.pack(pady=15)
        self.progress.start(10)  # Tốc độ quay

        # Thông báo
        self.label = tk.Label(
            frame,
            text=self.message,
            font=("Segoe UI", 10),
            bg="#f0f0f0",
            fg="#333333"
        )
        self.label.pack(pady=5)

        # Bắt đầu vòng lặp
        self.root.protocol("WM_DELETE_WINDOW", lambda: None)  # Không cho đóng
        self.root.mainloop()

    def update_message(self, new_message):
        """Cập nhật thông báo (gọi từ thread khác)"""
        if self.label and self.is_running:
            self.root.after(0, lambda: self.label.config(text=new_message))


# ========================
# VÍ DỤ SỬ DỤNG (TEST)
# ========================
def heavy_task(loading):
    """Mô phỏng công việc nặng 5 giây"""
    for i in range(6):
        time.sleep(1)
        loading.update_message(f"Đang xử lý... {i*20}%")
    loading.stop()

def start_processing():
    # Tạo form loading
    loading = LoadingForm(root, title="Đang tải dữ liệu", message="Khởi tạo...")
    loading.start()

    # Chạy task nặng trong thread riêng
    threading.Thread(target=heavy_task, args=(loading,), daemon=True).start()

# === GIAO DIỆN CHÍNH ===
root = tk.Tk()
root.title("Ứng dụng với Loading")
root.geometry("400x200")

tk.Label(root, text="Nhấn nút để bắt đầu xử lý", font=("Arial", 12)).pack(pady=30)
tk.Button(root, text="Bắt đầu xử lý", font=("Arial", 11), command=start_processing).pack(pady=20)

root.mainloop()