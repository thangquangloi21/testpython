import tkinter as tk
from tkinter import ttk
import random


class RandomGuessGame:
    def __init__(self, root):
        self.root = root
        root.title("Trò chơi đoán số — Random 1–9")
        root.geometry("500x450")
        root.resizable(False, False)

        # Style
        style = ttk.Style(root)
        try:
            style.theme_use('clam')
        except Exception:
            pass
        style.configure('TFrame', background='#f0f4ff')
        style.configure('Title.TLabel', font=('Segoe UI', 18, 'bold'), background='#f0f4ff', foreground='#1a3a52')
        style.configure('TButton', font=('Segoe UI', 10))
        style.configure('Info.TLabel', font=('Segoe UI', 11), background='#f0f4ff')

        # State
        self.target = random.randint(1, 9)
        self.attempts = 0
        self.game_over = False

        # Layout
        container = ttk.Frame(root, padding=16)
        container.pack(fill=tk.BOTH, expand=True)

        title = ttk.Label(container, text='🎮 Đoán số 1 → 9', style='Title.TLabel')
        title.pack(pady=(0, 12))

        # Info Box
        self.info_box = tk.Canvas(container, width=450, height=80, bg='#ffffff', highlightthickness=1, highlightbackground='#ccc')
        self.info_box.pack(pady=8)
        self.info_text = self.info_box.create_text(225, 25, text='Hãy đoán số mà máy đã chọn!', font=('Segoe UI', 12), fill='#333')
        self.result_text = self.info_box.create_text(225, 55, text='', font=('Segoe UI', 11, 'bold'), fill='#0066cc')

        # Input Frame
        input_frame = ttk.Frame(container)
        input_frame.pack(pady=12)
        ttk.Label(input_frame, text='Số của bạn:', style='Info.TLabel').grid(row=0, column=0, padx=8)
        self.guess_var = tk.StringVar()
        self.guess_entry = ttk.Entry(input_frame, textvariable=self.guess_var, width=5, font=('Segoe UI', 14))
        self.guess_entry.grid(row=0, column=1, padx=8)
        self.guess_entry.bind('<Return>', lambda e: self.check_guess())

        # Button Frame
        btn_frame = ttk.Frame(container)
        btn_frame.pack(pady=12)
        ttk.Button(btn_frame, text='✓ Đoán', command=self.check_guess).grid(row=0, column=0, padx=6)
        ttk.Button(btn_frame, text='🔄 Chơi lại', command=self.new_game).grid(row=0, column=1, padx=6)

        # Stats Frame
        stats_frame = ttk.Frame(container)
        stats_frame.pack(pady=12)
        self.attempts_label = ttk.Label(stats_frame, text='Lần đoán: 0', style='Info.TLabel')
        self.attempts_label.pack()

        # Footer
        footer = ttk.Label(container, text='Nhập số (1-9) và nhấn Enter hoặc nút Đoán', font=('Segoe UI', 9), background='#f0f4ff')
        footer.pack(pady=(12, 0))

    def check_guess(self):
        """Kiểm tra lần đoán của người chơi."""
        if self.game_over:
            return

        try:
            guess = int(self.guess_var.get())
            if guess < 1 or guess > 9:
                self.info_box.itemconfig(self.result_text, text='Vui lòng nhập số 1-9!', fill='#ff6600')
                return
        except ValueError:
            self.info_box.itemconfig(self.result_text, text='Nhập một số hợp lệ!', fill='#ff0000')
            return

        self.attempts += 1

        if guess == self.target:
            # Thắng
            self.game_over = True
            msg = f'🎉 Đúng rồi! Số là {self.target}.\nBạn dùng {self.attempts} lần đoán!'
            self.info_box.itemconfig(self.result_text, text=msg, fill='#00aa00')
            self.guess_entry.config(state='disabled')
        elif guess < self.target:
            msg = f'Số bạn đoán nhỏ hơn. Thử lại!'
            self.info_box.itemconfig(self.result_text, text=msg, fill='#0066cc')
        else:
            msg = f'Số bạn đoán lớn hơn. Thử lại!'
            self.info_box.itemconfig(self.result_text, text=msg, fill='#0066cc')

        self.attempts_label.config(text=f'Lần đoán: {self.attempts}')
        self.guess_var.set('')
        self.guess_entry.focus()

    def new_game(self):
        """Bắt đầu trò chơi mới."""
        self.target = random.randint(1, 9)
        self.attempts = 0
        self.game_over = False
        self.guess_var.set('')
        self.attempts_label.config(text='Lần đoán: 0')
        self.guess_entry.config(state='normal')
        self.guess_entry.focus()
        self.info_box.itemconfig(self.result_text, text='', fill='#0066cc')


if __name__ == '__main__':
    root = tk.Tk()
    app = RandomGuessGame(root)
    root.mainloop()
