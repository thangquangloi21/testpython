import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from pathlib import Path

class LotSummaryApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Tổng hợp LOT Sản xuất")
        self.root.geometry("1000x600")
        self.df = None
        self.summary_df = None
        
        # === PHẦN TRÊN: NÚT CHỌN FILE ===
        top_frame = ttk.Frame(root, padding="10")
        top_frame.pack(fill=tk.X)
        
        ttk.Label(top_frame, text="File Excel:").pack(side=tk.LEFT, padx=5)
        self.file_var = tk.StringVar(value=r"D:\4.DEV\KHSX\khsx.xlsx")
        file_entry = ttk.Entry(top_frame, textvariable=self.file_var, width=50)
        file_entry.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(top_frame, text="Chọn file", command=self.select_file).pack(side=tk.LEFT, padx=5)
        ttk.Button(top_frame, text="Tải dữ liệu", command=self.load_data).pack(side=tk.LEFT, padx=5)
        
        # === PHẦN LỌC ===
        filter_frame = ttk.Frame(root, padding="10")
        filter_frame.pack(fill=tk.X)
        
        ttk.Label(filter_frame, text="Lọc theo trạng thái:").pack(side=tk.LEFT, padx=5)
        self.filter_var = tk.StringVar(value="Tất cả")
        filter_combo = ttk.Combobox(filter_frame, textvariable=self.filter_var, 
                                     values=["Tất cả", "OK", "Còn nợ"], state="readonly", width=15)
        filter_combo.pack(side=tk.LEFT, padx=5)
        filter_combo.bind("<<ComboboxSelected>>", lambda e: self.apply_filter())
        
        # === PHẦN BẢNG HIỂN THỊ ===
        table_frame = ttk.Frame(root)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(table_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Treeview
        self.tree = ttk.Treeview(table_frame, yscrollcommand=scrollbar.set, height=20)
        scrollbar.config(command=self.tree.yview)
        
        # Định nghĩa cột
        self.tree['columns'] = ('LOT', 'Tổng SX', 'SLWIP', 'SLRM', 'slsx', 'Nợ PO', 'Trạng thái')
        self.tree.column('#0', width=0, stretch=tk.NO)
        self.tree.column('LOT', anchor=tk.W, width=100)
        self.tree.column('Tổng SX', anchor=tk.CENTER, width=80)
        self.tree.column('SLWIP', anchor=tk.CENTER, width=80)
        self.tree.column('SLRM', anchor=tk.CENTER, width=80)
        self.tree.column('slsx', anchor=tk.CENTER, width=80)
        self.tree.column('Nợ PO', anchor=tk.CENTER, width=80)
        self.tree.column('Trạng thái', anchor=tk.CENTER, width=100)
        
        # Heading
        self.tree.heading('#0', text='', anchor=tk.W)
        self.tree.heading('LOT', text='LOT', anchor=tk.W)
        self.tree.heading('Tổng SX', text='Tổng SX', anchor=tk.CENTER)
        self.tree.heading('SLWIP', text='SLWIP', anchor=tk.CENTER)
        self.tree.heading('SLRM', text='SLRM', anchor=tk.CENTER)
        self.tree.heading('slsx', text='slsx', anchor=tk.CENTER)
        self.tree.heading('Nợ PO', text='Nợ PO', anchor=tk.CENTER)
        self.tree.heading('Trạng thái', text='Trạng thái', anchor=tk.CENTER)
        
        self.tree.pack(fill=tk.BOTH, expand=True)
        
        # === PHẦN DƯỚI: THỐNG KÊ ===
        bottom_frame = ttk.Frame(root, padding="10")
        bottom_frame.pack(fill=tk.X)
        
        self.status_label = ttk.Label(bottom_frame, text="Sẵn sàng", foreground="blue")
        self.status_label.pack(side=tk.LEFT, padx=5)
    
    def select_file(self):
        """Chọn file Excel"""
        filename = filedialog.askopenfilename(
            title="Chọn file Excel",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.file_var.set(filename)
    
    def load_data(self):
        """Tải dữ liệu từ Excel"""
        try:
            input_file = self.file_var.get()
            self.status_label.config(text="Đang tải dữ liệu...", foreground="orange")
            self.root.update()
            
            # Đọc dữ liệu
            self.df = pd.read_excel(input_file, sheet_name='RM')
            
            # Xử lý dữ liệu
            self.summary_df = self.process_data()
            
            # Hiển thị dữ liệu
            self.display_data()
            
            self.status_label.config(
                text=f"✓ Tải thành công! {len(self.summary_df)} LOT", 
                foreground="green"
            )
        except FileNotFoundError:
            messagebox.showerror("Lỗi", f"Không tìm thấy file: {self.file_var.get()}")
            self.status_label.config(text="Lỗi: Không tìm thấy file", foreground="red")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Có lỗi xảy ra: {str(e)}")
            self.status_label.config(text=f"Lỗi: {str(e)}", foreground="red")
    
    def process_data(self):
        """Xử lý dữ liệu theo LOT"""
        if 'lotnumber' not in self.df.columns:
            raise ValueError("Sheet 'RM' không có cột 'lotnumber'!")
        
        tong_hop = pd.DataFrame()
        
        for lot in self.df['lotnumber'].unique():
            lot_df = self.df[self.df['lotnumber'] == lot].copy().reset_index(drop=True)
            
            # Tính toán
            tong_sp = lot_df['tongsanxuat'].iloc[0] if 'tongsanxuat' in lot_df.columns and not lot_df['tongsanxuat'].empty else 0
            tong_wip = lot_df['soluongwip'].iloc[0] if 'soluongwip' in lot_df.columns else 0
            tong_rm = lot_df['soluongrm'].sum() if 'soluongrm' in lot_df.columns else 0
            soluongsx = tong_wip + tong_rm
            nopo = tong_sp - tong_wip - tong_rm
            status = "OK" if nopo == 0 else "Còn nợ"
            
            tong_hop = pd.concat([tong_hop, pd.DataFrame([{
                'LOT': lot,
                'Tổng SX': int(tong_sp),
                'SLWIP': int(tong_wip),
                'SLRM': int(tong_rm),
                'slsx': int(soluongsx),
                'Nợ PO': int(nopo),
                'Trạng thái': status
            }])], ignore_index=True)
        
        return tong_hop.sort_values('LOT')
    
    def display_data(self, filter_status="Tất cả"):
        """Hiển thị dữ liệu trong bảng"""
        # Xóa dữ liệu cũ
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Lọc dữ liệu
        if filter_status == "Tất cả":
            display_df = self.summary_df
        elif filter_status == "OK":
            display_df = self.summary_df[self.summary_df['Trạng thái'] == 'OK']
        else:  # Còn nợ
            display_df = self.summary_df[self.summary_df['Trạng thái'] == 'Còn nợ']
        
        # Thêm dữ liệu vào bảng
        for idx, row in display_df.iterrows():
            tag = 'ok' if row['Trạng thái'] == 'OK' else 'nopo'
            self.tree.insert('', tk.END, values=(
                row['LOT'], row['Tổng SX'], row['SLWIP'], row['SLRM'],
                row['slsx'], row['Nợ PO'], row['Trạng thái']
            ), tags=(tag,))
        
        # Định dạng màu
        self.tree.tag_configure('ok', background='#90EE90')
        self.tree.tag_configure('nopo', background='#FFB6C6')
    
    def apply_filter(self):
        """Áp dụng bộ lọc"""
        if self.summary_df is not None:
            self.display_data(self.filter_var.get())


if __name__ == "__main__":
    root = tk.Tk()
    app = LotSummaryApp(root)
    root.mainloop()
