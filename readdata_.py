import os
import tkinter as tk
from tkinter import simpledialog, filedialog, messagebox

def read_files_in_directory(directory_path, output_file):
    """
    Đọc tất cả file trong thư mục và ghi thông tin vào file txt.
    Args:
        directory_path: Đường dẫn đến thư mục cần đọc.
        output_file: Đường dẫn đến file txt đầu ra.
    """
    try:
        with open(output_file, 'w', encoding='utf-8') as outfile:
            # Duyệt qua tất cả file và thư mục con
            for root, dirs, files in os.walk(directory_path):
                # Ghi thông tin thư mục
                outfile.write(f"\n=== Thư mục: {root} ===\n")
                
                # Nếu không có file trong thư mục
                if not files:
                    outfile.write("Không có file trong thư mục này.\n")
                
                # Duyệt qua từng file
                for file in files:
                    file_path = os.path.join(root, file)
                    outfile.write(f"\nTên file: {file}\n")
                    outfile.write(f"Đường dẫn: {file_path}\n")
                    
                    # Thử đọc nội dung file nếu là file văn bản
                    try:
                        with open(file_path, 'r', encoding='utf-8') as infile:
                            content = infile.read()
                            outfile.write("Nội dung:\n")
                            outfile.write(content + "\n")
                    except (UnicodeDecodeError, PermissionError, IOError):
                        outfile.write("Không thể đọc nội dung (có thể là file không phải văn bản hoặc lỗi quyền truy cập).\n")
                    outfile.write("-" * 50 + "\n")
                    
        return f"Đã ghi thông tin vào {output_file}"
    
    except Exception as e:
        return f"Lỗi: {str(e)}"

def get_user_input():
    """
    Hiển thị hộp thoại để nhập đường dẫn thư mục và file đầu ra.
    Returns:
        tuple: (directory_path, output_file) hoặc (None, None) nếu người dùng hủy.
    """
    root = tk.Tk()
    root.withdraw()  # Ẩn cửa sổ chính

    # Hiển thị hộp thoại chọn thư mục
    directory_path = filedialog.askdirectory(
        title="Chọn thư mục cần đọc"
    )
    
    # Kiểm tra nếu người dùng hủy
    if not directory_path:
        messagebox.showwarning("Cảnh báo", "Bạn chưa chọn thư mục!")
        root.destroy()
        return None, None

    # Hiển thị hộp thoại nhập tên file đầu ra
    output_file = simpledialog.askstring(
        title="Nhập tên file đầu ra",
        prompt="Nhập tên file đầu ra (ví dụ: output.txt):"
    )
    
    # Kiểm tra nếu người dùng hủy hoặc không nhập
    if not output_file:
        messagebox.showwarning("Cảnh báo", "Bạn chưa nhập tên file đầu ra!")
        root.destroy()
        return None, None
    
    # Đảm bảo file đầu ra có đuôi .txt
    if not output_file.endswith('.txt'):
        output_file += '.txt'

    root.destroy()
    return directory_path, output_file

if __name__ == "__main__":
    # Lấy đường dẫn thư mục và file đầu ra từ hộp thoại
    directory_path, output_file = get_user_input()
    
    # Kiểm tra nếu người dùng không hủy
    if directory_path and output_file:
        result = read_files_in_directory(directory_path, output_file)
        # Hiển thị thông báo kết quả
        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo("Kết quả", result)
        root.destroy()