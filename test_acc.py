# import tkinter as tk
# import pandas as pd
# import numpy as np

# # Đọc dữ liệu từ file Excel
# file_path = r"C:\Users\249533\Downloads\Tool 31.08.2025 (use).xlsx"
# df = pd.read_excel(file_path, sheet_name="Tools", header=4, usecols="A:DS")

# # Lọc các hàng có Stock (Amount) > 0
# filtered_df = df[df['Stock (Amount)'] > 0]

# # Lấy danh sách các giá trị duy nhất trong cột 'Dept mua' và loại bỏ nan
# unique_depts = filtered_df['Dept mua'].dropna().unique()

# # Tạo DataFrame chứa các bộ phận duy nhất
# bophan = pd.DataFrame(unique_depts, columns=['bộ phận'])

# # In danh sách các bộ phận
# print("Danh sách các bộ phận duy nhất:")
# print(bophan)

# # Tạo dictionary để lưu các DataFrame theo bộ phận
# dept_dfs = {}

# # Lọc và lưu dữ liệu theo từng bộ phận
# for _, row in bophan.iterrows():
#     dept = row['bộ phận']
#     print(f"\nĐang xử lý bộ phận: {dept}")
    
#     # Lọc dữ liệu cho bộ phận hiện tại
#     dept_data = filtered_df[filtered_df['Dept mua'] == dept]
    
#     # Lưu DataFrame vào dictionary
#     dept_dfs[dept] = dept_data
    
#     # Lưu DataFrame vào file Excel
#     output_file = f"filtered_dept_{dept}.xlsx"
#     dept_data.to_excel(output_file, index=False)
#     print(f"Đã lưu dữ liệu bộ phận {dept} vào file: {output_file}")
#     print(f"Số lượng hàng: {len(dept_data)}")

# # (Tùy chọn) In một số thông tin về các DataFrame đã tách
# print("\nThông tin các DataFrame đã tách:")
# for dept, dept_df in dept_dfs.items():
#     print(f"Bộ phận {dept}: {len(dept_df)} hàng")

import tkinter as tk
import pandas as pd
import numpy as np
import os

# Đường dẫn file Excel
file_path = r"C:\Users\249533\Downloads\Tool 31.08.2025 (use).xlsx"

# Đọc dữ liệu từ file Excel
df = pd.read_excel(file_path, sheet_name="Tools", header=4, usecols="A:DS")

# Lọc các hàng có Stock (Amount) > 0
filtered_df = df[df['Stock (Amount)'] > 0]

# Danh sách các cột cần giữ
required_columns = [
    'Ngày chứng từ', 'Số chứng từ', 'PO', 'Mã khách', 'Tên khách hàng', 
    'Budget No', 'Item code', 'Description', 'Diễn giải', 'Unit', 
    "Q'ty", 'Total (FC)', 'Cur', 'Total (LC)\n(USD)', 'G/L Account', 
    'Dist rule', 'Mục đích', 'Dept mua', 'Ngày đặt  hàng', 'Emp Code', 
    'Người đặt hàng', 'Stock (q\'ty)', 'Stock (Amount)'
]

# Kiểm tra xem các cột có tồn tại trong DataFrame không
missing_columns = [col for col in required_columns if col not in filtered_df.columns]
if missing_columns:
    print(f"Các cột không tồn tại trong DataFrame: {missing_columns}")
else:
    # Chỉ giữ các cột được yêu cầu
    filtered_df = filtered_df[required_columns]

    # Lấy danh sách các giá trị duy nhất trong cột 'Dept mua' và loại bỏ nan
    unique_depts = filtered_df['Dept mua'].dropna().unique()

    # Tạo thư mục tạm thời nếu chưa tồn tại
    temp_dir = "temp"
    if not os.path.exists(temp_dir):
        os.makedirs(temp_dir)

    # Tạo dictionary để lưu các DataFrame theo bộ phận
    dept_dfs = {}

    # Lọc và lưu dữ liệu theo từng bộ phận
    print("Danh sách các bộ phận duy nhất:")
    print(unique_depts)

    for dept in unique_depts:
        print(f"\nĐang xử lý bộ phận: {dept}")
        
        # Lọc dữ liệu cho bộ phận hiện tại
        dept_data = filtered_df[filtered_df['Dept mua'] == dept]
        
        # Lưu DataFrame vào dictionary
        dept_dfs[dept] = dept_data
        
        # Lưu DataFrame vào file Excel trong thư mục temp
        output_file = os.path.join(temp_dir, f"{dept}.xlsx")
        dept_data.to_excel(output_file, index=False)
        print(f"Đã lưu dữ liệu bộ phận {dept} vào file: {output_file}")
        print(f"Số lượng hàng: {len(dept_data)}")

    # Lưu toàn bộ dữ liệu đã lọc vào một file Excel trong thư mục temp
    all_data_file = os.path.join(temp_dir, "filtered_all_depts.xlsx")
    filtered_df.to_excel(all_data_file, index=False)
    print(f"\nĐã lưu toàn bộ dữ liệu đã lọc vào file: {all_data_file}")
    print(f"Tổng số hàng: {len(filtered_df)}")

    # (Tùy chọn) In thông tin về các DataFrame đã tách
    print("\nThông tin các DataFrame đã tách:")
    for dept, dept_df in dept_dfs.items():
        print(f"Bộ phận {dept}: {len(dept_df)} hàng")
        
