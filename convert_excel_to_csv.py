import pandas as pd

# Đọc file Excel
excel_file = r"C:\Users\249533\Downloads\bom1233.xlsx" # Thay bằng đường dẫn file Excel của bạn
df = pd.read_excel(excel_file)

# Chuyển sang file CSV
csv_file = 'bom_mes.csv'  # Tên file CSV đầu ra
df.to_csv(csv_file, index=False, encoding='utf-8')

print(f"Đã chuyển file {excel_file} sang {csv_file}")


