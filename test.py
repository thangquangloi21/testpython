import pandas as pd

# Dữ liệu mẫu
data = {
    'maxpalet/me': [7, 7, 7, 7, 7, 7, 7, 7, 7, 7, 7, 7, 7, 7, 7, 7, 7, 7, 7, 7],
    'Batch_ID': ['trays760', 'trays760', 'trays761', 'trays762', 'trays763', 'trays763', 'trays764', 'trays764', 'trays766', 'trays766', 'trays766', 'trayl718', 'trayl718', 'trayl718', 'trayl718', 'pouchs72', 'pouchs73', 'pouchs74', 'pouchl71', 'pouchl71'],
    'Batch_Total': [6.93, 6.93, 6.75, 6.75, 6.89, 6.89, 6.68, 6.68, 3.52, 3.52, 3.52, 5.13, 5.13, 5.13, 5.13, 5.0, 4.55, 4.4, 0.2, 0.2],
    'tilegepme': [99, 99, 96.43, 96.43, 98.43, 98.43, 95.43, 95.43, 50.29, 50.29, 50.29, 73.29, 73.29, 73.29, 73.29, 71.43, 65.0, 62.86, 2.86, 2.86]
}
df = pd.DataFrame(data)

# Tách df1: maxpalet/me = 5 và tilegepme < 79.9, hoặc maxpalet/me = 7 và tilegepme <= 85.5
condition_df1 = ((df['maxpalet/me'] == 5) & (df['tilegepme'] < 79.9)) | ((df['maxpalet/me'] == 7) & (df['tilegepme'] <= 85.5))
df1 = df[condition_df1].copy()
df1['group'] = 'group_1'  # Thêm nhãn nhóm (tùy chọn)

# Tách df2: các hàng không thỏa mãn điều kiện của df1
df2 = df[~condition_df1].copy()
df2['group'] = 'group_2'  # Thêm nhãn nhóm (tùy chọn)

# Hiển thị kết quả
print("DataFrame df1 (thỏa mãn điều kiện):")
print(df1)
print("\nDataFrame df2 (không thỏa mãn điều kiện):")
print(df2)