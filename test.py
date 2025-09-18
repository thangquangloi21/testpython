import pandas as pd

# Dữ liệu mẫu
data = {
    'a': [5, 5, 7, 7, 5, 3, 7,7],
    'B': [50, 65, 75, 55, 70, 45, 80,4]
}
df = pd.DataFrame(data)
print("Dữ liệu gốc:")
print(df)

# Tách dữ liệu theo các điều kiện
results = {}
conditions = {
    5: df[(df['a'] == 5) & (df['B'] > 60)],
    7: df[(df['a'] == 7) & (df['B'] > 70)]
}


# Hiển thị kết quả
for value, result_df in conditions.items():
    # print(f"\nDữ liệu khi a == {value} và B > {60 if value == 5 else 70}:")
    print("sau khi lọc")
    print(result_df)