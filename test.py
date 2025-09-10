import pandas as pd
import numpy as np
file_path1=r"Z:\pod_det_export.csv"
file_path2=r"Z:\po_mstr_export.csv"

df1 = pd.read_csv(file_path1)
df2 = pd.read_csv(file_path2)
# đọc 2 file csv bằng pandas trên rồi mege với nhau thành 1 df
merged_df = pd.merge(df1, df2, on='Purchase Order', how='inner')
# cộng cột Receipt Qty
merged_df['Open Qty'] = np.where(
    (merged_df['Status'] != 'X') & (merged_df['Status'] != 'C'),  # Điều kiện
    merged_df['Order Qty'] - merged_df['Receipt Qty'],               # Nếu đúng, trừ hai cột
    0 # Nếu sai, trả về 0
)

# tạo cột PO Ord Amt
merged_df["PO Ord Amt"] = merged_df["PO Ord Cost"] * merged_df["Order Qty"]

print(merged_df.head())
merged_df.to_csv('output.csv',index=False)