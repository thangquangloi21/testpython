import pandas as pd
import math
import numpy as np
from itertools import combinations

# Đường dẫn đến file Excel
file_path = 'Master_GW.xlsx'  # Cập nhật nếu file ở nơi khác, ví dụ: r'D:\4.DEV\TEST_python\MasterIK.xlsx'

# tính mẫu:
def tinhmau(max_lot_size,mkbd,chietxuat,mautinhnang,meog,maukhac):
    # Tổng các ô
    total = max_lot_size + mkbd + chietxuat + mautinhnang + meog + maukhac
    # print(maukhac)
    # Áp dụng logic
    if (total + 2) > 29:
        result = max(min(math.ceil(3 * total / 97), 10), 3)
    else:
        result = 2
    # print(f"The result is: {result}")
    return result


# lamtron de tinh palet
def round_up_to_divisible(n, divisor):
    return math.ceil(n / divisor) * divisor


def chiame_df_tray_l(df_tray_s, df_tray_l, pouchs, pouchl):
        # tray S
        df_trays = create_batches(df_tray_s, "trays",7)
        # tray L
        df_tray_l_5 = df_tray_l[df_tray_l['maxpalet/me'] == 5].copy()
        df_tray_l_7 = df_tray_l[df_tray_l['maxpalet/me'] == 7].copy()
        trayl5 = create_batches(df_tray_l_5, "trayl",5)
        trayl7 = create_batches(df_tray_l_7, "trayl",7)
        # gộp df tray L
        df_trayl= pd.concat([trayl5, trayl7], ignore_index=True)
        # Pouch S
        df_pouchs = create_batches(pouchs, "pouchs",7)
        # Pouch L
        df_pouchl = create_batches(pouchl, "pouchl",7)
        
        # gộp df
        dfs = [df_trays, df_trayl, df_pouchs, df_pouchl]
        df_combined = pd.concat(dfs, ignore_index=True)
        return df_combined

def create_batches(df, tray,max_slpalet):
    """
    Chia các dòng dữ liệu thành các mẻ sao cho tổng slpalet trong mỗi mẻ không vượt quá maxpalet/me,
    và thêm cột Batch_ID và Batch_Total vào DataFrame ban đầu, giữ nguyên các cột gốc.

    Args:
        df (pd.DataFrame): DataFrame chứa các cột 'mathanhpham', 'slpalet'
        max_slpalet (float): Giới hạn tổng slpalet của mỗi mẻ (mặc định là 7)

    Returns:
        pd.DataFrame: DataFrame ban đầu với thêm cột 'Batch_ID' và 'Batch_Total'
    """
    batches = []
    batch_id = 1
    remaining_items = df.sort_values(by='slpalet', ascending=False).to_dict('records')

    while remaining_items:
        current_batch = []
        current_sum = 0

        for item in remaining_items[:]:
            if current_sum + item['slpalet'] <= (max_slpalet - 0.01):
                current_batch.append(item)
                current_sum += item['slpalet']
                remaining_items.remove(item)

        if current_batch:
            for item in current_batch:
                item['Batch_ID'] = f'{tray}{max_slpalet}{batch_id}'
                item['Batch_Total'] = current_sum
                batches.append(item)
            batch_id += 1

    return pd.DataFrame(batches)



def read_excel_to_df(file_path):
     # Đọc sheet 'TEMP', từ hàng 4 (header ở index 3), cột A đến D
    df_temp = pd.read_excel(file_path, sheet_name='INPUT1', header=1, usecols='A:K')
    
    # Hiển thị dữ liệu từ sheet TEMP
    if not df_temp.empty:
        print("Dữ liệu từ sheet TEMP (từ A4 đến D4 và các hàng tiếp theo):")
        print(df_temp.head())
        print("\nTên cột của TEMP:", df_temp.columns.tolist())
    else:
        print("Không có dữ liệu trong sheet TEMP.")


    # Đọc sheet 'MASTER', từ hàng 1 (header ở index 0), cột A đến G
    df_master = pd.read_excel(file_path, sheet_name='MASTER_GW', header=1, usecols='A:H')
    
    # Hiển thị dữ liệu từ sheet MASTER
    if not df_master.empty:
        print("\nDữ liệu từ sheet MASTER (từ A1 đến G1 và các hàng tiếp theo):")
        print(df_master.head())
        print("\nTên cột của MASTER:", df_master.columns.tolist())
    else:
        print("Không có dữ liệu trong sheet MASTER.")

    merged_df = pd.merge(
    df_temp,
    df_master,
    how='left',
    left_on=['X1'],
    right_on=['Y'])


    # Hiển thị dữ liệu sau khi gộp
    if not merged_df.empty:
        print("\nDữ liệu sau khi gộp (left join dựa trên cột B):")
        print(merged_df.head())
    else:
        print("\nKhông có dữ liệu sau khi gộp.")
    
    # Lưu kết quả chia mẻ vào sheet Lot Splits
    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        merged_df.to_excel(writer, sheet_name='step1', index=False)
    print("\nĐã lưu kết quả chia mẻ vào sheet 'step1' trong file", file_path)
    
    return merged_df
    


def chialo(merged_df):
         # Kiểm tra và chuẩn bị dữ liệu cho sheet Lot Splits
        df_plan = merged_df.copy()
        # Chia mẻ theo cỡ lô max + tính mẫu
        lot_splits = []
        for _, row in df_plan.iterrows():
            index = row['X']
            thitruong = row['X0']
            code = row['X1']
            btp = row['X2']
            rm = row['X3']
            date = row['X5']
            pass_rm = row['X6']
            total_qty = int(row['X4'])
            max_lot_size = int(row['Y1'])
            chungloai = row['Y0']
            maukbd = row['X7']
            maueog = row['X8']
            maukhac = row['X9']
            mautinhnang = row['Y3']
            mauchietxuat = row['Y4']
            maunhatban = row['Y5']
            mauluutvc = row['Y6']
            print(chungloai)
            # Tính số mẻ đầy đủ và số lượng còn lại
            full_lots = total_qty // max_lot_size
            remainder = total_qty % max_lot_size
            
            # Thêm các mẻ đầy đủ
            for i in range(full_lots):
                lot_splits.append({
                    'stt' : index,
                    'solot': i,
                    'thitruong' : thitruong,
                    'mathanhpham': code,
                    'mabtp ': btp,
                    'daydan ': rm,
                    'chungloai': chungloai,
                    'soluonglot': max_lot_size,
                    'tongsoluongsx': total_qty,
                    'date': date,
                    'passrm': pass_rm,
                    'maukbd': maukbd,
                    'maueog': maueog,
                    'maukhac': maukhac,
                    'mautinhnang': mautinhnang,
                    'mauchietxuat': mauchietxuat,
                    'maunhatban': maunhatban,
                    'mauluutvc': mauluutvc,
                    'mauedotoxin': tinhmau(max_lot_size,maukbd,mauchietxuat,mautinhnang,maueog,maukhac),
                    'tongmau': tinhmau(max_lot_size,maukbd,mauchietxuat,mautinhnang,maueog,maukhac) + maukbd + maueog + maukhac + mautinhnang + mauchietxuat + maunhatban + mauluutvc   ,      
                    'tongsanxuat': max_lot_size + (tinhmau(max_lot_size,maukbd,mauchietxuat,mautinhnang,maueog,maukhac) + maukbd + maueog + maukhac + mautinhnang + mauchietxuat + maunhatban + mauluutvc)

                    # 'some': i + 1,
                    # 'soluong': max_lot_size,
                    # 'maukhuanbandau': mkbd,
                    # 'mauedotoxin': tinhmau(max_lot_size,mkbd,chietxuat,mautinhnang,meog,maukhac),
                    # 'mauchietxuat': chietxuat,
                    # 'mautinhnang': mautinhnang,
                    # 'maueog': meog,
                    # 'maukhac': maukhac,
                    # 'Tray': tray,
                    # 'tongmau': mkbd + tinhmau(max_lot_size,mkbd,chietxuat,mautinhnang,meog,maukhac)+ chietxuat+ mautinhnang+meog+maukhac,
                    # 'tongsanxuat': max_lot_size + (mkbd + tinhmau(max_lot_size,mkbd,chietxuat,mautinhnang,meog,maukhac)+ chietxuat+ mautinhnang+meog+maukhac),
                    # 'slpalet': round(round_up_to_divisible((max_lot_size + (mkbd + tinhmau(max_lot_size,mkbd,chietxuat,mautinhnang,meog,maukhac)+ chietxuat+ mautinhnang+meog+maukhac)), sltxx) / maxpalet,2),
                    # 'maxpalet/me': slmtt / maxpalet,
                    # 'quennong': quenong,
                    # 'tilechiemdung': (round(round_up_to_divisible((max_lot_size + (mkbd + tinhmau(max_lot_size,mkbd,chietxuat,mautinhnang,meog,maukhac)+ chietxuat+ mautinhnang+meog+maukhac)), sltxx) / maxpalet,2) / (slmtt / maxpalet)) * 100
                })
            
            # Thêm mẻ còn lại (nếu có)
            if remainder > 0:
                lot_splits.append({
                    'stt' : index,
                    'solot': i + 1,
                    'thitruong' : thitruong,
                    'mathanhpham': code,
                    'mabtp ': btp,
                    'daydan ': rm,
                    'chungloai': chungloai,
                    'soluonglot': remainder,
                    'tongsoluongsx': total_qty,
                    'date': date,
                    'passrm': pass_rm,
                    'maukbd': maukbd,
                    'maueog': maueog,
                    'maukhac': maukhac,
                    'mautinhnang': mautinhnang,
                    'mauchietxuat': mauchietxuat,
                    'maunhatban': maunhatban,
                    'mauluutvc': mauluutvc,
                    'mauedotoxin': tinhmau(remainder,maukbd,mauchietxuat,mautinhnang,maueog,maukhac),
                    'tongmau': tinhmau(remainder,maukbd,mauchietxuat,mautinhnang,maueog,maukhac) + maukbd + maueog + maukhac + mautinhnang + mauchietxuat + maunhatban + mauluutvc,
                    'tongsanxuat': remainder + (tinhmau(remainder,maukbd,mauchietxuat,mautinhnang,maueog,maukhac) + maukbd + maueog + maukhac + mautinhnang + mauchietxuat + maunhatban + mauluutvc)
                    
                    # 'stt' : index,
                    # 'thitruong' : thitruong,
                    # 'mathanhpham': code,
                    # 'some': full_lots + 1,
                    # 'soluong': remainder,
                    # 'maukhuanbandau': mkbd,
                    # 'mauedotoxin': tinhmau(remainder,mkbd,chietxuat,mautinhnang,meog,maukhac),
                    # 'mauchietxuat': chietxuat,
                    # 'mautinhnang': mautinhnang,
                    # 'maueog': meog,
                    # 'maukhac': maukhac,
                    # 'Tray': tray,
                    # 'tongmau': mkbd + tinhmau(remainder,mkbd,chietxuat,mautinhnang,meog,maukhac)+ chietxuat+ mautinhnang+meog+maukhac,
                    # 'tongsanxuat': remainder + (mkbd + tinhmau(remainder,mkbd,chietxuat,mautinhnang,meog,maukhac)+ chietxuat+ mautinhnang+meog+maukhac),
                    # 'slpalet': round(round_up_to_divisible((remainder + (mkbd + tinhmau(remainder,mkbd,chietxuat,mautinhnang,meog,maukhac)+ chietxuat+ mautinhnang+meog+maukhac)), sltxx) / maxpalet,2),
                    # 'maxpalet/me': slmtt / maxpalet,
                    # 'quennong': quenong,
                    # 'tilechiemdung': (round(round_up_to_divisible((remainder + (mkbd + tinhmau(remainder,mkbd,chietxuat,mautinhnang,meog,maukhac)+ chietxuat+ mautinhnang+meog+maukhac)), sltxx) / maxpalet,2) / (slmtt / maxpalet)) * 100
                })
        
        # Tạo DataFrame từ kết quả chia mẻ
        df_lot_splits = pd.DataFrame(lot_splits)
        
        # Hiển thị kết quả chia mẻ
        if not df_lot_splits.empty:
            print("\nKết quả chia mẻ theo cỡ lô max:")
            print(df_lot_splits)
        else:
            print("\nKhông có mẻ nào được chia.")

        # Lưu kết quả chia mẻ vào sheet Lot Splits
        with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df_lot_splits.to_excel(writer, sheet_name='Lot Splits', index=False)
        print("\nĐã lưu kết quả chia mẻ vào sheet 'Lot Splits' trong file", file_path)
        return df_lot_splits




def gopme(df_lot_splits):
    # # In kết quả
    # print(result)
    # Tách DataFrame thành 4 DataFrame dựa trên cột Tray
    df_tray_s = df_lot_splits[df_lot_splits['Tray'] == 'Tray S'].copy()
    df_tray_l = df_lot_splits[df_lot_splits['Tray'] == 'Tray L'].copy()
    df_pouch_s = df_lot_splits[df_lot_splits['Tray'] == 'Pouch S'].copy()
    df_pouch_l = df_lot_splits[df_lot_splits['Tray'] == 'Pouch L'].copy()

    mett = chiame_df_tray_l(df_tray_s,df_tray_l,df_pouch_s,df_pouch_l)
    
    tinhmett = mett[mett['Batch_Total'] < 4.90].copy()
    tinhmett2 = mett[mett['Batch_Total'] > 4.90].copy()
    cuoicung = create_batches(tinhmett,"gop", 7)
    
    # gộp df
    dfs = [tinhmett2,cuoicung]
    df_combined = pd.concat(dfs, ignore_index=True)
    df_combined['Batch_ID_encoded'] = pd.factorize(df_combined['Batch_ID'])[0]
    # df_pouchs1 = create_batches(tinhmett, "end",7)
#  # Lưu kết quả chia mẻ vào sheet Lot Splits
#     with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
#         cuoicung.to_excel(writer, sheet_name='meconlai', index=False)
#     print("\nĐã lưu kết quả chia mẻ vào sheet 'meconlai' trong file", file_path)
    # Lưu kết quả chia mẻ vào sheet chiame
    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df_combined.to_excel(writer, sheet_name='chiame', index=False)
    print("\nĐã lưu kết quả chia mẻ vào sheet 'chiame' trong file", file_path)
    




data1 = read_excel_to_df(file_path)
data2  = chialo(data1)
# gopme(data2)

