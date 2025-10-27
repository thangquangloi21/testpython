import pandas as pd
import math
import numpy as np
from itertools import combinations
from datetime import datetime, timedelta
import re

# Đường dẫn đến file Excel
file_path = 'Master_GW.xlsx'  # Cập nhật nếu file ở nơi khác, ví dụ: r'D:\4.DEV\TEST_python\MasterIK.xlsx'
file_path1 = 'inventory.xlsx'

# tính mẫu:
def tinhmau_edotoxin(max_lot_size,mkbd,chietxuat,mautinhnang,meog,maukhac):
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
    remaining_items = df.sort_values(by='tongsanxuat', ascending=False).to_dict('records')
    

    while remaining_items:
        current_batch = []
        current_sum = 0

        for item in remaining_items[:]:
            if current_sum + item['soluonglot'] <= (max_slpalet):
                current_batch.append(item)
                current_sum += item['soluonglot']
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
    df_temp = pd.read_excel(file_path, sheet_name='INPUT1', header=1, usecols='A:H')
    
    # # Hiển thị dữ liệu từ sheet TEMP
    # if not df_temp.empty:
    #     print("Dữ liệu từ sheet TEMP (từ A4 đến D4 và các hàng tiếp theo):")
    #     print(df_temp.head())
    #     print("\nTên cột của TEMP:", df_temp.columns.tolist())
    # else:
    #     print("Không có dữ liệu trong sheet TEMP.")


    # Đọc sheet 'MASTER', từ hàng 1 (header ở index 0), cột A đến G
    df_master = pd.read_excel(file_path, sheet_name='MASTER_GW', header=1, usecols='A:I')
    
    # # Hiển thị dữ liệu từ sheet MASTER
    # if not df_master.empty:
    #     print("\nDữ liệu từ sheet MASTER (từ A1 đến G1 và các hàng tiếp theo):")
    #     print(df_master.head())
    #     print("\nTên cột của MASTER:", df_master.columns.tolist())
    # else:
    #     print("Không có dữ liệu trong sheet MASTER.")

    merged_df = pd.merge(
    df_temp,
    df_master,
    how='left',
    left_on=['X1'],
    right_on=['Y'])


    # # Hiển thị dữ liệu sau khi gộp
    # if not merged_df.empty:
    #     print("\nDữ liệu sau khi gộp (left join dựa trên cột B):")
    #     print(merged_df.head())
    # else:
    #     print("\nKhông có dữ liệu sau khi gộp.")
    
    # Lưu kết quả chia mẻ vào sheet Lot Splits
    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        merged_df.to_excel(writer, sheet_name='step1', index=False)
        
    # print("\nĐã lưu kết quả chia mẻ vào sheet 'step1' trong file", file_path)
    
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
            # btp = row['X2']
            # rm = row['X3']
            # date = row['X5']
            # pass_rm = row['X6']
            wip = row['X3']
            total_qty = int(row['X4'])
            maukbd = row['X7']
            maueog = row['X8']
            maukhac = row['X9']
            chungloai = row['Y0']
            max_lot_size = int(row['Y1'])
            maxmett = row['Y2']
            mautinhnang = row['Y3']
            mauchietxuat = row['Y4']
            maunhatban = row['Y5']
            mauluutvc = row['Y6']
            maxpalet = row['Y7']
            # print(chungloai)

            # Tính số mẻ đầy đủ và số lượng còn lại
            full_lots = total_qty // max_lot_size
            remainder = total_qty % max_lot_size
            
            # Thêm các mẻ đầy đủ
            for i in range(full_lots):
                lot_splits.append({
                    'X' : index,
                    'X0' : thitruong,
                    'X1': code,
                    'X3': wip,
                    'X4': max_lot_size,
                    'X7': maukbd,
                    'X8': maueog,
                    'X9': maukhac,
                    'Y0': chungloai,
                    'Y1': max_lot_size,
                    'Y2'   : maxmett,
                    'Y3': mautinhnang,
                    'Y4': mauchietxuat,
                    'Y5': maunhatban,
                    'Y6': mauluutvc,
                    'Y7'   : maxpalet,
                    'solot': i,
                    'tongsx': total_qty
                    
                    # 'mauedotoxin': tinhmau(max_lot_size,maukbd,mauchietxuat,mautinhnang,maueog,maukhac),
                    # 'tongmau': tinhmau(max_lot_size,maukbd,mauchietxuat,mautinhnang,maueog,maukhac) + maukbd + maueog + maukhac + mautinhnang + mauchietxuat + maunhatban + mauluutvc   ,      
                    # 'tongsanxuat': max_lot_size + (tinhmau(max_lot_size,maukbd,mauchietxuat,mautinhnang,maueog,maukhac) + maukbd + maueog + maukhac + mautinhnang + mauchietxuat + maunhatban + mauluutvc)

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
                    'X' : index,
                    'X0' : thitruong,
                    'X1': code,
                    'X3': wip,
                    'X4': remainder,
                    'X7': maukbd,
                    'X8': maueog,
                    'X9': maukhac,
                    'Y0': chungloai,
                    'Y1': max_lot_size,
                    'Y2'   : maxmett,
                    'Y3': mautinhnang,
                    'Y4': mauchietxuat,
                    'Y5': maunhatban,
                    'Y6': mauluutvc,
                    'Y7'   : maxpalet,
                    'solot': i + 1,
                    'tongsx': total_qty
                    # 'stt' : index,
                    # 'thitruong' : thitruong,
                    # 'mathanhpham': code,
                    # 'chungloai': chungloai,
                    # 'soluonglot': remainder,
                    # 'maxlot': max_lot_size,
                    # 'tongsoluongsx': total_qty,
                    # 'maxmett'   : maxmett,
                    # 'maxpalet'   : maxpalet,
                    # 'maukbd': maukbd,
                    # 'maueog': maueog,
                    # 'maukhac': maukhac,
                    # 'mautinhnang': mautinhnang,
                    # 'mauchietxuat': mauchietxuat,
                    # 'maunhatban': maunhatban,
                    # 'mauluutvc': mauluutvc,
                    # 'solot': i + 1,
                    # 'mauedotoxin': tinhmau(remainder,maukbd,mauchietxuat,mautinhnang,maueog,maukhac),
                    # 'tongmau': tinhmau(remainder,maukbd,mauchietxuat,mautinhnang,maueog,maukhac) + maukbd + maueog + maukhac + mautinhnang + mauchietxuat + maunhatban + mauluutvc,
                    # 'tongsanxuat': remainder + (tinhmau(remainder,maukbd,mauchietxuat,mautinhnang,maueog,maukhac) + maukbd + maueog + maukhac + mautinhnang + mauchietxuat + maunhatban + mauluutvc)

                    
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
        # df_lot_splits["INDEX"] = range(1, len(df_lot_splits) + 1)
        # df_lot_splits['lot_id'] = df_lot_splits['INDEX'].astype(str) + df_lot_splits['stt'].astype(str) + df_lot_splits['solot'].astype(str)
        # df_lot_splits['chiempalett'] = round(df_lot_splits['tongsanxuat'] / df_lot_splits['maxpalet'],2)
        

        # # Hiển thị kết quả chia mẻ
        # if not df_lot_splits.empty:
        #     print("\nKết quả chia mẻ theo cỡ lô max:")
        #     print(df_lot_splits)
        # else:
        #     print("\nKhông có mẻ nào được chia.")

        # Lưu kết quả chia mẻ vào sheet Lot Splits
        # with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        #     df_lot_splits.to_excel(writer, sheet_name='chialo', index=False)
        # print("\nĐã lưu kết quả chia mẻ vào sheet 'chialo' trong file", file_path)
        return df_lot_splits
# tinhmau(remainder,mkbd,chietxuat,mautinhnang,meog,maukhac),

def tinhmau(data):
    df_plan = data.copy()
    # Áp dụng hàm cho từng hàng của df_plan
    df_plan['mauedotoxin'] = df_plan.apply(
        lambda row: tinhmau_edotoxin(
            row['X4'],  # max_lot_size
            row['X7'],  # mkbd
            row['Y4'],  # chietxuat
            row['Y3'],  # mautinhnang
            row['X8'],  # meog
            row['X9']   # maukhac
        ),
        axis=1
    )
    
    # tính tổng mẫu, và tổng sản xuất
    df_plan['tongmau'] = df_plan['X7'] + df_plan['X8'] + df_plan['X9'] + df_plan['Y3'] + df_plan['Y4'] + df_plan['Y5'] + df_plan['Y6'] + df_plan['mauedotoxin']
    df_plan['tongsanxuat'] = df_plan['tongmau'] + df_plan['X4']
    
    
    # df_plan['danhsolo'] = df_plan.index + 1
    # tỉ lệ max mẻ:
    df_plan['tilechiemlot'] = round((df_plan['tongsanxuat'] / df_plan['Y1']) * 100, 2)
    return df_plan
    

def gopme(df_lot_splits):
    # # In kết quả
    # tách mẫu ra khỏi sl sản xuất
    data_sample = df_lot_splits[['lot_id', 'mathanhpham' , 'tongmau']]
    #  tách theo chủng loại
    df_gen_d = df_lot_splits[df_lot_splits["chungloai"] == "D"]
    df_gen_N = df_lot_splits[df_lot_splits["chungloai"] == "N"]
    #  gộp mẻ theo chủng loại
    gop_dai = create_batches(df_gen_d, "D", 13800)
    gop_ngan = create_batches(df_gen_N, "N", 27500)

    # gộp df
    dfs = [gop_dai, gop_ngan]   

#     # print(result)
#     # Tách DataFrame thành 4 DataFrame dựa trên cột Tray
#     df_tray_s = df_lot_splits[df_lot_splits['Tray'] == 'Tray S'].copy()
#     df_tray_l = df_lot_splits[df_lot_splits['Tray'] == 'Tray L'].copy()
#     df_pouch_s = df_lot_splits[df_lot_splits['Tray'] == 'Pouch S'].copy()
#     df_pouch_l = df_lot_splits[df_lot_splits['Tray'] == 'Pouch L'].copy()

#     mett = chiame_df_tray_l(df_tray_s,df_tray_l,df_pouch_s,df_pouch_l)
    
#     tinhmett = mett[mett['Batch_Total'] < 4.90].copy()
#     tinhmett2 = mett[mett['Batch_Total'] > 4.90].copy()
#     cuoicung = create_batches(tinhmett,"gop", 7)
    
#     # gộp df
#     dfs = [tinhmett2,cuoicung]
    df_combined = pd.concat(dfs, ignore_index=True)
#     df_combined['Batch_ID_encoded'] = pd.factorize(df_combined['Batch_ID'])[0]
#     # df_pouchs1 = create_batches(tinhmett, "end",7)
# #  # Lưu kết quả chia mẻ vào sheet Lot Splits
# #     with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
# #         cuoicung.to_excel(writer, sheet_name='meconlai', index=False)
# #     print("\nĐã lưu kết quả chia mẻ vào sheet 'meconlai' trong file", file_path)
    # Lưu kết quả chia mẻ vào sheet chiame
    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df_combined.to_excel(writer, sheet_name='gepme', index=False)
    print("\nĐã lưu kết quả chia mẻ vào sheet 'gepme' trong file", file_path)
    
def khongchialo_func(khongchialo):
    # xu ly bang khong chia lo
    khongchialo['solot'] = '0'
    khongchialo['tongsx'] = khongchialo['X4']
    khongchialo = khongchialo.drop('Y', axis= 1)
    tinhmaudone = tinhmau(khongchialo)
    # đánh số lot
    # tinhmaudone['lot_number'] = ['LOT{}'.format(i) for i in range(1, len(tinhmaudone) + 1)]
     # Lưu kết quả chia mẻ vào sheet khongchialo
    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        tinhmaudone.to_excel(writer, sheet_name='tinhmau_khongchialo', index=False)
    return tinhmaudone


def phaichialo_func(phaichialo):
    
    print(len(phaichialo))
    # chia lô
    chialodone  = chialo(phaichialo)
    # tính mẫu
    tinhmaudone = tinhmau(chialodone)
    # đánh số lot
    # tinhmaudone['lot_number'] = ['LOTS{}'.format(i) for i in range(1, len(tinhmaudone) + 1)]
    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        tinhmaudone.to_excel(writer, sheet_name='tinhmau_chialo', index=False)
    return tinhmaudone

       
       
def show_(iventory):
    if not iventory.empty:
        print("Dữ liệu từ sheet TEMP (từ A4 đến D4 và các hàng tiếp theo):")
        print(iventory.head())
        print("\nTên cột của TEMP:", iventory.columns.tolist())
    else:
        print("Không có dữ liệu trong sheet TEMP.")
        
def laytonkho(item,bom, inventory):
    # lấy item con của mã
    get = bom['component'][bom['item'] == item]
    print(get)
    
    # kiểm tra tồn của sản phẩm
    inven = inventory[inventory['masanpham'].isin(get)]
    # print(inven)
    # print(len(inven))
    
    tonkho = inven.copy()  # Tạo bản sao độc lập
    tonkho['ngayhethan'] = pd.to_datetime(tonkho['ngayhethan'], format='%d/%m/%Y')
    # Sắp xếp DataFrame theo cột ngayhethan, giảm dần (ngày gần nhất trước)
    tonkho = tonkho.sort_values(by='ngayhethan', ascending=True)
    return tonkho #trả về lượng tồn của sản phẩm


def load_bom_inven():
    
    inventory = pd.read_excel(file_path1, sheet_name='inventory', header=0, usecols='A:G')
    bom = pd.read_excel(file_path1, sheet_name='BOM', header=0, usecols='A:C')
    # lấy dữ liệu tồn và bom
    return bom, inventory
    
    
def laylot(df, bom, inventory):
    lotrm = []
    for _, row in df.iterrows():
        itemcode = row['itemncode']
        lotnumber = row['lot_number']
        wip = row['banthanhpham']
        tongsanxuat = row['tongsanxuat']
        wipdalay = row['soluonglay']
        soconlai = row['soconlai']
        wiplot = row['wiplot']
        ngayhethanwip = row["ngayhethan"]
        rm = row['component']
        
        # print(f"{itemcode} + {lotnumber} + {tongsanxuat} + {get}")
        # lấy tồn kho của wip theo item
        tonkho = laytonkho(itemcode, bom, inventory)
        # Chuyển cột ngayhethan thành định dạng datetime

        print(tonkho)
        if (len(tonkho) == 0):
            pass
        else:
            # Lặp qua các hàng của tonkho
            for _, row1 in tonkho.iterrows():
                # print(row['stt'], row['masanpham'], row['solo'], row['soluong'], row['ngayhethan'], row['ngaypass'], row['kieu'])
                # print(f"tongsanxuat truoc khi tru {tongsanxuat}")
                # tonkho = laytonkho(itemcode, bom, inventory)
                wipitem = row1['masanpham']
                wiplot = row1['solo']
                wipsoluong = row1['soluong']
                ngayhethan = row1['ngayhethan']
                ngaypass = row1['ngaypass']
                kieu = row1['kieu']
                # print(f"tongsanxuat {itemcode} sau khi tru {tongsanxuat}")
                soconlai = tongsanxuat - wipsoluong
                tonkhodu = tongsanxuat % wipsoluong
                # wipsoluong = wipsoluong - wipsoluong
                # xóa sau khi lấy
                if soconlai > 0:
                    lotrm.append(
                        {
                            'itemncode': itemcode,
                            'lotnumber': lotnumber,
                            'banthanhpham': wip,
                            'tongsanxuat': tongsanxuat,
                            'soluongsudung': wipsoluong,
                            'soconlai': soconlai,
                            'wipitem': wipitem,
                            'wiplot': wiplot,
                            'tonkho': wipsoluong,
                            'ngayhethan': ngayhethan,
                            'ngaypass': ngaypass,
                            'kieu': kieu
                        })
                    # xong thì xóa khỏi df
                    inventory = inventory.drop(inventory[(inventory['masanpham'] == wipitem) & (inventory['solo'] == wiplot)].index)
                
                if soconlai <= 0:
                    lotrm.append(
                    {
                        'itemncode': itemcode,
                        'lotnumber': lotnumber,
                        'banthanhpham': wip,
                        'tongsanxuat': tongsanxuat,
                        'soluongsudung': tonkhodu,
                        'soconlai': tongsanxuat - tonkhodu,
                        'wipitem': wipitem,
                        'wiplot': wiplot,
                        'tonkho': wipsoluong,
                        'ngayhethan': ngayhethan,
                        'ngaypass': ngaypass,
                        'kieu': kieu
                    })
                    # cập nhật tồn vào inventory
                    inventory.loc[(inventory['masanpham'] == wipitem) & (inventory['solo'] == wiplot), 'soluong'] = wipsoluong - tonkhodu
                    
    return pd.DataFrame(lotrm),inventory


# lấy ngày hết hạn của btp
def get_wip(df, inventory):
    """
    - Duyệt từng dòng nhu cầu (df): lấy BTP 'wip' và số lượng 'tongsanxuat'
    - Lấy tồn của 'wip' trong inventory theo FEFO (ngày hết hạn sớm trước)
    - Xuất các dòng đã lấy (mỗi lô một dòng), đồng thời cập nhật inventory:
        * Nếu lấy hết lô: drop lô khỏi inventory
        * Nếu lấy một phần: giảm 'soluong' của lô đó
    Trả về:
        (DataFrame kết quả, inventory đã cập nhật)
    """
    out_rows = []

    # ✅ Khuyến nghị: parse ngày 1 lần trước khi gọi hàm (nhanh hơn)
    # Nhưng nếu cần parse ở đây, dùng .loc trên inventory để tránh cảnh báo
    if inventory['ngayhethan'].dtype == object:
        inventory = inventory.copy()
        inventory.loc[:, 'ngayhethan'] = pd.to_datetime(
            inventory['ngayhethan'], format='%d/%m/%Y', errors='coerce'
        )

    for _, r in df.iterrows():
        itemcode    = r['X1']
        lotnumber   = r['lot_number']
        wip         = r['X3']              # mã bán thành phẩm cần dùng
        demand      = int(r['tongsanxuat'])  # nhu cầu cần lấy

        # Lấy các lô tồn của đúng mã BTP, sort FEFO (hết hạn sớm trước)
        inven = inventory.loc[inventory['masanpham'] == wip, :].copy()
        if inven.empty:
            # Không có tồn để cấp
            out_rows.append({
                'itemncode': itemcode,
                'lotnumber': lotnumber,
                'tongsanxuat': demand,
                'banthanhpham': wip,
                'wiplot': '',
                'soluongwip': 0,
                'ngayhethan': pd.NaT,
                'ngaypass': '',
                'kieu': 'WIP',
                'tongsxconlai': demand,
            })
            continue

        inven = inven.sort_values(by='ngayhethan', ascending=True)  # FEFO

        remaining = demand
        for _, inv in inven.iterrows():
            if remaining <= 0:
                break

            wipitem     = inv['masanpham']
            wiplot      = inv['solo']
            lot_qty     = int(inv['soluong'])
            ngayhethan  = inv['ngayhethan']
            ngaypass    = inv.get('ngaypass', '')
            kieu        = inv.get('kieu', 'WIP')

            take = min(remaining, lot_qty)   # ✅ số lấy thực sự từ lô này
            remaining_after = remaining - take

            out_rows.append({
                'itemncode': itemcode,
                'lotnumber': lotnumber,
                'tongsanxuat': demand,
                'banthanhpham': wip,
                'wiplot': wiplot,
                'soluongwip': take,
                'ngayhethan': ngayhethan,
                'ngaypass': ngaypass,
                'kieu': kieu,
                'tongsxconlai': remaining_after,
            })

            # ✅ Cập nhật inventory bằng .loc để tránh SettingWithCopyWarning
            mask = (inventory['masanpham'] == wipitem) & (inventory['solo'] == wiplot)
            if take == lot_qty:
                # Lấy hết lô → drop hàng đó
                inventory = inventory.drop(index=inventory.loc[mask].index)
            else:
                # Lấy một phần → giảm số lượng
                inventory.loc[mask, 'soluong'] = lot_qty - take

            remaining = remaining_after

        # # Nếu vẫn thiếu sau khi đi hết các lô
        # if remaining > 0:
        #     out_rows.append({
        #         'itemncode': itemcode,
        #         'lotnumber': lotnumber,
        #         'banthanhpham': wip,
        #         'tongsanxuat': demand,
        #         'soluonglay': 0,
        #         'soconlai': remaining,
        #         'wiplot': '',
        #         'ngayhethan': pd.NaT,
        #         'ngaypass': '',
        #         'kieu': 'WIP'
        #     })

    return pd.DataFrame(out_rows), inventory

def parse_lot_number(lot):
    """
    Tách số lô thành (phần chữ, phần số) để sắp xếp đúng
    VD: 'Y2508172' -> ('Y', 2508172)
        '250901' -> ('', 250901)
    """
    match = re.match(r'([A-Z]*)(\d+)', str(lot))
    if match:
        prefix = match.group(1)  # Phần chữ
        number = int(match.group(2))  # Phần số
        return (prefix, number)
    return (str(lot), 0)

def get_rm(df, inventory):
    out_rows = []
    
    # Parse date một lần
    if inventory['ngayhethan'].dtype == object:
        inventory = inventory.copy()
        inventory.loc[:, 'ngayhethan'] = pd.to_datetime(
            inventory['ngayhethan'], format='%d/%m/%Y', errors='coerce'
        )
    
    for _, row in df.iterrows():
        itemcode = row['itemncode']
        lotnumber = row['lotnumber']
        tongsanxuat = row['tongsanxuat']
        wip = row['banthanhpham']
        wiplot = row['wiplot']
        wipdalay = row['soluongwip']
        ngayhethanwip = pd.to_datetime(row["ngayhethan"], format='%d/%m/%Y', errors='coerce')
        soconlai = row['tongsxconlai']
        rm = row['component']
        
        # ✅ Dùng biến tạm thay vì ghi đè inventory
        available_lots = inventory.loc[inventory['masanpham'] == rm, :].copy()
        
        if available_lots.empty:
            out_rows.append({
                'itemncode': itemcode,
                'lotnumber': lotnumber,
                'tongsanxuat': tongsanxuat,
                'banthanhpham': wip,
                'wiplot': wiplot,
                'soluongwip': wipdalay,
                'ngayhethanwip': ngayhethanwip,
                'soconlai': soconlai,
                'component': rm,
                'lotrm': '',
                'soluongrm': 0,
                'ngayhethanrm': '',
                'ngaypassrm': '',
                'tonconlai': soconlai,
                'status': 'Khongcoton'
            })
            continue
        
        # ✅ Sắp xếp FEFO: ngày hết hạn trước, nếu trùng thì sắp xếp theo số lô
        available_lots['_sort_key'] = available_lots['solo'].apply(parse_lot_number)
        available_lots = available_lots.sort_values(
            by=['ngayhethan', '_sort_key'], 
            ascending=[True, True]
        )
        remaining = soconlai
        
        for _, inv in available_lots.iterrows():
            if remaining <= 0:
                break
            
            rmitem = inv['masanpham']
            rmlot = inv['solo']
            lot_qtyrm = int(inv['soluong'])
            ngayhethanrm = inv['ngayhethan']
            ngaypassrm = pd.to_datetime(inv.get('ngaypass', ''), format='%d/%m/%Y', errors='coerce')
            tonconlai = remaining
            take = min(remaining, lot_qtyrm)
            remaining -= take
            
            # Kiểm tra điều kiện ngày
            if pd.notna(ngaypassrm) and pd.notna(ngayhethanwip):
                if ngaypassrm + timedelta(days=4) > ngayhethanwip:
                    if wipdalay > 10:
                        status = 'nopo'
                        take = 0
                        ngayhethanrm = None
                        rmlot = None
                        ngaypassrm = None
                        
                    else:
                        # take = min(remaining, tongsanxuat)
                        take = tonconlai + wipdalay  # Lấy theo tongsanxuat
                        # remaining = 
                        status = 'OK Bỏ tồn WIP'  # Cập nhật status nếu lấy được
                else:
                    status = 'OK'
            else:
                status = 'OK'
            
            out_rows.append({
                'itemncode': itemcode,
                'lotnumber': lotnumber,
                'tongsanxuat': tongsanxuat,
                'banthanhpham': wip,
                'wiplot': wiplot,
                'soluongwip': wipdalay,
                'ngayhethanwip': ngayhethanwip,
                'soconlai': soconlai,
                'component': rm,
                'lotrm': rmlot,
                'soluongrm': take,
                'ngayhethanrm': ngayhethanrm,
                'ngaypassrm': ngaypassrm,
                'tonconlai': remaining,
                'status': status
            })
            
            # ✅ Cập nhật inventory gốc
            mask = (inventory['masanpham'] == rmitem) & (inventory['solo'] == rmlot)
            if take == lot_qtyrm:
                inventory = inventory.drop(index=inventory.loc[mask].index)
            else:
                inventory.loc[mask, 'soluong'] -= take
        
        # Nếu vẫn thiếu sau khi phân bổ
        if remaining > 0:
            out_rows.append({
                'itemncode': itemcode,
                'lotnumber': lotnumber,
                'tongsanxuat': tongsanxuat,
                'banthanhpham': wip,
                'wiplot': wiplot,
                'soluongwip': wipdalay,
                'ngayhethanwip': ngayhethanwip,
                'soconlai': soconlai,
                'component': rm,
                'lotrm': '',
                'soluongrm': 0,
                'ngayhethanrm': '',
                'ngaypassrm': '',
                'tonconlai': remaining,
                'status': 'nopo'
            })
    
    return pd.DataFrame(out_rows), inventory




# def get_rm(df, inventory):
    out_rows = []
    
    # Parse date một lần
    if inventory['ngayhethan'].dtype == object:
        inventory = inventory.copy()
        inventory.loc[:, 'ngayhethan'] = pd.to_datetime(
            inventory['ngayhethan'], format='%d/%m/%Y', errors='coerce'
        )
    
    for _, row in df.iterrows():
        itemcode = row['itemncode']
        lotnumber = row['lotnumber']
        tongsanxuat = row['tongsanxuat']
        wip = row['banthanhpham']
        wiplot = row['wiplot']
        wipdalay = row['soluongwip']
        ngayhethanwip = row["ngayhethan"]
        soconlai = row['tongsxconlai']
        rm = row['component']
        
        # ✅ Dùng biến tạm thay vì ghi đè inventory
        available_lots = inventory.loc[inventory['masanpham'] == rm, :].copy()
        
        if available_lots.empty:
            out_rows.append({
                'itemncode': itemcode,
                'lotnumber': lotnumber,
                'tongsanxuat': tongsanxuat,
                'banthanhpham': wip,
                'wiplot': wiplot,
                'soluongwip': wipdalay,
                'ngayhethanwip': ngayhethanwip,
                'soconlai': soconlai,
                'component': rm,
                'lotrm': '',
                'soluongrm': 0,
                'ngayhethanrm': '',
                'ngaypassrm': '',
                'tonconlai': soconlai,
                'status': 'Khongcoton'
            })
            continue
         # ✅ Sắp xếp FEFO: ngày hết hạn trước, nếu trùng thì sắp xếp theo số lô
        available_lots['_sort_key'] = available_lots['solo'].apply(parse_lot_number)
        available_lots = available_lots.sort_values(
            by=['ngayhethan', '_sort_key'], 
            ascending=[True, True]
        )
        available_lots = available_lots.sort_values(by='ngayhethan', ascending=True)
        remaining = soconlai
        
        for _, inv in available_lots.iterrows():
            if remaining <= 0:
                break
            
            rmitem = inv['masanpham']
            rmlot = inv['solo']
            lot_qtyrm = int(inv['soluong'])
            ngayhethanrm = inv['ngayhethan']
            ngaypassrm = inv.get('ngaypass', '')
            
            take = min(remaining, lot_qtyrm)
            remaining -= take
            
            
             # kiểm tra xem ngày (ngaypassrm) pass của (rm) + 4 ngày có lớn hơn ngày hết hạn ngayhethanwip của banthanhpham không
            
            # nếu lớn hơn thì xem tồn cần lấy của wip nếu (soluongwip > 10 => nợ po) nếu soluongwip nhỏ hơn 10 thì bỏ wip lấy số lượng theo tổng sản xuất tongsanxuat
            out_rows.append({
                'itemncode': itemcode,
                'lotnumber': lotnumber,
                'tongsanxuat': tongsanxuat,
                'banthanhpham': wip,
                'wiplot': wiplot,
                'soluongwip': wipdalay,
                'ngayhethanwip': ngayhethanwip,
                'soconlai': soconlai,
                'component': rm,
                'lotrm': rmlot,
                'soluongrm': take,
                'ngayhethanrm': ngayhethanrm,
                'ngaypassrm': ngaypassrm,
                'tonconlai': remaining,
                'status': 'OK'
            })
            
            # ✅ Cập nhật inventory gốc
            mask = (inventory['masanpham'] == rmitem) & (inventory['solo'] == rmlot)
            if take == lot_qtyrm:
                inventory = inventory.drop(index=inventory.loc[mask].index)
            else:
                inventory.loc[mask, 'soluong'] -= take
        
        # Nếu vẫn thiếu
        if remaining > 0:
            # kiểm tra xem ngày (ngaypassrm) pass của (rm) + 4 ngày có lớn hơn ngày hết hạn ngayhethanwip của banthanhpham không
            
            # nếu lớn hơn thì xem tồn cần lấy của wip nếu (soluongwip > 10 => nợ po) nếu soluongwip nhỏ hơn 10 thì bỏ wip lấy số lượng theo tổng sản xuất tongsanxuat
            out_rows.append({
                'itemncode': itemcode,
                'lotnumber': lotnumber,
                'tongsanxuat': tongsanxuat,
                'banthanhpham': wip,
                'wiplot': wiplot,
                'soluongwip': wipdalay,
                'ngayhethanwip': ngayhethanwip,
                'soconlai': soconlai,
                'component': rm,
                'lotrm': '',
                'soluongrm': 0,
                'ngayhethanrm': '',
                'ngaypassrm': '',
                'tonconlai': remaining,  # ✅ Dùng remaining thay vì remaining_after
                'status': 'nopo'
            })
    
    return pd.DataFrame(out_rows), inventory




startdate = "01/11/2025"
enddate = "15/11/2025"
ngaydatkhpass = datetime.strptime(enddate, "%d/%m/%Y")
ngaydatkhpass = ngaydatkhpass - timedelta(days = 4)
ngaydatkh = ""



# Đọc dữ liệu từ file Excel
data1 = read_excel_to_df(file_path)
# lấy những mã không cần chia lô ra:
khongchialo = data1[data1['X4'] < data1['Y1']].copy()
phaichialo = data1[data1['X4'] >  data1['Y1']].copy()


# tính mẫu
phaichialo = phaichialo_func(phaichialo)
khongchialo = khongchialo_func(khongchialo)

lotle = phaichialo[phaichialo['tilechiemlot'] < 100].copy()
phaichialo = phaichialo[phaichialo['tilechiemlot'] >= 100].copy()

goplole = [khongchialo, lotle]
khongchialo1 = pd.concat(goplole, ignore_index=True)

# đánh số mẻ gán label
phaichialo['lot_number'] = ['LOT{}'.format(i) for i in range(1, len(phaichialo) + 1)]
phaichialo['typelot'] = 'full'

khongchialo1['lot_number'] = ['LOTS{}'.format(i) for i in range(1, len(khongchialo1) + 1)]
khongchialo1['typelot'] = 'gep'
# lấy những lô không đạt 100% ra để gép với không chia lô

# load bom và inventory
bom, inventory = load_bom_inven()

# gộp dữ liệu lại
dfs = [phaichialo, khongchialo1]
gopdata = pd.concat(dfs, ignore_index=True)

#lấy tồn WIP và ngày hết hạn của wip
datewip,inve = get_wip(gopdata, inventory)

# lấy item theo bom rồi gắn vào df
datewip = datewip.merge(bom[['item', 'component']], left_on='banthanhpham', right_on='item', how='left')
# xóa cột item thừa
datewip = datewip.drop('item', axis=1)

# lấy thông tin rm
# lọc những mã có ngày pass trước end date 4 ngày:
# Lọc các hàng có ngayhethan trước 03/11/2025

inve['ngaypass'] = pd.to_datetime(inve['ngaypass'], format='%d/%m/%Y', errors='coerce')
inve = inve[(inve['ngaypass'] <= ngaydatkhpass) | (inve['ngaypass'].isna())]
inve = inve[inve['kieu'] == 'RM']

with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        inve.to_excel(writer, sheet_name='tonnvltruoc', index=False)

# lấy nvl
getrm,inve = get_rm(datewip, inve)


with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        getrm.to_excel(writer, sheet_name='getrm', index=False)

with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        gopdata.to_excel(writer, sheet_name='gopdata', index=False)

with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        datewip.to_excel(writer, sheet_name='datewip', index=False)

with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        inve.to_excel(writer, sheet_name='tonnvlsau', index=False)

