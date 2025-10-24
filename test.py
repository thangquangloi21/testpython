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