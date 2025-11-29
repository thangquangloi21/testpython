import oracledb
import pandas as pd
import os
from datetime import datetime

# Thiết lập thư mục Oracle Client
oracle_client_path = r"C:\oracle\23.26\instantclient_23_0"

print("=" * 50)
print("Oracle Table Counter")
print("=" * 50)

# Khởi tạo Oracle Client
try:
    oracledb.init_oracle_client(lib_dir=oracle_client_path)
    print(f"✓ Oracle Client đã khởi tạo: {oracle_client_path}\n")
except Exception as e:
    print(f"✗ Lỗi khi khởi tạo Oracle Client: {e}\n")
    exit()

# Kết nối với Oracle Database
try:
    print("Đang kết nối đến Oracle Database...")
    connection = oracledb.connect(
        user='CIMV',
        password='CIMV',
        # host='10.239.1.26',
        host='10.1.34.130',
        port=1521,
        service_name='CIMV1'
    )
    print("✓ Kết nối thành công!\n")
    
except Exception as e:
    print(f"✗ Lỗi kết nối: {e}")
    print("\nHãy kiểm tra:")
    print("1. Đường dẫn Oracle Client đúng chưa?")
    print("2. Thông tin kết nối (username, password, host, port) đúng chưa?")
    print("3. Oracle Server có chạy không?")
    exit()

try:
    # Lấy tất cả các bảng
    print("Đang lấy danh sách các bảng...")
    query = "SELECT table_name FROM user_tables ORDER BY table_name"
    cursor = connection.cursor()
    cursor.execute(query)
    tables = [row[0] for row in cursor.fetchall()]
    cursor.close()
    
    print(f"✓ Tìm thấy {len(tables)} bảng\n")
    
    # Tạo danh sách để lưu kết quả
    results = []
    total_rows = 0
    
    # Đếm dữ liệu trong mỗi bảng
    print(f"{'Tên Bảng':<40} {'Số Hàng':<15}")
    print("-" * 55)
    
    for idx, table_name in enumerate(tables, 1):
        try:
            cursor = connection.cursor()
            count_query = f"SELECT COUNT(*) FROM {table_name}"
            cursor.execute(count_query)
            count = cursor.fetchone()[0]
            cursor.close()
            
            results.append({
                'STT': idx,
                'Tên Bảng': table_name,
                'Số Hàng': count
            })
            total_rows += count
            print(f"{table_name:<40} {count:<15,}")
            
        except Exception as e:
            results.append({
                'STT': idx,
                'Tên Bảng': table_name,
                'Số Hàng': f'Lỗi: {str(e)[:20]}'
            })
            print(f"{table_name:<40} {'Lỗi':<15}")
    
    # Tạo DataFrame từ kết quả
    result_df = pd.DataFrame(results)
    
    # Xuất ra file Excel
    output_file = f'dulieu/table_counts_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
    result_df.to_excel(output_file, index=False, sheet_name='Tables')
    
    print("-" * 55)
    print(f"{'Tổng cộng':<40} {total_rows:<15,}")
    print("-" * 55)
    print(f"\n✓ Dữ liệu đã được xuất ra file: {output_file}")
    
    # Đóng kết nối
    connection.close()
    print("✓ Kết nối đã được đóng")
    print("\n" + "=" * 50)

except Exception as e:
    print(f"\n✗ Lỗi xảy ra: {e}")
    import traceback
    traceback.print_exc()
    connection.close()
finally:
    print("Chương trình kết thúc")