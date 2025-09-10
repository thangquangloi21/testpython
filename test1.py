import os
import json
import pandas as pd

# Dữ liệu JSON của máy móc đã sửa lỗi
machines_data = {
    "machines": [
        {"name": "V1", "supported_types": "4D,4G,4K,5D,5G,5K,5N,6D,6G,6K,6N,7K,7N,8K,8N", "production_per_day": 4000},
        {"name": "V2", "supported_types": "4D,4G,4K,5D,5G,5K,5N,6D,6G,6K,6N,7G", "production_per_day": 4000},
        {"name": "V3", "supported_types": "4G,4K,5G,5K,5N,6G,6K,6N,7K,7N,8K,8N", "production_per_day": 4000},
        {"name": "V4", "supported_types": "5G,5K,5N,6G,6K,6N,8K,8N", "production_per_day": 4000},
        {"name": "V5", "supported_types": "5G,5K,5N,6G,6K,6N,8K", "production_per_day": 4000},
        {"name": "V6", "supported_types": "4A,5A,6A,7D", "production_per_day": 4300},
        {"name": "V7", "supported_types": "4D,5G,5K,6D,6G", "production_per_day": 4300},
        {"name": "V8", "supported_types": "4A,5A,5G,6A,6G,7D", "production_per_day": 4300},
        {"name": "V9", "supported_types": "5D,5K,5N,6G,6K,8N", "production_per_day": 4300},
        {"name": "V10", "supported_types": "4K,5K,6G,7K", "production_per_day": 4300},
        {"name": "V11", "supported_types": "4K,5D,6N,7N", "production_per_day": 4300},
        {"name": "V12", "supported_types": "4G,6G,7G", "production_per_day": 4300},
        {"name": "V13", "supported_types": "4A,5A,6A,6G,7D", "production_per_day": 4300},
        {"name": "V14", "supported_types": "6A,6G,6K,7K,7N,8N", "production_per_day": 4300}
    ]
}

# Lấy danh sách máy
machines = machines_data['machines']

# Khởi tạo năng suất ban đầu cho mỗi máy
initial_machine_capacity = {machine['name']: machine['production_per_day'] for machine in machines}
remaining_capacity = initial_machine_capacity.copy()

# Hàm tìm máy phù hợp với năng suất còn lại
def find_suitable_machine(quennong, current_remaining_capacity, all_machines):
    for machine in all_machines:
        if quennong in machine['supported_types'].split(',') and current_remaining_capacity[machine['name']] > 0:
            return machine
    return None

# Đường dẫn file JSON sản phẩm
product_data_file = 'all_products_combined.json' 

try:
    with open(product_data_file, 'r', encoding='utf-8') as f:
        product_data = json.load(f)
    print(f"Đã đọc thành công dữ liệu sản phẩm từ file: {product_data_file}")
except FileNotFoundError:
    print(f"Lỗi: Không tìm thấy file dữ liệu sản phẩm tại đường dẫn '{product_data_file}'.")
    print("Vui lòng đảm bảo bạn đã chạy script tạo file 'all_products_combined.json' trước đó.")
    exit() 
except json.JSONDecodeError:
    print(f"Lỗi: Không thể giải mã JSON từ file '{product_data_file}'. Vui lòng kiểm tra định dạng file.")
    exit()
except Exception as e:
    print(f"Một lỗi không xác định đã xảy ra khi đọc file '{product_data_file}': {e}")
    exit()

# Phân bổ mã và chia lô theo ngày
schedule = []
daily_capacity_snapshots = [] 
day_id = 1

# Ghi lại năng suất ban đầu của ngày đầu tiên
daily_capacity_snapshots.append({
    'day_id': day_id,
    **{f"{machine['name']}_capacity": initial_machine_capacity[machine['name']] for machine in machines}
})

for product in product_data:
    quennong = product['quennong']
    tongsanxuat = product['tongsanxuat']
    batch_id_encoded = str(product['Batch_ID_encoded'])
    
    remaining_quantity = tongsanxuat
    while remaining_quantity > 0:
        suitable_machine = find_suitable_machine(quennong, remaining_capacity, machines)
        
        if not suitable_machine:
            remaining_capacity = initial_machine_capacity.copy()
            day_id += 1
            daily_capacity_snapshots.append({
                'day_id': day_id,
                **{f"{machine['name']}_capacity": initial_machine_capacity[machine['name']] for machine in machines}
            })
            suitable_machine = find_suitable_machine(quennong, remaining_capacity, machines)
            
            if not suitable_machine:
                print(f"Không tìm thấy máy phù hợp cho quennong {quennong} ngay cả sau khi chuyển sang ngày mới. Sản phẩm bị bỏ qua hoặc cần xem xét lại.")
                break 
        
        machine_capacity = suitable_machine['production_per_day']
        available_capacity = remaining_capacity[suitable_machine['name']]
        
        lot_size = min(machine_capacity, available_capacity, remaining_quantity)
        
        if lot_size > 0:
            new_record = product.copy()
            new_record['tongsanxuat_phanbo'] = lot_size 
            new_record['Batch_ID_phanbo'] = f"{product['Batch_ID']}_{day_id}" 
            new_record['Batch_ID_encoded_phanbo'] = f"{batch_id_encoded}_{day_id}" 
            new_record['assigned_machine'] = suitable_machine['name']
            new_record['day_id'] = day_id
            
            schedule.append(new_record)
            remaining_capacity[suitable_machine['name']] -= lot_size
            remaining_quantity -= lot_size
        else:
            pass 

# Chuẩn bị dữ liệu để xuất ra Excel
final_excel_data = []

for snapshot in daily_capacity_snapshots:
    final_excel_data.append(snapshot)

for record in schedule:
    original_product = next((p for p in product_data if p['stt'] == record['stt']), None)
    if original_product:
        row_data = original_product.copy()
        row_data['tongsanxuat'] = record['tongsanxuat_phanbo'] 
        row_data['Batch_ID'] = record['Batch_ID_phanbo']
        row_data['Batch_ID_encoded'] = record['Batch_ID_encoded_phanbo']
        row_data['assigned_machine'] = record['assigned_machine']
        row_data['day_id'] = record['day_id']
        final_excel_data.append(row_data)

# Xuất ra file Excel
output_file = 'production_schedule.xlsx'
df = pd.DataFrame(final_excel_data)

all_possible_columns = [
    'day_id', 'Batch_ID_encoded_phanbo', 'Batch_ID_phanbo', 'quennong', 'tongsanxuat', 'assigned_machine',
    'stt', 'thitruong', 'mathanhpham', 'soluong', 'maukhuanbandau', 'mauedotoxin',
    'mauchietxuat', 'mautinhnang', 'maueog', 'maukhac', 'Tray', 'tongmau',
    'slpalet', 'maxpalet/me', 'Batch_Total'
]
for machine in machines:
    all_possible_columns.append(f"{machine['name']}_capacity")

existing_columns = [col for col in all_possible_columns if col in df.columns]
df = df[existing_columns]

df.to_excel(output_file, index=False, sheet_name='Schedule')
print(f"Lịch trình và dữ liệu đầu vào đã được lưu vào file: {output_file}")

# In lịch trình ra màn hình
print("\n--- Lịch trình sản xuất chi tiết ---")
for record in schedule:
    print(f"Ngày ID: {record['day_id']}")
    print(f"Mã ID phân bổ: {record['Batch_ID_encoded_phanbo']}")
    print(f"Batch_ID phân bổ: {record['Batch_ID_phanbo']}")
    print(f"quennong: {record['quennong']}")
    print(f"Tổng sản xuất phân bổ: {record['tongsanxuat_phanbo']}")
    print(f"Máy được gán: {record['assigned_machine']}")
    print("---")

print("\n--- Năng suất còn lại của các máy (đơn vị/ngày) sau phân bổ cuối cùng ---")
for machine_name, capacity in remaining_capacity.items():
    print(f"Máy {machine_name}: {capacity} đơn vị")