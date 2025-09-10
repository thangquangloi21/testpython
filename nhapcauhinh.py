import socket

printer_ip = "10.239.9.16"  # MES-PRT-01
port = 9100
input_file = "config.zpl"

try:
    with open(input_file, 'r', encoding='ascii') as f:
        config_data = f.read()
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.settimeout(5)
        s.connect((printer_ip, port))
        s.sendall(config_data.encode('ascii'))
        print(f"Đã gửi cấu hình tới {printer_ip}")
        s.sendall(b"^XA~JC^XZ")
        print("Đã gửi lệnh hiệu chỉnh cảm biến giấy")
except FileNotFoundError:
    print(f"Lỗi: Tệp {input_file} không tồn tại")
except socket.timeout:
    print(f"Lỗi: Timeout khi kết nối tới {printer_ip}:{port}")
except socket.error as e:
    print(f"Lỗi kết nối: {e}")
except Exception as e:
    print(f"Lỗi không xác định: {e}")