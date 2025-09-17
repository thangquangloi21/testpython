import psutil
import socket

def get_hostname():
    """Lấy tên máy tính (hostname)."""
    try:
        hostname = socket.gethostname()
        return hostname
    except Exception as e:
        return f"Lỗi khi lấy tên máy tính: {str(e)}"

def get_wifi_mac_address():
    """Lấy địa chỉ MAC của card WiFi."""
    try:
        # Lấy thông tin tất cả các giao diện mạng
        interfaces = psutil.net_if_addrs()
        stats = psutil.net_if_stats()

        # Duyệt qua các giao diện mạng
        for interface, addrs in interfaces.items():
            # Kiểm tra xem giao diện có phải là WiFi (thường chứa 'Wi-Fi', 'wlan', hoặc 'Wireless')
            if 'Wi-Fi' in interface or 'wlan' in interface.lower() or 'Wireless' in interface:
                # Kiểm tra xem giao diện có đang hoạt động
                if interface in stats and stats[interface].isup:
                    for addr in addrs:
                        # Lấy địa chỉ MAC (family = AF_LINK hoặc AF_PACKET)
                        if addr.family == psutil.AF_LINK:
                            return addr.address
        return "Không tìm thấy giao diện WiFi hoạt động."
    except Exception as e:
        return f"Lỗi khi lấy địa chỉ MAC WiFi: {str(e)}"

def main():
    print("Thông tin hệ thống:")
    print(f"Tên máy tính: {get_hostname()}")
    print(f"Địa chỉ MAC WiFi: {get_wifi_mac_address()}")

if __name__ == "__main__":
    main()