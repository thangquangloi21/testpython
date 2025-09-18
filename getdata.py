import psutil
import time
import py3nvml

# Khởi tạo NVML cho GPU
gpu_available = False
try:
    py3nvml.nvmlInit()
    gpu_available = True
except Exception as e:
    print(f"Không phát hiện GPU NVIDIA hoặc lỗi NVML: {str(e)}")

# Hàm lấy và hiển thị thông tin CPU
def display_cpu_usage():
    cpu = psutil.cpu_percent(interval=1)
    print(f"CPU Usage: {cpu}%")

# Hàm lấy và hiển thị thông tin RAM
def display_ram_usage():
    ram = psutil.virtual_memory().percent
    print(f"RAM Usage: {ram}%")

# Hàm lấy và hiển thị thông tin GPU (nếu có NVIDIA)
def display_gpu_usage():
    if not gpu_available:
        print("GPU: Không có GPU NVIDIA")
        return
    try:
        device_count = py3nvml.nvmlDeviceGetCount()
        if device_count == 0:
            print("GPU: Không có GPU")
            return
        gpu_usage = 0
        for i in range(device_count):
            handle = py3nvml.nvmlDeviceGetHandleByIndex(i)
            usage = py3nvml.nvmlDeviceGetUtilizationRates(handle).gpu
            gpu_usage += usage
        avg_gpu_usage = gpu_usage / device_count if device_count > 0 else 0
        print(f"GPU Usage: {avg_gpu_usage}%")
    except Exception as e:
        print(f"GPU: Lỗi - {str(e)}")

# Hàm chính để hiển thị thông tin
def display_system_info():
    while True:
        print("\n" + "="*40)
        print(f"Thời gian: {time.strftime('%Y-%m-%d %H:%M:%S')}")
        display_cpu_usage()
        display_ram_usage()
        display_gpu_usage()
        print("="*40)
        time.sleep(5)  # Cập nhật mỗi 5 giây

if __name__ == "__main__":
    print("Bắt đầu hiển thị thông tin hệ thống...")
    display_system_info()