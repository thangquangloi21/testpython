import subprocess
import pyautogui
import time
import sys

# Danh sách tên máy in
print_name = {"MES-PRT-01", "MES-PRT-02", "MES-PRT-03", "MES-PRT-04", "MES-PRT-05", "MES-PRT-06", "MES-PRT-16"}

# Thiết lập pyautogui
pyautogui.PAUSE = 0.5
pyautogui.FAILSAFE = True

print("Chương trình đang chạy. Nhấn Enter và nhập 'q' để dừng sau mỗi máy in.")

# Lặp qua từng máy in
for printer in print_name:
    print(f"Đang xử lý máy in: {printer}")
    command = f'rundll32 printui.dll,PrintUIEntry /e /n "{printer}"'
    try:
        process = subprocess.Popen(command, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        time.sleep(3)

        if process.poll() is not None:
            error_output = process.stderr.read()
            raise Exception(f"Lỗi khi mở cửa sổ thuộc tính máy in: {error_output}")

        pyautogui.press('tab', presses=2, interval=0.2)
        pyautogui.write("3.15")
        pyautogui.press('tab')
        pyautogui.write("1.969")
        pyautogui.press('enter')
        print(f"Đã nhập thông số 3.15, 1.969 cho {printer} và nhấn Enter.")
        
        time.sleep(0.2)

        # Kiểm tra dừng bằng cách nhập 'q'
        user_input = input("Nhập 'q' để dừng hoặc nhấn Enter để tiếp tục: ").strip().lower()
        if user_input == 'q':
            print("Dừng chương trình theo yêu cầu người dùng.")
            sys.exit(0)

    except Exception as e:
        print(f"Lỗi khi xử lý {printer}: {e}")
        pyautogui.hotkey('alt', 'f4')
        
    time.sleep(1)

print("Đã hoàn tất xử lý tất cả máy in.")