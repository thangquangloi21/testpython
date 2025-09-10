import pyautogui
import time
import os
import pygetwindow as gw
import pyperclip


# tự động mở app và login
Vega32 = "D:\LOIII\MES VS QAD\MES\MES APP\MES THAT (Production Environment)\Vega32.rdp"
Operator = 'D:\LOIII\MES VS QAD\MES\MES APP\MES THAT (Production Environment)\Operator.rdp'
# QAD
QADprod = 'C:\ProgramData\Microsoft\Windows\Start Menu\Programs\QAD Enterprise Applications\QAD 2018EE PROD'
QADtest = 'C:\ProgramData\Microsoft\Windows\Start Menu\Programs\QAD Enterprise Applications\QAD 2018EE TEST'
QADtest_OPEN = 'tvctest: TVC TVC [USD] > TVC TVC (1) - QAD Enterprise Applications'
QADprod_OPEN = 'tvcprod: TVC TVC [USD] > TVC TVC (1) - QAD Enterprise Applications'
username = 'mfg'
passw = 'tvcadmin'
text_to_paste = r'Code.txt'

# title app
VEGA_title = "CIMVisionVEGA"
OPERATOR_title = "CIMVisionPharms (Remote)"
QADprod_title = "tvcprod: TVC TVC [USD] > TVC TVC - QAD Enterprise Applications"
QADtest_title = "tvctest: TVC TVC [USD] > TVC TVC - QAD Enterprise Applications"
btn_progess = r"1image.png"
img = r"image.png"

openappqad = "Login"
help = "Help"

def getallwindow():
    windows = gw.getAllTitles()
    print("Danh sách các cửa sổ đang mở:")
    for title in windows:
        if title == "":
            continue
        else:
            print(f"- {title}")


def open_and_login(app_path, username, password, open_name):  # (Đường dẫn, TK, MK, Tên ứng dụng khi đã mở hoàn tất)
    openappqad = "Login"
    try:
        # Mở ứng dụng
        os.startfile(app_path)
        # time.sleep(0.5)  # Đợi một chút để ứng dụng mở

        # Kiểm tra xem cửa sổ "Login" có mở không
        while True:
            windows = gw.getAllTitles()
            if openappqad in windows:
                print("Cửa sổ Login đã mở.")
                break
            # Nếu gặp update QAD
            elif "Update QAD (3.4.0.41)?" in windows:
                window = gw.getWindowsWithTitle("Update QAD (3.4.0.41)?")
                if window:
                    window[0].close()  # Đóng cửa sổ
                    print(f"Cửa sổ '{"Update QAD (3.4.0.41)?"}' đã được đóng.")
                else:
                    print(f"Cửa sổ '{"Update QAD (3.4.0.41)?"}' không tìm thấy.")
                print("Thoát Update")
            else:
                print("Cửa sổ Login chưa mở, đang chờ...")
                time.sleep(1)  # Đợi một chút trước khi kiểm tra lại

        # Thực hiện đăng nhập
        pyautogui.write(username)  # Nhập tên người dùng
        pyautogui.press('tab')  # Nhấn Tab để chuyển đến trường mật khẩu
        pyautogui.write(password)  # Nhập mật khẩu
        pyautogui.press('enter')  # Nhấn Enter để đăng nhập

        # Kiểm tra xem ứng dụng đã mở chưa
        if open_name in windows:
            print("Khởi động APP Thành công")

    except FileNotFoundError:
        print("Không tìm thấy ứng dụng (Sai Link)")
    except Exception as e:
        print(f"Lỗi: {e}")


def check_export_sc(img):
    try:
        # Tìm vị trí của hình ảnh trên màn hình
        button_location = pyautogui.locateOnScreen(img, confidence=0.8)
        if button_location:
            # Lấy tọa độ trung tâm của nút
            button_center = pyautogui.center(button_location)

            # Di chuyển chuột đến tọa độ trung tâm của nút
            pyautogui.moveTo(button_center)

            # Nhấn vào nút
            pyautogui.click(button_center)
            print("Xuất thành công")
            return True
        else:
            print("Xuất thất bại: Không tìm thấy nút")
            return False
    except pyautogui.ImageNotFoundException:
        print("Đang xử lý")
        return False
    except Exception as e:
        print(f"Xuất thất bại: {e}")
        return False


def click_to_img(img):
    # Tìm vị trí của nút trên màn hình
    button_location = pyautogui.locateOnScreen(img, confidence=0.8)
    print(button_location)
    if button_location:
        # Nhấn vào nút
        pyautogui.moveTo(button_location)
        pyautogui.press("enter")
        print("Đã nhấn vào nút!")
    else:
        print("Không tìm thấy nút.")


# doc du lieu tu file .txt roi paste
def paste_data(file_name):
    with open(file_name, 'r', encoding='utf-8') as file:
        text_to_paste = file.read()
    # Dán văn bản vào Notepad
    # Dán văn bản vào Notepad
    pyautogui.write(text_to_paste)


# Đẩy lên đầu tiên
def bring_app_to_front(window_title):
    try:

        # Tìm cửa sổ theo tiêu đề
        window = gw.getWindowsWithTitle(window_title)[0]
        # Khôi phục cửa sổ nếu nó đang ở trạng thái minimize
        if window.isMinimized:
            window.restore()
            # Đưa cửa sổ lên trên cùng
        window.activate()
        print(f"Cửa sổ '{window_title}' đã được đưa lên trên cùng.")
        pyautogui.hotkey('alt', 'f4')
    except IndexError:
        print(f"Cửa sổ '{window_title}' không tìm thấy.")


if __name__ == "__main__":
    # mes
    Vega32 = "D:\LOIII\MES VS QAD\MES\MES APP\MES THAT (Production Environment)\Vega32.rdp"
    Operator = 'D:\LOIII\MES VS QAD\MES\MES APP\MES THAT (Production Environment)\Operator.rdp'
    # QAD
    QADprod = 'C:\ProgramData\Microsoft\Windows\Start Menu\Programs\QAD Enterprise Applications\QAD 2018EE PROD'
    QADtest = 'C:\ProgramData\Microsoft\Windows\Start Menu\Programs\QAD Enterprise Applications\QAD 2018EE TEST'
    QADtest_OPEN = 'tvctest: TVC TVC [USD] > TVC TVC (1) - QAD Enterprise Applications'
    QADprod_OPEN = 'tvcprod: TVC TVC [USD] > TVC TVC (1) - QAD Enterprise Applications'
    username = 'mfg'
    passw = 'tvcadmin'
    text_to_paste = r'Code.txt'

    # title app
    VEGA_title = "CIMVisionVEGA"
    OPERATOR_title = "CIMVisionPharms (Remote)"
    QADprod_title = "tvcprod: TVC TVC [USD] > TVC TVC - QAD Enterprise Applications"
    QADtest_title = "tvctest: TVC TVC [USD] > TVC TVC - QAD Enterprise Applications"
    btn_progess = r"1image.png"
    img = r"image.png"

    openappqad = "Login"
    help = "Help"

    # kiem tra app da mo chua
    getallwindow()
    bring_app_to_front(help)
    bring_app_to_front(openappqad)
    bring_app_to_front(QADtest_title)
    bring_app_to_front(QADtest_OPEN)
    # Mở app rồi đăng nhập
    open_and_login(app_path=QADtest, username=username, password=passw, open_name=QADprod_title)  # QAD
    time.sleep(2)
    # tìm progess editor

    pyautogui.write("Progress Editor")

    time.sleep(2)
    # ấn enter
    pyautogui.press("enter")
    # click_to_img(btn_progess)
    time.sleep(2)
    # paste data
    paste_data(text_to_paste)
    print("Nhập xong chờ ấn f1")
    # nhập xong ấn f1
    time.sleep(2)
    pyautogui.press('f1')
    # đợi xuất ra
    time.sleep(60)
    print("Xuất thành công đợi 3s nữa đóng cửa sổ")
    time.sleep(2)
    pyautogui.press('enter')
    time.sleep(2)
    pyautogui.hotkey('alt', 'f4')

    # open_and_login(app_path=Operator, username=username,password=passw,open_name=OPERATOR_title) #Operator
    # open_and_login(app_path=Vega32, username=username,password=passw,open_name=VEGA_title) #vega32
    # getallwindow()