import requests

# Địa chỉ của server Flask
url = "http://172.31.98.63:5000/generate"

print("Bắt đầu chat live với API. Nhập 'exit' để thoát.")

while True:
    # Nhập prompt từ người dùng
    prompt = input("Bạn: ")
    
    # Thoát nếu nhập 'exit'
    if prompt.lower() == 'exit':
        print("Kết thúc cuộc trò chuyện.")
        break
    
    # Gửi yêu cầu GET đến API
    try:
        response = requests.get(url, params={"prompt:": prompt}, timeout=30)
        
        # Kiểm tra và in kết quả
        if response.status_code == 200:
            data = response.json()
            print("BOT:", data["response"])
        else:
            print("Lỗi:", response.status_code, response.text)
    except Exception as e:
        print("Lỗi kết nối:", str(e))
