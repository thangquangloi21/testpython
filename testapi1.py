from flask import Flask, send_file
import os

app = Flask(__name__)

# Đường dẫn đến file bạn muốn tải
FILE_PATH = 'uploads/drives_all.zip'  # Thay 'your_file.txt' bằng tên file của bạn

@app.route('/')
def download_file():
    if os.path.exists(FILE_PATH):
        return send_file(FILE_PATH, as_attachment=True)
    return 'File not found', 404

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=555)