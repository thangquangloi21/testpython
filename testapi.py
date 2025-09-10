from flask import Flask, request, jsonify

app = Flask(__name__)

@app.route('/nhan-json', methods=['POST'])
def nhan_json():
    data = request.get_json()
    print("Dữ liệu nhận được:", data)
    return jsonify({"message": "Dữ liệu đã được in ra màn hình", "data": data}), 200

if __name__ == '__main__':
    app.run(debug=True)
