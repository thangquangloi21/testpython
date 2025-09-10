import http.server
import socketserver

# Địa chỉ IP của card Wi-Fi (thay bằng IP thật của bạn)
LISTEN_IP = "172.31.98.248"
LISTEN_PORT = 8080

class SimpleProxy(http.server.SimpleHTTPRequestHandler):
    def do_GET(self):
        print(f"[+] Yêu cầu GET đến: {self.path}")
        super().do_GET()

if __name__ == "__main__":
    with socketserver.TCPServer((LISTEN_IP, LISTEN_PORT), SimpleProxy) as httpd:
        print(f"[+] Proxy HTTP đang chạy tại {LISTEN_IP}:{LISTEN_PORT}")
        httpd.serve_forever()
