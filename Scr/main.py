# main.py
from working_thread import DataProcessor
import queue
import time

class MainApplication:
    def __init__(self, num_threads=50000):
        self.data_queue = queue.Queue()
        self.threads = []
        
        # Tạo các luồng xử lý dữ liệu
        for i in range(num_threads):
            thread = DataProcessor(f"Thread-{i+1}", self.data_queue)
            thread.start()
            self.threads.append(thread)
    
    def add_data(self, data):
        """Thêm dữ liệu vào hàng đợi để các luồng xử lý"""
        self.data_queue.put(data)
    
    def stop_threads(self):
        """Dừng tất cả các luồng"""
        for _ in self.threads:
            self.data_queue.put(None)  # Tín hiệu để các luồng dừng
        
        for thread in self.threads:
            thread.join()
    
    def run(self):
        """Phương thức chạy ứng dụng chính"""
        try:
            # Ví dụ về việc thêm dữ liệu
            for i in range(1000000):
                self.add_data(f"Dữ liệu {i+1}")
                # time.sleep(0.1)  # Mô phỏng việc thêm dữ liệu theo thời gian
        except KeyboardInterrupt:
            print("\nDừng ứng dụng...")
        finally:
            self.stop_threads()

if __name__ == "__main__":
    app = MainApplication()
    app.run()