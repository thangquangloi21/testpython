# working_thread.py
import queue
import threading
import time
import random

class DataProcessor(threading.Thread):
    def __init__(self, name, data_queue):
        """
        Khởi tạo luồng xử lý dữ liệu
        
        :param name: Tên của luồng
        :param data_queue: Hàng đợi chứa dữ liệu để xử lý
        """
        threading.Thread.__init__(self)
        self.name = name
        self.data_queue = data_queue
        self.daemon = True  # Cho phép luồng kết thúc khi chương trình chính kết thúc
    
    def process_data(self, data):
        """
        Xử lý dữ liệu 
        
        :param data: Dữ liệu cần xử lý
        """
        print(f"{self.name} đang xử lý: {data}")
        # Mô phỏng thời gian xử lý
        time.sleep(random.uniform(0.5, 2))
        print(f"{self.name} đã hoàn thành xử lý: {data}")
    
    def run(self):
        """
        Phương thức chính được gọi khi start() được gọi
        Liên tục lấy dữ liệu từ hàng đợi và xử lý
        """
        while True:
            try:
                # Lấy dữ liệu từ hàng đợi, chờ tối đa 1 giây
                data = self.data_queue.get(timeout=1)
                
                # Nếu nhận được None, đó là tín hiệu để dừng luồng
                if data is None:
                    break
                
                # Xử lý dữ liệu
                self.process_data(data)
                
                # Đánh dấu nhiệm vụ đã hoàn thành
                self.data_queue.task_done()
            
            except queue.Empty:
                # Nếu hàng đợi trống, tiếp tục vòng lặp
                continue