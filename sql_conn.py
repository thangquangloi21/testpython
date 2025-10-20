import pyodbc

class SQLQuery:
    def __init__(self, server, database, username, password, driver='{SQL Server}'):
        """Khởi tạo kết nối với cơ sở dữ liệu SQL Server"""
        self.server = server
        self.database = database
        self.username = username
        self.password = password
        self.driver = driver
        self.conn = None
        self.cursor = None
        self.connect()

    def connect(self):
        """Kết nối tới cơ sở dữ liệu"""
        try:
            conn_str = f'DRIVER={self.driver};SERVER={self.server};DATABASE={self.database};UID={self.username};PWD={self.password}'
            self.conn = pyodbc.connect(conn_str)
            self.cursor = self.conn.cursor()
            print("Kết nối thành công!")
        except pyodbc.Error as e:
            print(f"Lỗi kết nối: {e}")
            raise

    def execute_query(self, query, params=None):
        """Thực hiện truy vấn SQL, hỗ trợ cả truy vấn có tham số"""
        try:
            if params:
                self.cursor.execute(query, params)
            else:
                self.cursor.execute(query)
            # Kiểm tra xem truy vấn có trả về dữ liệu không
            if self.cursor.description:  # Nếu có dữ liệu trả về
                rows = self.cursor.fetchall()
                return rows
            self.conn.commit()  # Commit cho các truy vấn như INSERT, UPDATE, DELETE
            print("Truy vấn thực hiện thành công!")
            return None
        except pyodbc.Error as e:
            print(f"Lỗi khi thực hiện truy vấn: {e}")
            raise

    def close(self):
        """Đóng kết nối"""
        try:
            if self.cursor:
                self.cursor.close()
            if self.conn:
                self.conn.close()
                print("Đã đóng kết nối.")
        except pyodbc.Error as e:
            print(f"Lỗi khi đóng kết nối: {e}")
            raise

