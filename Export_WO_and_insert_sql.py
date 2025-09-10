import pandas as pd
from sqlalchemy import create_engine, text
import pyautogui
import time
import os
import pygetwindow as gw
import datetime
import logging
import sys
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import logging
import subprocess

class auto_export_qad:
    def __init__(self):
        self.setup_logging()
        self.logger.debug("Application initialized")
        # self.export_his()
        self.export_wo()
        self.copydata()


    def copydata(self):
        try:
            conect = self.conn()
            query = """
            drop table [Export_data_QAD].[dbo].[trans_export]
            SELECT * INTO [Export_data_QAD].[dbo].[trans_export]
            FROM [QAD_FG_Management].[dbo].[trans_export]
            WHERE 1 = 0;
            INSERT INTO [Export_data_QAD].[dbo].[trans_export]
            SELECT * FROM [QAD_FG_Management].[dbo].[trans_export];
            """
            with conect.begin() as connection:
                # Thực thi câu lệnh SQL
                result = connection.execute(text(query))
                # Commit giao dịch
                connection.commit()
                self.logger.debug(f"Đã cập nhật bản ghi.")
            print("Sao chép bảng hoàn tất!")
        except Exception as err:
            self.logger.error(f"Error: {err}")
            print(err)
    def setup_logging(self):
        """
        Configure logging with both file and console output
        """
        # Create a logger
        self.logger = logging.getLogger('AutoExportQAD')
        self.logger.setLevel(logging.DEBUG)

        # Create logs directory if it doesn't exist
        log_dir = 'logs'
        os.makedirs(log_dir, exist_ok=True)

        # Create a file handler with utf-8 encoding
        log_file_path = os.path.join(log_dir, f'auto_export_qad_{datetime.datetime.now().strftime("%Y%m%d")}.log')
        file_handler = logging.FileHandler(log_file_path, encoding='utf-8')
        file_handler.setLevel(logging.DEBUG)

        # Create console handler with utf-8 encoding
        console_handler = logging.StreamHandler(sys.stdout.buffer)
        console_handler.setLevel(logging.INFO)

        # Create formatter
        formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        file_handler.setFormatter(formatter)
        console_handler.setFormatter(formatter)

        # Add handlers to logger
        self.logger.addHandler(file_handler)
        self.logger.addHandler(console_handler)

    def conn(self):
        # Database connection and data processing
        try:
            server = '10.239.1.54'
            database = 'Export_data_QAD'
            username = 'sa'
            password = '123456'
            driver = 'ODBC Driver 18 for SQL Server'

            connection_string = f'mssql+pyodbc://{username}:{password}@{server}/{database}?driver={driver}&TrustServerCertificate=yes'
            engine = create_engine(connection_string)
            self.logger.debug("connect success")
            return engine
        except Exception as e:
            self.logger.error(f"Error connecting to database: {e}")

    def delete_file(self, file_path):
        if os.path.exists(file_path):
            os.remove(file_path)
            self.logger.debug(f"Tệp {file_path} đã được xóa.")
        else:
            self.logger.error("Tệp không tồn tại.")


    def check_table_data(self, table):
        try:
            conect = self.conn()
            # if data cout > 0 delete table
            with conect.begin() as connection:
                result = connection.execute(text(f"SELECT COUNT(*) FROM {table}"))
                row_count = result.scalar()
                if row_count > 0:
                    self.logger.debug("Có dữ liệu trong bảng.")
                    self.logger.debug(f"The table '{table}' has {row_count} rows.")
                    self.logger.debug(f"DELETE FROM {table}")
                    with conect.begin() as connection:
                        connection.execute(text(f"DELETE FROM {table}"))
                        return
                else:
                    self.logger.debug(f"The table '{table}' is empty.")
                    self.logger.debug("Không có dữ liệu trong bảng")
                    return
        except Exception as e:
            self.logger.error(f"error: {e}")

    # readdata and insert data into sql Sever
    def insert_trans_to_sql(self, table, file_path):
        try:
            if table == 'wo_browse':
                # Read CSV file
                self.logger.debug("Đang đọc dữ liệu từ tệp CSV...")
                df1 = pd.read_csv(file_path, delimiter=",", encoding='ISO-8859-1',dtype={'column_name': str})
                df1['order_date'] = pd.to_datetime(df1['order_date'], format='%d/%m/%y', errors='coerce')
                df1['due_date'] = pd.to_datetime(df1['due_date'], format='%d/%m/%y', errors='coerce')
                # check an drop data in table
                self.check_table_data(table)
                # Insert data into SQL Server
                conect = self.conn()          
                # pd.read_sql(text("DROP TABLE IF EXISTS wo_bill_browse"), conect)
                self.logger.debug("Đang Insert data vào SQL")
                df1.to_sql(table, con=conect, if_exists='append', index=False)
                self.logger.debug("Data processed and inserted into database successfully")

            if table == 'wo_bill_browse':
                # Read CSV file
                self.logger.debug("Đang đọc dữ liệu từ tệp CSV...")
                df1 = pd.read_csv(file_path, delimiter=",", encoding='ISO-8859-1',dtype={'column_name': str})
                df1['issue_date'] = pd.to_datetime(df1['issue_date'], format='%d/%m/%y', errors='coerce')
                # check an drop data in table
                self.check_table_data(table)
                # Insert data into SQL Server
                conect = self.conn()          
                self.logger.debug("Đang Insert data vào SQL")
                df1.to_sql(table, con=conect, if_exists='append', index=False)
                self.logger.debug("Data processed and inserted into database successfully")
            else:
                self.logger.debug(f"Không có bảng {table}")
            self.delete_file(file_path)
            email_body = f"<p>Xuất dữ liệu {table} thành công</p>"
            self.send_email("Export auto QAD:", email_body, "tvc_adm_it@terumo.co.jp")
        except Exception as e:
            self.logger.debug(f"error: {e}")
            email_body = f"<p>Lỗi trong quá trình xuất: {str(e)}</p>"
            # self.send_email("Export auto trans_history Error:", email_body, "tvc_adm_it@terumo.co.jp")

    def open_and_login(self, app_path, username, password, open_name):  # (Đường dẫn, TK, MK, Tên ứng dụng khi đã mở hoàn tất)
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
                    time.sleep(2)  # Đợi một chút trước khi kiểm tra lại

            # Thực hiện đăng nhập
            pyautogui.write(username)  # Nhập tên người dùng
            pyautogui.press('tab')  # Nhấn Tab để chuyển đến trường mật khẩu
            pyautogui.write(password)  # Nhập mật khẩu
            pyautogui.press('enter')  # Nhấn Enter để đăng nhập

            # Kiểm tra xem ứng dụng đã mở chưa
            if open_name in windows:
                self.logger.error("Khởi động APP Thành công")

        except FileNotFoundError:
            self.logger.error("Không tìm thấy ứng dụng (Sai Link)")
        except Exception as e:
            self.logger.error(f"Lỗi: {e}")


    # Tắt ứng dụng nếu đang mở
    def bring_app_to_front(self, window_title):
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

    def open_progess_editor(self):
        try:
            # QAD
            QADprod = r'C:\ProgramData\Microsoft\Windows\Start Menu\Programs\QAD Enterprise Applications\QAD 2018EE PROD'
            QADtest = r'C:\ProgramData\Microsoft\Windows\Start Menu\Programs\QAD Enterprise Applications\QAD 2018EE TEST'
            QADtest_OPEN = r'tvctest: TVC TVC [USD] > TVC TVC (1) - QAD Enterprise Applications'
            QADprod_OPEN = r'tvcprod: TVC TVC [USD] > TVC TVC (1) - QAD Enterprise Applications'
            username = 'mfg'
            passw = 'tvcadmin'

            # title app
            QADprod_title = r"tvcprod: TVC TVC [USD] > TVC TVC - QAD Enterprise Applications"
            QADtest_title = r"tvctest: TVC TVC [USD] > TVC TVC - QAD Enterprise Applications"
            openappqad = "Login"
            help = "Help"
            
            # kiểm tra app đã mở chưa
            self.bring_app_to_front(help)
            self.bring_app_to_front(openappqad)
            self.bring_app_to_front(QADprod_title)
            self.bring_app_to_front(QADprod_OPEN)
            self.open_and_login(app_path=QADprod, username=username, password=passw, open_name=QADprod_title)  # QAD
            time.sleep(5)
            # nhập progess editor
            pyautogui.write("Progress Editor")
            time.sleep(2)
            pyautogui.press('enter')
            time.sleep(5)
        
        except Exception as e:
            self.logger.error(f"Lỗi: {e}")
            email_body = f"<p>Lỗi trong quá trình xuất: {str(e)}</p>"
            self.send_email("Export auto QAD Error:", email_body, "tvc_adm_it@terumo.co.jp")

    def get_time_modified(self, file_path):
            modified_time = os.path.getmtime(file_path)
            modified_date = datetime.datetime.fromtimestamp(modified_time)
            rounded_date = modified_date.replace(minute=0, second=0, microsecond=0)
            formatted_date = rounded_date.strftime('%Y-%m-%d %H:%M:%S')
            self.logger.debug(f'File modified time (rounded to hour): {formatted_date}')
            return formatted_date

    def get_current_rounded_hour(self):
            current_time = datetime.datetime.now()
            rounded_time = current_time.replace(minute=0, second=0, microsecond=0)
            formatted_time = rounded_time.strftime('%Y-%m-%d %H:%M:%S')
            self.logger.debug(f"Current rounded hour: {formatted_time}")
            return formatted_time
    def send_email(self, subject, body, to_email):
        try:
            smtp_server = "10.98.28.206"
            smtp_port = 25
            from_email = "Export_auto_QAD@terumo.co.jp"

            msg = MIMEMultipart()
            msg['From'] = from_email
            msg['To'] = to_email
            msg['Subject'] = subject

            msg.attach(MIMEText(body, 'html'))

            server = smtplib.SMTP(smtp_server, smtp_port)
            server.send_message(msg)
            self.logger.debug(f"Email sent successfully to {to_email}")

            server.quit()
        except Exception as e:
            self.logger.error(f"Failed to send email: {e}", exc_info=True)

    def export_wo(self):
        try:
            from datetime import timedelta,datetime
            self.open_progess_editor()
            today = datetime.now()
            # Calculate yesterday's date
            yesterday = today - timedelta(days=1)
            yesterday = yesterday.strftime("%d/%m/%y")
            # exit unikey
            try:
                self.logger.debug("Check unikey is open?")
                subprocess.call(["taskkill", "/F", "/IM", "UniKeyNT.exe"])
            except Exception as e:
                 self.logger.debug("Unikey is not opening")
            # yesterday = "26/03/25"
            # Print yesterday's date in the desired format
            # nhập code
            code = f'''
            OUTPUT TO "/home/mfg/WO_browse.csv".
            PUT UNFORMATTED "id,work_order,item_number,wo_status,lot,order_qty,qty_comp,order_date,due_date" SKIP.
            FOR EACH wo_mstr NO-LOCK
            WHERE wo_mstr.wo_due_date >= 01/01/25:
            PUT UNFORMATTED
            (IF wo_mstr.wo_lot = ? THEN "" ELSE wo_mstr.wo_lot) ","
            (IF wo_mstr.wo_nbr = ? THEN "" ELSE wo_mstr.wo_nbr) ","
            (IF wo_mstr.wo_part = ? THEN "" ELSE wo_mstr.wo_part) ","
            (IF wo_mstr.wo_status = ? THEN "" ELSE wo_mstr.wo_status) ","
            (IF wo_mstr.wo_lot_next = ? THEN "" ELSE wo_mstr.wo_lot_next) ","
            (IF wo_mstr.wo_qty_ord = ? THEN 0 ELSE wo_mstr.wo_qty_ord) ","
            (IF wo_mstr.wo_qty_comp = ? THEN 0 ELSE wo_mstr.wo_qty_comp) ","
            (IF wo_mstr.wo_ord_date = ? THEN "" ELSE STRING(wo_mstr.wo_ord_date, "99/99/99")) ","
            (IF wo_mstr.wo_due_date = ? THEN "" ELSE STRING(wo_mstr.wo_due_date, "99/99/99"))
            SKIP.
            END.
            OUTPUT CLOSE.

            OUTPUT TO "/home/mfg/WO_bill_browse.csv".
            PUT UNFORMATTED "id,work_order,part_number,qty_req,qty_to_issue,qty_issued,issue_date" SKIP.
            FOR EACH wod_det NO-LOCK
            WHERE wod_det.wod_iss_date >= 01/01/25:
            PUT UNFORMATTED
            (IF wod_det.wod_lot = ? THEN "" ELSE wod_det.wod_lot) ","
            (IF wod_det.wod_nbr = ? THEN "" ELSE wod_det.wod_nbr) ","
            (IF wod_det.wod_part = ? THEN "" ELSE wod_det.wod_part) ","
            (IF wod_det.wod_qty_req = ? THEN 0 ELSE wod_det.wod_qty_req) ","
            (IF wod_det.wod_qty_chg = ? THEN 0 ELSE wod_det.wod_qty_chg) ","
            (IF wod_det.wod_qty_iss = ? THEN 0 ELSE wod_det.wod_qty_iss) ","
            (IF wod_det.wod_iss_date = ? THEN "" ELSE STRING(wod_det.wod_iss_date, "99/99/99"))
            SKIP.
            END.
            OUTPUT CLOSE.
            '''
            file_path1=r"Z:\WO_browse.csv"
            file_path2=r"Z:\WO_bill_browse.csv"
            pyautogui.write(code)
            time.sleep(5)
            pyautogui.press('f1')
            time.sleep(30)
            pyautogui.press('enter')
            time.sleep(2)
            pyautogui.hotkey('alt', 'f4')
            get_time_Date = self.get_current_rounded_hour()
            get_time_update_file1 = self.get_time_modified(file_path1)
            get_time_update_file2 = self.get_time_modified(file_path2)
            if get_time_update_file1 and get_time_update_file2 == get_time_Date:
                self.logger.debug("Files exported successfully")
                self.insert_trans_to_sql(file_path=file_path1,table='wo_browse')
                self.insert_trans_to_sql(file_path=file_path2,table='wo_bill_browse')
            else:
                self.logger.error("Files exported failed")
        except Exception as e:
            self.logger.error(f"Error: {e}")
            email_body = f"<p>Lỗi trong quá trình xuất: {str(e)}</p>"
            self.send_email(f"Export auto WO Error:", email_body, "tvc_adm_it@terumo.co.jp")

if __name__ == "__main__":
   app = auto_export_qad()