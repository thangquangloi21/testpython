import pandas as pd
from sqlalchemy import create_engine
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

class auto_export_qad:
    def __init__(self):
        self.setup_logging()
        self.logger.debug("Application initialized")
        self.export_his()
    # connect db
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
            database = 'QAD_FG_Management'
            username = 'sa'
            password = '123456'
            driver = 'ODBC Driver 18 for SQL Server'

            connection_string = f'mssql+pyodbc://{username}:{password}@{server}/{database}?driver={driver}&TrustServerCertificate=yes'
            engine = create_engine(connection_string)
            return engine
        except Exception as e:
            self.logger.error(f"Error connecting to database: {e}")

    def delete_file(self, file_path):
        if os.path.exists(file_path):
            os.remove(file_path)
            self.logger.debug(f"Tệp {file_path} đã được xóa.")
        else:
            self.logger.error("Tệp không tồn tại.")

    # readdata and insert data into sql Sever
    def insert_trans_to_sql(self):
        try:
            # Read CSV file
            self.logger.debug("Đang đọc dữ liệu từ tệp CSV...")
            file_path = r"Z:\trans_export.csv"
            df1 = pd.read_csv(file_path, delimiter=",", encoding='ISO-8859-1')
            # Convert date columns to datetime format
            df1['date'] = pd.to_datetime(df1['date'], format='%d/%m/%y', errors='coerce')
            df1['effdate'] = pd.to_datetime(df1['effdate'], format='%d/%m/%y', errors='coerce')
            df1['time'] = pd.to_datetime(df1['time'], format='%H:%M:%S', errors='coerce').dt.time
            self.logger.debug("Đang Insert data vào QAD")
            # Insert data into SQL Server
            # conect = self.conn()
            # df1.to_sql('trans_export1', con=conect, if_exists='append', index=False)
            # self.logger.debug("Data processed and inserted into database successfully")
            
            # Establish connection
            conect = self.conn()
            
            try:
                # Read existing data from the table to check for duplicates
                existing_data = pd.read_sql(f"SELECT DISTINCT effdate FROM trans_export", conect)
                
                # Convert existing dates to datetime for comparison
                existing_dates = pd.to_datetime(existing_data['effdate'])
                
                # Filter out rows with tr_effdate that already exist in the table
                df_to_insert = df1[~df1['effdate'].isin(existing_dates)]
                
                # Insert only new data
                if not df_to_insert.empty:
                    self.logger.debug(f"Inserting {len(df_to_insert)} new rows")
                    df_to_insert.to_sql('trans_export', con=conect, if_exists='append', index=False)
                else:
                    self.logger.debug("No new data to insert")
            
            except Exception as e:
                self.logger.error(f"Error during data insertion: {e}")
                raise
            finally:
                # Close the connection
                conect.dispose()
            self.delete_file(file_path)
            email_body = f"<p>Xuất dữ liệu thành công</p>"
            self.send_email("Export auto trans_history:", email_body, "tvc_adm_it@terumo.co.jp")
        except Exception as e:
            self.logger.error(f"error: {e}")
            email_body = f"<p>Lỗi trong quá trình xuất: {str(e)}</p>"
            self.send_email("Export auto trans_history Error:", email_body, "tvc_adm_it@terumo.co.jp")

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
            QADprod = 'C:\ProgramData\Microsoft\Windows\Start Menu\Programs\QAD Enterprise Applications\QAD 2018EE PROD'
            QADtest = 'C:\ProgramData\Microsoft\Windows\Start Menu\Programs\QAD Enterprise Applications\QAD 2018EE TEST'
            QADtest_OPEN = 'tvctest: TVC TVC [USD] > TVC TVC (1) - QAD Enterprise Applications'
            QADprod_OPEN = 'tvcprod: TVC TVC [USD] > TVC TVC (1) - QAD Enterprise Applications'
            username = 'mfg'
            passw = 'tvcadmin'

            # title app
            QADprod_title = "tvcprod: TVC TVC [USD] > TVC TVC - QAD Enterprise Applications"
            QADtest_title = "tvctest: TVC TVC [USD] > TVC TVC - QAD Enterprise Applications"
            openappqad = "Login"
            help = "Help"
            
            # kiểm tra app đã mở chưa
            self.bring_app_to_front(help)
            self.bring_app_to_front(openappqad)
            self.bring_app_to_front(QADprod_title)
            self.bring_app_to_front(QADprod_OPEN)
            self.open_and_login(app_path=QADprod, username=username, password=passw, open_name=QADprod_title)  # QAD
            time.sleep(60)
            # nhập progess editor
            pyautogui.write("Progress Editor")
            time.sleep(2)
            pyautogui.press('enter')
            time.sleep(10)
        
        except Exception as e:
            self.logger.error(f"Lỗi: {e}")
            email_body = f"<p>Lỗi trong quá trình xuất: {str(e)}</p>"
            self.send_email("Export auto trans_history Error:", email_body, "tvc_adm_it@terumo.co.jp")

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
            from_email = "Export_Inventory_QAD@terumo.co.jp"

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

    def export_his(self):
        try:
            from datetime import timedelta,datetime
            self.open_progess_editor()
            today = datetime.now()
            # Calculate yesterday's date
            yesterday = today - timedelta(days=1)
            yesterday = yesterday.strftime("%d/%m/%y")
            # yesterday = "26/03/25"
            # Print yesterday's date in the desired format
            # nhập code
            code = f'''
            OUTPUT TO "/home/mfg/trans_export.csv".
            PUT UNFORMATTED "item_number,Lot,location,tran_type,order,date,time,effdate,change_qty,inven_status" SKIP.
            FOR EACH tr_hist NO-LOCK WHERE tr_hist.tr_effdate = DATE("{yesterday}"):
            PUT UNFORMATTED 
            STRING(tr_hist.tr_part) + "," + 
            STRING(tr_hist.tr_serial) + "," + 
            STRING(tr_hist.tr_loc) + "," + 
            STRING(tr_hist.tr_type) + "," + 
            STRING(tr_hist.tr_nbr) + "," + 
            STRING(tr_hist.tr_date) + "," + 
            STRING(tr_hist.tr_time, "HH:MM:SS") + "," + 
            STRING(tr_hist.tr_effdate, "99/99/99") + "," + 
            STRING(tr_hist.tr_qty_chg) + "," + 
            STRING(tr_hist.tr_status) SKIP.
            END.
            OUTPUT CLOSE.'''
            pyautogui.write(code)
            time.sleep(5)
            pyautogui.press('f1')
            time.sleep(300)
            pyautogui.press('enter')
            time.sleep(2)
            pyautogui.hotkey('alt', 'f4')
            get_time_Date = self.get_current_rounded_hour()
            get_time_update_file = self.get_time_modified(file_path=r"Z:\trans_export.csv")
            if get_time_update_file == get_time_Date:
                self.logger.debug("Files exported successfully")
                self.insert_trans_to_sql()
            else:
                self.logger.error("Files exported failed")
        except Exception as e:
            self.logger.error(f"Error: {e}")
            email_body = f"<p>Lỗi trong quá trình xuất: {str(e)}</p>"
            self.send_email("Export auto trans_his Error:", email_body, "tvc_adm_it@terumo.co.jp")

if __name__ == "__main__":
   app = auto_export_qad()