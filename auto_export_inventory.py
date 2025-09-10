import tkinter as tk
import datetime
import time
import threading
import Auto_Log as Auto_Log
import os
import pandas as pd # pip install pandas
from sqlalchemy import create_engine # pip install sqlalchemy
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import logging
import sys

class auto_export_qad:
    def __init__(self):
        # Set up logging
        self.setup_logging()
        self.logger.info("Application initialized")
        self.run_function()

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

        # Create a file handler
        log_file_path = os.path.join(log_dir, f'auto_export_qad_{datetime.datetime.now().strftime("%Y%m%d")}.log')
        file_handler = logging.FileHandler(log_file_path)
        file_handler.setLevel(logging.DEBUG)

        # Create console handler
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setLevel(logging.INFO)

        # Create formatter
        formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        file_handler.setFormatter(formatter)
        console_handler.setFormatter(formatter)

        # Add handlers to logger
        self.logger.addHandler(file_handler)
        self.logger.addHandler(console_handler)

    def run_function(self):
        try:
            self.logger.info("Starting run_function")
            
            # Existing run_function code with logging added
            Auto_Log.getallwindow()
            self.logger.debug("Retrieved all windows")
            
            Auto_Log.bring_app_to_front(Auto_Log.help)
            Auto_Log.bring_app_to_front(Auto_Log.openappqad)
            Auto_Log.bring_app_to_front(Auto_Log.QADtest_title)
            Auto_Log.bring_app_to_front(Auto_Log.QADprod_OPEN)
            self.logger.debug("Brought all required applications to front")

            # Open and login
            Auto_Log.open_and_login(
                app_path=Auto_Log.QADprod, 
                username=Auto_Log.username, 
                password=Auto_Log.passw, 
                open_name=Auto_Log.QADprod_title
            )
            self.logger.info("Logged into QAD application")
            
            time.sleep(3)
            Auto_Log.pyautogui.write("Progress Editor")
            self.logger.debug("Write 'Progress Editor'")

            time.sleep(2)
            Auto_Log.pyautogui.press("enter")
            self.logger.debug("Pressed enter")

            time.sleep(2)
            Auto_Log.paste_data(Auto_Log.text_to_paste)
            self.logger.info("Pasted data into Progress Editor")

            time.sleep(2)
            Auto_Log.pyautogui.press('f1')
            self.logger.debug("Pressed F1")

            time.sleep(15)
            self.logger.info("Export process completed")

            time.sleep(2)
            Auto_Log.pyautogui.press('enter')
            time.sleep(2)
            Auto_Log.pyautogui.hotkey('alt', 'f4')
            self.logger.debug("Closed application windows")

            get_time_Date = self.get_current_rounded_hour()
            file = r"Z:\ld_det_export.csv"
            file_path = r"Z:\pt_mstr_export.csv"
            get_time_update_file1 = self.get_time_modified(file)
            get_time_update_file2 = self.get_time_modified(file_path)

            if get_time_update_file1 and get_time_update_file2 == get_time_Date:
                self.logger.info("Files exported successfully")
                
                # Database connection and data processing
                server = '10.239.1.54'
                database = 'QAD_FG_Management'
                username = 'sa'
                password = '123456'
                driver = 'ODBC Driver 18 for SQL Server'

                connection_string = f'mssql+pyodbc://{username}:{password}@{server}/{database}?driver={driver}&TrustServerCertificate=yes'
                engine = create_engine(connection_string)
                self.logger.debug("Database connection established")

                # Read CSV files
                df1 = pd.read_csv(r"Z:\pt_mstr_export.csv", delimiter="'", encoding='ISO-8859-1')
                df2 = pd.read_csv(r"Z:\ld_det_export.csv", delimiter="'", encoding='ISO-8859-1')
                self.logger.debug("CSV files read successfully")

                # Merge and filter data
                merged_df = pd.merge(df1, df2, on='Part')
                filtered_df = merged_df[(merged_df['PartType'] == 'FG') | (merged_df['PartType'].str.startswith('HFG')) ]

                current_datetime = datetime.datetime.now()
                filtered_df.loc[:, 'Merge_Date'] = current_datetime.date()
                filtered_df.loc[:, 'Merge_Time'] = current_datetime.time().strftime('%H:%M:%S')
                filtered_df.save("123.csv")
                # Uncomment to insert data
                
                # filtered_df.to_sql('merged_output_with_datetime', con=engine, if_exists='append', index=False)
                self.logger.info("Data processed and ready for database insertion")

                # Send success email
                email_body = "<p>Xuất Tồn kho thành công !</p>"
                self.send_email("Export auto Inventory:", email_body, "tvc_adm_it@terumo.co.jp")
            else:
                self.logger.warning("Export failed")
                email_body = "<p>Xuất Tồn kho thất bại !</p>"
                self.send_email("Export auto Inventory:", email_body, "tvc_adm_it@terumo.co.jp")

        except Exception as e:
            self.logger.error(f"An error occurred: {e}", exc_info=True)
            # Send error email
            email_body = f"<p>Lỗi trong quá trình xuất: {str(e)}</p>"
            self.send_email("Export auto Inventory Error:", email_body, "tvc_adm_it@terumo.co.jp")

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
            self.logger.info(f"Email sent successfully to {to_email}")

            server.quit()
        except Exception as e:
            self.logger.error(f"Failed to send email: {e}", exc_info=True)

    def get_current_rounded_hour(self):
        current_time = datetime.datetime.now()
        rounded_time = current_time.replace(minute=0, second=0, microsecond=0)
        formatted_time = rounded_time.strftime('%Y-%m-%d %H:%M:%S')
        self.logger.debug(f"Current rounded hour: {formatted_time}")
        return formatted_time

    def get_time_modified(self, file_path):
        modified_time = os.path.getmtime(file_path)
        modified_date = datetime.datetime.fromtimestamp(modified_time)
        rounded_date = modified_date.replace(minute=0, second=0, microsecond=0)
        formatted_date = rounded_date.strftime('%Y-%m-%d %H:%M:%S')
        self.logger.debug(f'File modified time (rounded to hour): {formatted_date}')
        return formatted_date

    def schedule_function(self, hour, minute):
        now = datetime.datetime.now()
        target_time = now.replace(hour=hour, minute=minute, second=0, microsecond=0)
        if target_time < now:
            target_time += datetime.timedelta(days=1)
        delay = (target_time - now).total_seconds()
        threading.Timer(delay, self.run_function).start()
        self.logger.info(f"Function scheduled to run at {target_time}")

if __name__ == "__main__":
    app = auto_export_qad()