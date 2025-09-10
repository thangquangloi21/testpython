import os
import pyautogui
import time
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import smtplib
from datetime import datetime

def open_and_capture():
    path = r"C:\Users\Public\Desktop\mRemoteNG.lnk"
    os.startfile(path)

    time.sleep(5)
    pyautogui.click(100, 200)
    # Press key down
    pyautogui.press("down")
    time.sleep(1)

    pyautogui.press("enter")
    time.sleep(5)
    region = (210, 600, 1050, 260)
    
    # GET TIME AND DATE
    # Get the current date and time
    now = datetime.now()

# Format the date and time as a string
    formatted_now = now.strftime("%Y%m%d%H%M%S")
    # print(formatted_now)
    # Take a screenshot of the specified region
    screenshot_path = f'region_screenshot.png'
    pyautogui.screenshot(screenshot_path, region=region)

    email_body = f"""
    <p>Ảnh Backup sever QAD:</p>
    <img src="cid:image1">
    """
    with open('list_mail.txt', 'r', encoding='utf-8') as file:
        # Read each line in the file
        for mail in file:
            # Print the line
            print(mail.strip())
            send_email("Hình ảnh sever QAD_Backup:", email_body, mail, screenshot_path)

    # Tắt mRemoteNG
    pyautogui.move(1920,1)
    pyautogui.click(1920,1)
def send_email(subject, body, to_email, image_path):
    try:
        smtp_server = "10.98.28.206"
        smtp_port = 25
        from_email = "Sever_QAD_backup@terumo.co.jp"

        msg = MIMEMultipart()
        msg['From'] = from_email
        msg['To'] = to_email
        msg['Subject'] = subject

        msg.attach(MIMEText(body, 'html'))

        # Attach the image
        with open(image_path, 'rb') as img_file:
            img = MIMEImage(img_file.read())
            img.add_header('Content-ID', '<image1>')
            msg.attach(img)

        server = smtplib.SMTP(smtp_server, smtp_port)
        server.send_message(msg)
        print(f"Email sent successfully to {to_email}")
    except Exception as e:
        print(f"Failed to send email: {e}")

# Call the function to execute the script
open_and_capture()