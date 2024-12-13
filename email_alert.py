import smtplib
import imaplib
import email
from email.mime.text import MIMEText
from excel_handler import ExcelHandler
from db import Database

class EmailAlert:
    def __init__(self):
        self.excel_handler = ExcelHandler()
        self.db = Database()

    def send_warning(self, student_email, subject, message):
        msg = MIMEText(message)
        msg['Subject'] = subject
        msg['From'] = "your-email@gmail.com"
        msg['To'] = student_email

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login("your-email@gmail.com", "your-password")
            smtp.send_message(msg)

    def process_incoming_emails(self):
        mail = imaplib.IMAP4_SSL("imap.gmail.com")
        mail.login("your-email@gmail.com", "your-password")
        mail.select("inbox")
        
        result, data = mail.search(None, "UNSEEN")
        email_ids = data[0].split()

        for email_id in email_ids:
            result, msg_data = mail.fetch(email_id, "(RFC822)")
            raw_email = msg_data[0][1].decode("utf-8")
            email_message = email.message_from_string(raw_email)

            subject = email_message["subject"]
            if "Báo cáo vắng học" in subject:
                for part in email_message.walk():
                    if part.get_content_type() == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
                        file_name = part.get_filename()
                        file_data = part.get_payload(decode=True)

                        with open(f"./data/{file_name}", "wb") as file:
                            file.write(file_data)

                        students_data = self.excel_handler.import_data(f"./data/{file_name}")
                        for student in students_data:
                            self.db.add_or_update_student(
                                student["Lớp"], student["Môn học"], student["Họ tên"], student["MSSV"], student["Số buổi vắng"], student["Ngày vắng"]
                            )
