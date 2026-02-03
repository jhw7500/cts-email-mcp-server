import poplib
import smtplib
import email
import os
import sys
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from typing import List, Dict, Optional, Any
from . import utils

class EmailClient:
    def __init__(self):
        self.server = "pop3.hiworks.co.kr"
        self.port = 995
        self.user = os.environ.get("EMAIL_USER")
        self.password = os.environ.get("EMAIL_PASSWORD")
        
        if not self.user or not self.password:
            raise ValueError("EMAIL_USER 및 EMAIL_PASSWORD 환경 변수가 설정되어야 합니다.")
        self.connection = None

    def connect(self):
        try:
            self.connection = poplib.POP3_SSL(self.server, self.port)
            self.connection.user(self.user)
            self.connection.pass_(self.password)
            return True
        except Exception as e:
            print(f"연결 실패 ({self.user}): {e}", file=sys.stderr)
            return False

    def disconnect(self):
        if self.connection:
            try: self.connection.quit()
            except: pass
            self.connection = None

    def list_emails(self, count=10):
        if not self.connect(): return []
        try:
            num_messages = len(self.connection.list()[1])
            start = max(1, num_messages - count + 1)
            results = []
            for i in range(num_messages, start - 1, -1):
                try:
                    resp, lines, _ = self.connection.top(i, 0)
                    msg = email.message_from_bytes(b'\r\n'.join(lines))
                    results.append({
                        "id": i,
                        "date": utils.clean_header(msg.get("Date")),
                        "from": utils.clean_header(msg.get("From")),
                        "subject": utils.clean_header(msg.get("Subject"))
                    })
                except: continue
            return results
        finally: self.disconnect()

    def search_emails(self, keyword: str, limit: int = 50):
        if not self.connect(): return []
        try:
            num_messages = len(self.connection.list()[1])
            start = max(1, num_messages - limit + 1)
            results = []
            for i in range(num_messages, start - 1, -1):
                try:
                    resp, lines, _ = self.connection.top(i, 0)
                    msg = email.message_from_bytes(b'\r\n'.join(lines))
                    subj = utils.clean_header(msg.get("Subject"))
                    frm = utils.clean_header(msg.get("From"))
                    if keyword.lower() in subj.lower() or keyword.lower() in frm.lower():
                        results.append({
                            "id": i,
                            "date": utils.clean_header(msg.get("Date")),
                            "from": frm,
                            "subject": subj
                        })
                except: continue
            return results
        finally: self.disconnect()

    def get_email(self, email_id: int):
        if not self.connect(): return None
        try:
            _, lines, _ = self.connection.retr(email_id)
            msg = email.message_from_bytes(b'\r\n'.join(lines))
            return {
                "id": email_id,
                "subject": utils.clean_header(msg.get("Subject")),
                "from": utils.clean_header(msg.get("From")),
                "date": utils.clean_header(msg.get("Date")),
                "body": utils.get_body(msg),
                "attachments": utils.get_attachment_list(msg)
            }
        finally: self.disconnect()

    def download_file(self, email_id: int, filename: str, save_path: str):
        if not self.connect(): return "Error: Connection failed"
        try:
            _, lines, _ = self.connection.retr(email_id)
            msg = email.message_from_bytes(b'\r\n'.join(lines))
            for part in msg.walk():
                part_filename = part.get_filename()
                if part_filename and utils.clean_header(part_filename) == filename:
                    if not os.path.exists(save_path): os.makedirs(save_path)
                    filepath = os.path.join(save_path, filename)
                    with open(filepath, 'wb') as f:
                        f.write(part.get_payload(decode=True))
                    return f"Success: {filename} saved to {filepath}"
            return "Error: File not found"
        finally: self.disconnect()

    def send_email(self, to_email: str, subject: str, body: str):
        """SMTP를 통해 이메일을 발송합니다."""
        smtp_server = "smtp.hiworks.co.kr"
        smtp_port = 465
        
        try:
            msg = MIMEMultipart()
            msg['From'] = self.user
            msg['To'] = to_email
            msg['Subject'] = subject
            msg.attach(MIMEText(body, 'plain'))
            
            with smtplib.SMTP_SSL(smtp_server, smtp_port) as server:
                server.login(self.user, self.password)
                server.send_message(msg)
            return f"Success: 이메일이 {to_email}에게 성공적으로 발송되었습니다."
        except Exception as e:
            return f"Error: 이메일 발송 실패 - {str(e)}"