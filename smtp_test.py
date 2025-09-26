# smtp_test.py
import os, smtplib
from email.message import EmailMessage

SMTP_USER = os.environ.get("SMTP_USERNAME") or os.environ.get("SMTP_USER")
SMTP_PASSWORD = os.environ.get("SMTP_PASSWORD")
print("SMTP_USER:", SMTP_USER)
print("SMTP_PASSWORD length:", len(SMTP_PASSWORD) if SMTP_PASSWORD else "None")

msg = EmailMessage()
msg["From"] = SMTP_USER
msg["To"] = SMTP_USER
msg["Subject"] = "SMTP debug test"
msg.set_content("test")

try:
    with smtplib.SMTP("smtp.gmail.com", 587, timeout=30) as s:
        s.set_debuglevel(1)   # <<-- 這會把 SMTP 交易印出來（很重要）
        s.ehlo()
        s.starttls()
        s.ehlo()
        s.login(SMTP_USER, SMTP_PASSWORD)
        s.send_message(msg)
    print("Sent OK")
except Exception as e:
    print("ERROR:", repr(e))