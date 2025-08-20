# smtp_test_hardcoded.py
# 放在 C:\Users\User\activity_management 下執行
import smtplib
from email.message import EmailMessage

# ----- 把下面密碼改成你實際的 App Password（測試用） -----
SMTP_SERVER = "smtp.office365.com"
SMTP_PORT = 587
SMTP_USER = "tanyayu@mpat.org.tw"
SMTP_PASS = "@mpat.org.tw"   # <<--- 把這裡換成你剛產生的 16 位 App password
# ----------------------------------------------------------

# 收件測試用（預設寄給自己）
TO_ADDR = SMTP_USER
SUBJECT = "SMTP test (hardcoded credentials)"
BODY = "This is a test email sent by smtp_test_hardcoded.py"

def main():
    msg = EmailMessage()
    msg["From"] = SMTP_USER
    msg["To"] = TO_ADDR
    msg["Subject"] = SUBJECT
    msg.set_content(BODY)

    print(f"Using SMTP server: {SMTP_SERVER}:{SMTP_PORT}, username: {SMTP_USER}")
    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=30) as s:
            s.set_debuglevel(1)   # prints SMTP dialog for debugging
            s.ehlo()
            # upgrade to TLS on port 587
            s.starttls()
            s.ehlo()
            s.login(SMTP_USER, SMTP_PASS)
            s.send_message(msg)
        print("Sent OK")
    except smtplib.SMTPAuthenticationError as e:
        print("SMTPAuthenticationError:", e)
        print("-> Likely: wrong password or Google rejected credentials (BadCredentials).")
        print("-> If using Google, ensure you used an App Password (enable 2-Step Verification, then create App password).")
    except Exception as e:
        print("SEND FAILED:", repr(e))

if __name__ == "__main__":
    main()
