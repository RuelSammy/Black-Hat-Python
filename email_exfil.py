import smtplib
import time
import win32com.client
import os

smtp_server = 'smtp.example.com'
smtp_port = 587
smtp_acct = os.getenv('SMTP_ACCT', 'tim@example.com')  # Use environment variable for email
smtp_password = os.getenv('SMTP_PASSWORD', 'seKret')  # Use environment variable for password
tgt_accts = ['tim@elsewhere.com']

def plain_email(subject, contents):
    message = f'Subject: {subject}\nFrom: {smtp_acct}\n'
    message += f'To: {", ".join(tgt_accts)}\n\n{contents}'
    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(smtp_acct, smtp_password)
        server.sendmail(smtp_acct, tgt_accts, message)
        server.quit()
        print("Email sent successfully.")
    except Exception as e:
        print(f"Failed to send email: {e}")
    time.sleep(1)

def outlook(subject, contents):
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        message = outlook.CreateItem(0)
        message.DeleteAfterSubmit = True
        message.Subject = subject
        message.Body = contents
        message.To = tgt_accts[0]
        message.Send()
        print("Outlook email sent successfully.")
    except Exception as e:
        print(f"Failed to send Outlook email: {e}")

if __name__ == '__main__':
    plain_email('test2 message', 'attack at dawn.')
