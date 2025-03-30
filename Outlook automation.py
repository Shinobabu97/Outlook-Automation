import win32com.client
import time
import os

def send_outlook_email(to_recipients, cc_recipients, bcc_recipients, subject, body, attachments=None):
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")

        if not outlook.Session.Accounts:
            return "Outlook is not logged in. Please log in and try again."

        mail = outlook.CreateItem(0)  # 0 = Mail Item

        # Set basic fields
        mail.To = to_recipients
        mail.CC = cc_recipients
        mail.BCC = bcc_recipients
        mail.Subject = subject

        # Display to load signature
        mail.Display()
        time.sleep(1)
        signature = mail.HTMLBody

        # Set body with signature
        mail.HTMLBody = f"<p>{body}</p>{signature}"

        # Attach files if provided
        if attachments:
            for file_path in attachments:
                if os.path.exists(file_path):
                    mail.Attachments.Add(file_path)
                else:
                    print(f"Attachment not found: {file_path}")

        mail.Send() #mail.Display() to send manually
        return "Email sent successfully."

    except Exception as e:
        return f"Failed to send email. Error: {str(e)}"

# MAIN CALL
if __name__ == "__main__":
    to = "abc@gmail.com; xyz@gmail.com"
    cc = "qwerty@gmail.com; asdf@gmail.com"
    bcc = "hiddenemail@example.com"  # Optional. If not required, leave the string empty
    subject = "Testing Purpose"
    body = "This is a test session with attachments and BCC."

    # Optional: list of file paths
    attachments = [
        r"C:\Path\To\Your\File1.pdf",
        r"C:\Path\To\Your\File2.png"
    ]

    result = send_outlook_email(to, cc, bcc, subject, body, attachments)
    print(result)
