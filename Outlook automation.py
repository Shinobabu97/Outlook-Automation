import win32com.client
import os
import time

def get_signature_html(signature_name):
    """Return the HTML contents of a named Outlook signature."""
    sig_folder = os.path.join(os.environ["APPDATA"], "Microsoft", "Signatures")
    sig_path = os.path.join(sig_folder, f"{signature_name}.htm")

    if os.path.exists(sig_path):
        with open(sig_path, 'r', encoding='utf-8') as f:
            return f.read()
    else:
        print(f"Signature '{signature_name}' not found.")
        return ""

def send_outlook_email(to, cc, subject, body, bcc="", attachments=None, signature_name=None):
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        if not outlook.Session.Accounts:
            return "Outlook is not logged in. Please log in and try again."

        mail = outlook.CreateItem(0)  # Mail item

        mail.To = to
        mail.CC = cc
        mail.BCC = bcc
        mail.Subject = subject

        # Load specific signature
        signature_html = get_signature_html(signature_name) if signature_name else ""

        # Set email body with signature
        mail.HTMLBody = f"<p>{body}</p>{signature_html}"

        # Add attachments if any
        if attachments:
            for file_path in attachments:
                if os.path.exists(file_path):
                    mail.Attachments.Add(file_path)
                else:
                    print(f"Attachment not found: {file_path}")

        mail.Display() #mail.Send() to send automatically
        return "Outlook email draft opened successfully."

    except Exception as e:
        return f"Failed to open Outlook or create email. Error: {str(e)}"

# Example usage
if __name__ == "__main__":
    to = "abc@gmail.com; xyz@gmail.com"
    cc = "qwerty@gmail.com; asdf@gmail.com"
    subject = "Testing Purpose"
    body = "This is a test session using the Shino signature."

    r"""attachments = [
        r"C:\Path\To\File1.pdf",
        r"C:\Path\To\File2.docx"
    ]"""

    result = send_outlook_email(
        to=to,
        cc=cc,
        subject=subject,
        body=body,
        bcc="",
        signature_name="Shino (shino@gmail.com)"  # 👈 use this to load specific signature go and check the name used for the signature here C:\Users\shino\AppData\Roaming\Microsoft\Signatures
    )

    print(result)
