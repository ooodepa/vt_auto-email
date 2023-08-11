import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from tkinter import filedialog
import ssl


def get_mail(sender, subject):
    html_file_path = filedialog.askopenfilename(
        filetypes=[("HTML Files", "*.html")],
        initialdir=os.path.join(os.getcwd(), "HTML Mails"),
        title="HTML")

    with open(html_file_path, "r", encoding="utf-8") as file:
        html_content = file.read()

    message = MIMEMultipart()
    message["From"] = sender
    message["Subject"] = subject

    message.attach(MIMEText(html_content, "html"))

    files_dir_path = os.path.join(os.getcwd(), "Files")

    for file in os.listdir(files_dir_path):
        file_path = os.path.join(files_dir_path, file)
        with open(file_path, "rb") as f:
            part = MIMEApplication(f.read())
            part.add_header("Content-Disposition", f"attachment; filename={file}")
            message.attach(part)

    return message


def send_email(sender, receivers, password, message, port=465):
    smtp_server = "smtp." + sender.split('@')[-1]

    # Create an SSL context to use SSL for SMTP
    context = ssl.create_default_context()

    with smtplib.SMTP_SSL(host=smtp_server, port=port, context=context) as server:
        server.login(sender, password)
        for receiver in receivers:
            message["To"] = receiver
            server.send_message(message)
