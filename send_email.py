import mimetypes
import os
import smtplib
from email.message import EmailMessage
from typing import Iterable


def send_mail(send_from: str, send_to: list, subject: str, body: str, attachments: Iterable[str], smtp_url: str, port: int, username: str, password: str, starttls=True):

    # Create the email message
    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = send_from
    msg['To'] = ', '.join(send_to)
    # Set email content
    msg.set_content(body)

    for filepath in attachments:
        filename = os.path.basename(filepath)

        if os.path.exists(filepath):
            ctype, encoding = mimetypes.guess_type(filepath)
            if ctype is None or encoding is not None:
                # No guess could be made, or the file is encoded (compressed), so
                # use a generic bag-of-bits type.
                ctype = 'application/octet-stream'
            maintype, subtype = ctype.split('/', 1)
            # Add email attachment
            with open(filepath, 'rb') as fp:
                msg.add_attachment(fp.read(),
                                   maintype=maintype,
                                   subtype=subtype,
                                   filename=filename)

    smtp = smtplib.SMTP(smtp_url, port)
    if starttls:
        smtp.starttls()
    smtp.login(username, password)
    smtp.send_message(msg)
    smtp.quit()


def check_emails_available_outlook(outlook, email: str) -> bool:
    if not outlook.Session.Accounts[email]:
        print(f"{email}: NOT found on Outlook.")
        return False
    return True


def send_mail_outlook(outlook, send_from: str, send_to: list, subject: str, body: str, attachments: Iterable[str]):
    account = outlook.Session.Accounts[send_from]
    # construct the email item object
    message = outlook.CreateItem(0)
    message.Subject = subject
    message.Body = body
    message.To = ";".join(send_to)
    message._oleobj_.Invoke(*(64209, 0, 8, 0, account))
    for attachment in attachments:
        abs_path = os. path. abspath(attachment)
        print(abs_path)
        message.Attachments.Add(abs_path)
    message.Save()
    message.Send()
