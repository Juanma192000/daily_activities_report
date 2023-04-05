import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
from config.config import email_password, email, email_confirm_receiver

def send_email(msg_body):
    recipients = [email_confirm_receiver]
    msg = MIMEMultipart()
    msg['From'] = email
    msg['To'] = ", ".join(recipients)
    msg['Date'] = formatdate(localtime = True)
    msg['Subject'] = 'Trabajo_remoto'
    msg.attach(MIMEText(msg_body))
    mailServer=smtplib.SMTP('smtp-mail.outlook.com', 587)
    mailServer.starttls()
    mailServer.login(email , email_password)
    mailServer.sendmail(email, recipients , msg.as_string())
    print(" \n Email has been sent!")  
    mailServer.quit()