import logging

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

logger = logging.getLogger(__name__)

username = 'jonas4784@gmail.com'
password = 'ohiybdlbutjsptoa'

to = 'jonas4784@gmail.com'
filelist = ['a.pdf', 'b.pdf', 'c.pdf']
logger.info(f'Enviando correo de confirmación a {to} {filelist}...')
sender = 'jonas4784@gmail.com'
receivers = ['jonas4784@gmail.com']
files = '\n'.join(filelist)
msg = MIMEMultipart('alternative')
msg['From'] = f'From: {sender}'
msg['To'] = f'To <{to}>'
msg['Subject'] = 'Confirmación de impresión'
message = f"""
        <html>
        <head></head>
        <body>
        <p>Se han impreso tus documentos</p>
        {files}
        </body>
        </html>
        """

msg_body = MIMEText(message, 'html')
msg.attach(msg_body)

try:
    smtpObj = smtplib.SMTP('smtp.gmail.com', 587)
    smtpObj.ehlo()
    smtpObj.starttls()
    smtpObj.login(username, password)
    smtpObj.sendmail(sender, receivers, msg.as_string())
    smtpObj.quit()
except Exception as e:
    logger.error(e)
