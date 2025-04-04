import imaplib
import logging
import os
import socket
import subprocess
import time
import traceback
import sys
import unidecode
from imap_tools import A, MailBox, MailboxLoginError, MailboxLogoutError
from converter import DocToPDF, PptxToPDF, XlsxToPDF
from logging.handlers import RotatingFileHandler
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

LOG_FILENAME = os.path.join(os.getcwd(), 'monitor.log')
stdout_handler = logging.StreamHandler(stream=sys.stdout)
size_handler = RotatingFileHandler(LOG_FILENAME, backupCount=3, encoding='utf-8')
handlers = [stdout_handler, size_handler]
logging.basicConfig(
    level=logging.INFO,
    format='[%(asctime)s] %(levelname)s - %(message)s',
    handlers=handlers,
)
logger = logging.getLogger('mailmonitor')

CURRENT_PATH = os.getcwd() + '/'


class MailMonitor():


    def __init__(self, username, password, default_printer, verypdf_folder, allowed_domains, clean_attachments, imap_server, imap_port):
        self.password = password
        self.username = username
        self.default_printer = default_printer
        self.verypdf_folder = verypdf_folder
        self.allowed_domain = [x.strip() for x in allowed_domains.split(',')]
        self.clean_attachments = clean_attachments
        self.imap_server = imap_server
        self.imap_port = imap_port

    def send_confirmation_mail(self, to, filelist):
        logger.info(f'Enviando correo de confirmación a {to}...')
        sender = self.username
        receivers = [to]
        fl = [file[1] for file in filelist]
        files = '\n'.join(fl)
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
        <p>No respondas este correo ya que será ignorado.</p>
        </body>
        </html>
        """

        msg_body = MIMEText(message, 'html')
        msg.attach(msg_body)

        try:
            smtpObj = smtplib.SMTP('smtp.gmail.com', 587)
            smtpObj.ehlo()
            smtpObj.starttls()
            smtpObj.login(self.username, self.password)
            smtpObj.sendmail(sender, receivers, msg.as_string())
            smtpObj.quit()
        except Exception as e:
            logger.error(e)




    def delete_files(self, t_file):
        del_dir = os.path.join('', t_file[3])
        try:
            pObj = subprocess.Popen('rmdir /S /Q %s' % del_dir, shell=True, stdout=subprocess.PIPE,
                                    stderr=subprocess.PIPE)
            # recreate the deleted parent dir in case it get deleted
        except Exception as e:
            logger.error(f'{e}')

    def download_file(self, filename, msguser, msgdate, payload):
        path = os.path.join(CURRENT_PATH, msguser)
        if not os.path.exists(path):
            os.mkdir(path)
        hora = '0' + str(msgdate.hour) if msgdate.hour < 10 else str(msgdate.hour)
        minutos = '0' + str(msgdate.minute) if msgdate.minute < 10 else str(msgdate.minute)
        segundos = '0' + str(msgdate.second) if msgdate.second < 10 else str(msgdate.second)
        dia = '0' + str(msgdate.day) if msgdate.day < 10 else str(msgdate.day)
        mes = '0' + str(msgdate.month) if msgdate.month < 10 else str(msgdate.month)
        anio = str(msgdate.year)
        att_filename = msguser + hora + minutos + segundos + dia + mes + anio + filename
        att_filename = unidecode.unidecode(att_filename)
        att_filename = os.path.join(path, att_filename)
        try:
            with open(att_filename, 'wb') as f:
                f.write(payload)
                logger.debug(f'Archivo guardado en {att_filename}')
            return True, att_filename
        except Exception as e:
            return False, ''

    def print_docx_file(self, filename, user, printer, jobname):
        logger.info(f'Convirtiendo Word {filename}')
        obj = DocToPDF(doc_file_path=filename)
        ok, filepath = obj.convert_docx_to_pdf()
        logger.debug(f'{ok} {filepath}')
        if ok:
            self.print_file(filepath, user, printer, jobname=jobname)
            logger.info(f'{filename} Convertido correctamente {filepath}')
        else:
            logger.error(f'Error al convertir {filename}  -> {filepath}')

    def print_pptx_file(self, filename, user, printer, jobname):
        logger.info(f'Convirtiendo PowerPoint {filename}')
        obj = PptxToPDF(pp_file_path=filename)
        ok, filepath = obj.convert_pptx_to_pdf()
        logger.debug(f'{ok} {filepath}')
        if ok:
            self.print_file(filepath, user, printer, jobname=jobname)
            logger.debug(f'{filename} Convertido correctamente {filepath}')
        else:
            logger.error(f'Error al convertir {filename}  -> {filepath}')

    def print_excel_file(self, filename, user, printer, jobname):
        logger.info(f'Convirtiendo Excel {filename}')
        obj = XlsxToPDF(xlsx_file_path=filename)
        ok, filepath = obj.convert_xlsx_to_pdf()
        logger.debug(f'{ok} {filepath}')
        if ok:
            self.print_file(filepath, user, printer, jobname=jobname)
            logger.info(f'{filename} Convertido correctamente {filepath}')
        else:
            logger.error(f'Error al convertir {filename}  -> {filepath}')

    def print_file(self, filename, user, printer, jobname=None):
        # if jobname:
        #     jobdocname = jobname[:-5]
        #     cmd = f'{self.verypdf_folder} -jobdocname "{jobdocname}" -printer "{printer}" -jobusername "{user}" "{filename}"'
        # else:
        nombre, extension = os.path.splitext(jobname)
        cmd = f'{self.verypdf_folder} -printer "{printer}" -jobusername "{user}" -jobdocname "{nombre}" "{filename}"'
        logger.info(f'Imprimiendo {filename}')
        logger.debug(f'{cmd}')
        subprocess.run(cmd)

    def start_monitor(self):
        logger.info(f'Monitoreo iniciado en {self.username}')
        done = False
        logger.info(f'Allowed domains{self.allowed_domain}')
        logger.info(f'Default printer {self.default_printer}')
        logger.info(f'Clear files {self.clean_attachments}')
        logger.info(f'user {self.username}')
        login_attempts = 0
        while not done:
            connection_start_time = time.monotonic()
            connection_live_time = 0.0
            try:
                with MailBox(self.imap_server).login(self.username, self.password) as mailbox:
                    logger.info(f'@@ new connection {time.asctime()}')
                    login_attempts = 0
                    while connection_live_time < 29 * 60:
                        try:
                            responses = mailbox.idle.wait(timeout=3 * 60)
                            logger.debug(f'IDLE responses: {responses}')
                            if responses:
                                for msg in mailbox.fetch(A(seen=False)):
                                    msg_domain = msg.from_.split('@')[1]
                                    msg_user = msg.from_.split('@')[0]
                                    msg_to = msg.from_
                                    if msg_domain in self.allowed_domain or self.allowed_domain == []:
                                        pdf_list = []
                                        img_list = []
                                        doc_list = []
                                        excel_list = []
                                        pp_list = []
                                        file_list = []
                                        logger.info(f'-> {msg.date} {msg.subject} {msg.from_} {msg.uid}')
                                        for att in msg.attachments:
                                            type = att.content_type
                                            file_name, file_extension = os.path.splitext(att.filename)
                                            file_extension = file_extension.lower()
                                            result, filename = self.download_file(filename=att.filename,
                                                                                  msguser=msg_user,
                                                                                  msgdate=msg.date,
                                                                                  payload=att.payload)
                                            if result:
                                                file_list.append((filename, att.filename, file_extension, msg_user))
                                                if file_extension in ['.pdf']:
                                                    pdf_list.append((filename, att.filename))
                                                elif file_extension in ['.jpg', '.jpeg', '.tif', '.tiff', '.gif',
                                                                        '.png', '.pcx']:
                                                    img_list.append((filename, att.filename))
                                                elif file_extension in ['.xls', '.xlsx', '.xlsm']:
                                                    excel_list.append((filename, att.filename))
                                                elif file_extension in ['.doc', '.docx', '.rtf', '.txt', '.xml']:
                                                    doc_list.append((filename, att.filename))
                                                elif file_extension in ['.ppt', '.pptx', '.pps', '.ppsx']:
                                                    pp_list.append((filename, att.filename))
                                                else:
                                                    logger.warning(f'Tipo no admitido {type}')
                                        for pdf in pdf_list:
                                            self.print_file(filename=pdf[0], user=msg_user,
                                                            printer=self.default_printer,
                                                            jobname=f'{msg_user}-{pdf[1]}')
                                        for img in img_list:
                                            self.print_file(filename=img[0], user=msg_user,
                                                            printer=self.default_printer,
                                                            jobname=f'{msg_user}-{img[1]}')
                                        for excel in excel_list:
                                            self.print_excel_file(filename=excel[0], user=msg_user,
                                                                  printer=self.default_printer,
                                                                  jobname=f'{msg_user}-{excel[1]}')
                                        for word in doc_list:
                                            self.print_docx_file(filename=word[0], user=msg_user,
                                                                 printer=self.default_printer,
                                                                 jobname=f'{msg_user}-{word[1]}')
                                        for pp in pp_list:
                                            self.print_pptx_file(filename=pp[0], user=msg_user,
                                                                 printer=self.default_printer,
                                                                 jobname=f'{msg_user}-{pp[1]}')
                                        if self.clean_attachments:
                                            for t_file in file_list:
                                                logger.info(f'Limpiando el directorio ----- {t_file[3]}')
                                                self.delete_files(t_file)
                                            if file_list:
                                                self.send_confirmation_mail(msg_to,file_list)
                                    else:
                                        logger.warning(
                                            f'Mail from not allowed domain-> {msg.date} {msg.subject} {msg.from_} {msg.uid}')
                            else:
                                logger.info('No hay nada nuevo')
                        except KeyboardInterrupt:
                            logger.error('~KeyboardInterrupt')
                            done = True
                            break
                        except FileNotFoundError:
                            logger.error('Error al guardar el archivo')
                        except Exception as e:
                            logger.error(e)
                            logger.error('Reconectando en 1 minúto...')
                            mailbox.idle.stop()
                            time.sleep(60)
                            connection_live_time = 2000
                            continue
                        connection_live_time = time.monotonic() - connection_start_time
            except (TimeoutError, ConnectionError,
                    imaplib.IMAP4.abort, MailboxLoginError, MailboxLogoutError,
                    socket.herror, socket.gaierror, socket.timeout) as e:
                logger.error(f'## Error en la comunicación con el servidor de correo. Reconectando en 1 minuto.')
                time.sleep(60)
                continue
            except imaplib.IMAP4.error:
                logger.error('Error al iniciar sesión. Verifique la contraseña y el correo.')


