import logging
import os
import sqlite3
import sys
from logging.handlers import RotatingFileHandler

from main_monitor import MailMonitor

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

username = ''
password = ''
allowed_domain = ''
default_printer = ''
very_pdf_folder = ''
clear_files = True

from db import Database


def load_settings():
    global_config = []
    try:
        connection = sqlite3.connect('monitor.db')
        db = Database(connection)
        global_config = db.get_global_config()
    except sqlite3.OperationalError:
        logger.error('Could not connect to database')
    except Exception as e:
        logger.error('Error reading configuration.')
    finally:
        db.close_connection()
    return global_config


if __name__ == '__main__':
    result = load_settings()
    if result:
        logger.info('Monitor started')
        try:
            mm = MailMonitor(username=result["application_user"], password=result["application_password"],
                             default_printer=result["printer"], verypdf_folder=result["verypdf_folder"],
                             allowed_domains=result["allowed_domains"], clean_attachments=result["clean_attachments"],
                             imap_server=result["imap_server"], imap_port=result["imap_port"],
                             smtp_server=result["smtp_server"], smtp_port=result["smtp_port"])
            mm.start_monitor()
        except Exception as e:
            logger.error(e)
    else:
        print('Error al leer archivo de configuración')
        logger.error('Error al leer archivo de configuración')
