import logging
import os
import shlex
import sqlite3
import subprocess
import sys
from logging.handlers import RotatingFileHandler
from tkinter import Tk, Widget, Button, Label, Entry, StringVar, BooleanVar, ttk, filedialog, IntVar, WORD, \
    END
from tkinter.scrolledtext import ScrolledText

import win32print
from PIL import ImageTk, Image
from pyexpat.errors import messages

from db import Database

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


class GUI(Tk):
    current_col = 0
    current_row = 0
    DEFAULT_PAD_X = 5
    DEFAULT_PAD_Y = 3
    DEFAULT_FILL = 'x'

    def add_component_grid(self, component: Widget) -> None:
        component.grid(row=self.current_row, column=self.current_col)
        self.current_col += 1
        # self.current_row += 1

    def add_component_pack(self, component: Widget, padx: int = None, pady: int = None, fill: str = None) -> None:
        pad_x = padx or self.DEFAULT_PAD_X
        pad_y = pady or self.DEFAULT_PAD_Y
        fill = fill or self.DEFAULT_FILL
        component.pack(expand=False, fill=fill, padx=pad_x, pady=pad_y)


class Container(ttk.Frame):
    current_col = 0
    current_row = 0
    DEFAULT_PAD_X = 5
    DEFAULT_PAD_Y = 3
    DEFAULT_FILL = 'x'

    def add_component_pack(self, component: Widget, padx: int = None, pady: int = None, fill: str = None) -> None:
        pad_x = padx or self.DEFAULT_PAD_X
        pad_y = pady or self.DEFAULT_PAD_Y
        fill = fill or self.DEFAULT_FILL
        component.pack(expand=False, fill=fill, padx=pad_x, pady=pad_y)

    def add_component_grid(self, component: Widget, col: int = None, row: int = None) -> None:
        current_row = 0
        current_col = 0
        if col is not None:
            if col >= 0:
                current_col = col
        else:
            current_col = self.current_col
        if row is not None:
            if row >= 0:
                current_row = row
        else:
            current_row = self.current_row
        component.grid(row=current_row, column=current_col, padx=self.DEFAULT_PAD_X, pady=self.DEFAULT_PAD_Y)
        self.current_col = current_col + 1
        self.current_row = current_row


def load_printers():
    logger.debug('loading printers...')
    load_system_printers()


def load_system_printers():
    global printer_options
    printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL, None, 1)
    printers += win32print.EnumPrinters(win32print.PRINTER_ENUM_CONNECTIONS, None, 1)
    printer_options = [x[2] for x in printers]
    entry_printer.configure(
        values=printer_options,
    )


def load_verypdf_folder():
    verypdf_path = filedialog.askopenfilename()  # Open a folder selection dialog
    if verypdf_path:
        var_verypdf_folder.set(verypdf_path)
    else:
        var_verypdf_folder.set('')


def stop_app():
    logger.info('Finalizado')
    sys.exit()


def show_password():
    global SHOWED_CHARACTER
    if SHOWED_CHARACTER == '*':
        SHOWED_CHARACTER = ''
        button_show_password.configure(
            image=image_hide_password.__str__()
        )
    else:
        SHOWED_CHARACTER = '*'
        button_show_password.configure(
            image=image_show_password.__str__()
        )
    entry_password.configure(show=SHOWED_CHARACTER)


def save_settings():
    stop_monitor()
    connection = sqlite3.connect('monitor.db')
    db = Database(connection)
    try:
        id_provider = [x for x in PROVIDERS if x[1].lower() == var_provider.get().lower()]
        if id_provider:
            id_provider = id_provider[0][0]
            db.insert_or_update_oauth_provider(var_provider.get(), var_imap_server.get(), var_imap_port.get())
        else:
            id_provider = None
        db.create_or_update_settings([
            ("verypdf_folder", var_verypdf_folder.get()),
            ("application_user", var_email.get()),
            ("application_password", var_password.get()),
            ("allowed_domains", var_allowed_domains.get()),
            ("printer", var_default_printer.get()),
            ("clean_attachments", checkbox_clean_value.get()),
            ("active", 1),
            ("provider_id", id_provider)
        ])
    except sqlite3.IntegrityError:
        logger.error('Error creating settings')
    finally:
        db.close_connection()


def start_service():
    global procces_id
    logger.info('Starting service...')
    error = save_settings()
    if error:
        logger.error('Configure all settings')
    else:
        try:
            cmd = "pcountermailmonitorgnsys.exe"
            path_mm = os.path.join('mon',cmd)
            cmds = shlex.split(cmd)
            logger.info(path_mm)
            pid = subprocess.Popen(path_mm, start_new_session=True, creationflags=subprocess.DETACHED_PROCESS)
            logger.debug(f'Monitor started with PID {pid.pid}')
        except Exception as e:
            logger.error(e)



def stop_monitor():
    logger.info('Stopping monitor...')
    stop_command = 'TASKKILL /F /IM pcountermailmonitorgnsys.exe'
    result = subprocess.run(stop_command, capture_output=True, text=True)
    if result.stderr:
        logger.error('Error stopping monitor.')
    else:
        logger.debug(f'Se finalizó el monitoreo')


def on_change_provider(index, value, op):
    try:
        imap_server_conf = [x for x in PROVIDERS if x[1].lower() == var_provider.get().lower()]
        imap_server = imap_server_conf[0][2] if imap_server_conf[0][2] else ''
        imap_port = imap_server_conf[0][3] if imap_server_conf[0][3] else 0
        var_imap_server.set(imap_server)
        var_imap_port.set(imap_port)
    except IndexError as ie:
        logger.error('Error loading providers...')


def load_settings(global_settings):
    if global_settings:
        var_verypdf_folder.set(global_settings["verypdf_folder"])
        var_email.set(global_settings["application_user"])
        var_password.set(global_settings["application_password"])
        var_allowed_domains.set(global_settings["allowed_domains"])
        var_default_printer.set(global_settings["printer"])
        checkbox_clean_value.set(global_settings["clean_attachments"])
        var_imap_server.set(global_settings["imap_server"])
        var_imap_port.set(global_settings["imap_port"])
        var_provider.set(global_settings["provider_name"])


def get_providers():
    global PROVIDERS, PROVIDER_VALUES
    connexion = sqlite3.connect('monitor.db')
    db = Database(connexion)
    providers_result = db.get_oauth_providers()
    PROVIDERS = providers_result
    PROVIDER_VALUES = [provider[1] for provider in providers_result]
    entry_provider.configure(values=PROVIDER_VALUES)


def show_logs():
    global log_length
    log_file = open(LOG_FILENAME, 'r', encoding="utf8")
    lines = log_file.readlines()
    new_log_length = len(lines)
    if new_log_length > log_length:
        for line in lines[log_length:new_log_length]:
            tag = 'DEBUG' if 'DEBUG' in line else 'WARNING' if 'WARNING' in line else 'ERROR' if 'ERROR' in line else 'INFO'
            log_display.insert(END, line, tag)
            log_length = new_log_length
        log_display.yview(END)
    container_logs.after(1000, show_logs)


if __name__ == '__main__':
    log_length = 0
    # Constantes
    WINDOW_WIDTH = 580
    WINDOW_HEIGHT = 380
    ENTRY_WIDTH = 140
    LABEL_WIDTH = 15
    BUTTON_WIDTH = 16
    SQUARE_BUTTON_WIDTH = 15
    SQUARE_BUTTON_HEIGHT = 15
    SHOWED_CHARACTER = '*'
    PROVIDERS = []
    PROVIDER_VALUES = []


    printers = []
    root = GUI()
    root.resizable(False, False)
    root.title('Monitor de correo')
    tab_control = ttk.Notebook()
    tab_imap = Container()

    tab_control.add(tab_imap, text='Configuración IMAP')

    # Variables
    var_provider = StringVar()
    var_imap_server = StringVar()
    var_imap_port = IntVar()
    var_provider.trace('w', on_change_provider)
    var_email = StringVar()
    var_password = StringVar()
    var_default_printer = StringVar()
    var_verypdf_folder = StringVar()
    var_allowed_domains = StringVar()
    checkbox_clean_value = BooleanVar()
    # Sección proveedor IMAP

    container_provider = Container()
    lbl_provider = Label(container_provider, text='Proveedor IMAP:', width=LABEL_WIDTH)
    entry_provider = ttk.Combobox(container_provider, values=PROVIDER_VALUES, state='readonly',
                                  width=ENTRY_WIDTH - 3,
                                  textvariable=var_provider)
    container_provider.add_component_grid(lbl_provider)
    container_provider.add_component_grid(entry_provider)

    # Sección configuracion IMAP
    container_imap = Container()
    lbl_imap_server = Label(container_imap, text='Server:', width=(int(LABEL_WIDTH / 2) - 3))
    entry_imap_server = Entry(container_imap, width=(int(ENTRY_WIDTH / 2) + 3), textvariable=var_imap_server)
    lbl_imap_port = Label(container_imap, text='Port:', width=(int(LABEL_WIDTH / 2) - 3))
    entry_imap_port = Entry(container_imap, width=int(ENTRY_WIDTH / 2), textvariable=var_imap_port)
    container_imap.add_component_grid(lbl_imap_server)
    container_imap.add_component_grid(entry_imap_server)
    container_imap.add_component_grid(lbl_imap_port)
    container_imap.add_component_grid(entry_imap_port)

    # Sección Correo
    container_email = Container()
    lbl_email = Label(container_email, text='Usuario:', width=LABEL_WIDTH)
    entry_email = Entry(container_email, width=ENTRY_WIDTH, textvariable=var_email)
    container_email.add_component_grid(lbl_email)
    container_email.add_component_grid(entry_email)

    # Sección Password
    container_password = Container()
    label_password = Label(container_password, text='Password:', width=LABEL_WIDTH)
    entry_password = Entry(container_password, show=SHOWED_CHARACTER, width=ENTRY_WIDTH, textvariable=var_password)
    image_show_password = ImageTk.PhotoImage(
        Image.open('resources/img/view.png').resize((SQUARE_BUTTON_WIDTH, SQUARE_BUTTON_HEIGHT)))
    image_hide_password = ImageTk.PhotoImage(
        Image.open('resources/img/hide.png').resize((SQUARE_BUTTON_WIDTH, SQUARE_BUTTON_HEIGHT)))
    button_show_password = Button(container_password, text='', command=show_password,
                                  image=image_show_password.__str__())
    container_password.add_component_grid(label_password)
    container_password.add_component_grid(entry_password)
    container_password.add_component_grid(button_show_password)

    # Sección Impresora
    container_printer = Container()
    label_printer = Label(container_printer, text='Impresora', width=LABEL_WIDTH)
    # ttk.Combobox(tab_1, state='readonly', values=PRINTERS, width=ENTRYWIDTH - 3, textvariable=var_default_printer)
    entry_printer = ttk.Combobox(container_printer, state='readonly', width=ENTRY_WIDTH - 3,
                                 textvariable=var_default_printer)
    image_reload_printers = ImageTk.PhotoImage(
        Image.open('resources/img/refresh.png').resize((SQUARE_BUTTON_WIDTH, SQUARE_BUTTON_HEIGHT)))
    button_refresh_printers = Button(container_printer, text='', command=load_printers,
                                     image=image_reload_printers.__str__())
    container_printer.add_component_grid(label_printer)
    container_printer.add_component_grid(entry_printer)
    container_printer.add_component_grid(button_refresh_printers)

    # Sección VeryPDF
    container_very_pdf = Container()
    label_very_pdf = Label(container_very_pdf, text='Very PDF', width=LABEL_WIDTH)
    entry_very_pdf = Entry(container_very_pdf, width=ENTRY_WIDTH, textvariable=var_verypdf_folder)
    image_folder = ImageTk.PhotoImage(
        Image.open('resources/img/folder.png').resize((SQUARE_BUTTON_WIDTH, SQUARE_BUTTON_HEIGHT)))
    button_open_very_pdf = Button(container_very_pdf, text='', command=load_verypdf_folder,
                                  image=image_folder.__str__())
    container_very_pdf.add_component_grid(label_very_pdf)
    container_very_pdf.add_component_grid(entry_very_pdf)
    container_very_pdf.add_component_grid(button_open_very_pdf)

    # Sección Dominios
    container_domains = Container()
    label_domains = Label(container_domains, text='Dominios Permitidos:', width=LABEL_WIDTH)
    entry_domains = Entry(container_domains, width=ENTRY_WIDTH, textvariable=var_allowed_domains)
    container_domains.add_component_grid(label_domains)
    container_domains.add_component_grid(entry_domains)

    # Sección Check borrar adjuntos
    container_check = Container()
    checkbox_clean = ttk.Checkbutton(container_check, text='Limpiar archivos adjuntos automáticamente',
                                     variable=checkbox_clean_value)
    container_check.add_component_grid(checkbox_clean)

    # seccion mensajes
    container_messages = Container()
    label_msg = Label(container_messages, text='Messages:', width=LABEL_WIDTH)
    container_messages.add_component_grid(label_msg)

    # Sección logs
    container_logs = Container()
    log_display = ScrolledText(container_logs, wrap=WORD, height=10, width=ENTRY_WIDTH-18, background='black')
    log_display.tag_config('DEBUG', foreground='blue')
    log_display.tag_config('WARNING', foreground='yellow')
    log_display.tag_config('ERROR', foreground='red')
    log_display.tag_config('CRITICAL', foreground='red')
    log_display.tag_config('INFO', foreground='#74DE57')
    container_logs.add_component_grid(log_display)
    container_logs.after(100, show_logs)

    # Sección botones
    container_buttons = Container()
    button_stop_monitor = Button(container_buttons, command=stop_monitor, text='Detener Monitoreo', width=BUTTON_WIDTH)
    button_start_monitor = Button(container_buttons, command=start_service, text='Iniciar Monitoreo',
                                  width=BUTTON_WIDTH)
    button_exit = Button(container_buttons, text='Salir', width=BUTTON_WIDTH, command=stop_app)
    button_save = Button(container_buttons, command=save_settings, text='Guardar', width=BUTTON_WIDTH)
    container_buttons.add_component_grid(button_stop_monitor)
    container_buttons.add_component_grid(button_start_monitor)
    container_buttons.add_component_grid(button_exit)
    container_buttons.add_component_grid(button_save)

    root.add_component_pack(tab_control)
    root.add_component_pack(container_provider)
    root.add_component_pack(container_imap)
    root.add_component_pack(container_email)
    root.add_component_pack(container_password)
    root.add_component_pack(container_printer)
    root.add_component_pack(container_very_pdf)
    root.add_component_pack(container_domains)
    root.add_component_pack(container_check)
    root.add_component_pack(container_messages)
    root.add_component_pack(container_logs)
    root.add_component_pack(container_buttons)
    load_printers()
    try:
        get_providers()
    except sqlite3.OperationalError as e:
        logger.error(e)
    connexion = sqlite3.connect("monitor.db")
    db = Database(connexion)
    try:
        db.create_database()
    except sqlite3.OperationalError:
        logger.error('Error al crear la base de datos')
    finally:
        db.close_connection()
    connexion = sqlite3.connect("monitor.db")
    db = Database(connexion)
    try:
        global_config = db.get_global_config()
        load_settings(global_config)
    except sqlite3.OperationalError as oe:
        logger.error(oe)
        logger.error('No existe la configuración')
    except IndexError as ie:
        logger.error(ie)
        logger.error('No existe la configuración')
    finally:
        db.close_connection()
    root.mainloop()
