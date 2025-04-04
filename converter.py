import logging
import os
import sys
from logging.handlers import RotatingFileHandler

import win32com.client
from PIL import Image


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
wdFormatPDF = 17
pptFormatPDF = 32
xlsxFormatPDF = 47


class DocToPDF():
    doc_file_path = ''
    ok = None

    def __init__(self, doc_file_path):
        self.doc_file = doc_file_path
        self.ok = True

    def convert_docx_to_pdf(self):
        output = self.doc_file.replace('.docx', '.pdf')
        output = output.replace('.doc', '.pdf')
        output = output.replace('.txt', '.pdf')
        word = win32com.client.Dispatch('Word.Application')
        try:
            logger.info(f'Convirtiendo archivo {self.doc_file} a {output}')
            # convert(self.doc_file, output)
            doc = word.Documents.Open(self.doc_file)
            doc.SaveAs(output, FileFormat=wdFormatPDF)
            doc.Close()
            word.Quit()
            logger.info(f'Archivo convertido con éxito {output}')
        except Exception as e:
            word.Quit()
            self.ok = False
            logger.error(f'error al convertir {self.doc_file}')
        return self.ok, output


class PptxToPDF():
    pp_file_path = ''
    ok = None

    def __init__(self, pp_file_path):
        self.pp_file_path = pp_file_path
        self.ok = True

    def convert_pptx_to_pdf(self):
        output = self.pp_file_path.replace('.pptx', '.pdf')
        output = output.replace('.ppt', '.pdf')
        pp = win32com.client.Dispatch('Powerpoint.Application')
        try:
            logger.info(f'Convirtiendo archivo {self.pp_file_path} a {output}')
            pres = pp.Presentations.Open(self.pp_file_path)
            pres.SaveAs(output, FileFormat=pptFormatPDF)
            pres.Close()
            pp.Quit()
            # docx2pdf.convert(self.doc_file, output)
            logger.info(f'Archivo convertido con éxito {output}')
        except Exception as e:
            self.ok = False
            pp.Quit()
            logger.error(e)
            logger.error(f'error al convertir {self.pp_file_path}')
        return self.ok, output


class XlsxToPDF():
    xlsx_file_path = ''
    ok = None

    def __init__(self, xlsx_file_path):
        self.xlsx_file_path = xlsx_file_path
        self.ok = True

    def convert_xlsx_to_pdf(self):
        output = self.xlsx_file_path.replace('.xlsx', '.pdf')
        output = output.replace('.xls', '.pdf')
        excel = win32com.client.Dispatch('Excel.Application')
        try:
            logger.info(f'Convirtiendo archivo {self.xlsx_file_path} a {output}')
            wb = excel.Workbooks.Open(self.xlsx_file_path)
            wb.ExportAsFixedFormat(0, output)
            wb.Close()
            excel.Quit()
            logger.info(f'Archivo convertido con éxito {output}')
        except Exception as e:
            self.ok = False
            excel.Quit()
            logger.error(e)
            logger.error(f'error al convertir {self.xlsx_file_path}')
        return self.ok, output

class ImgToPDF():
    img_file_paths = []
    ok = None
    pdf_name = ''

    def __init__(self, img_file_paths, pdf_name=None):
        self.img_file_paths = img_file_paths
        self.ok = True
        self.pdf_name = pdf_name if pdf_name else 'file.pdf'

    def convert_imgs_to_pdf(self):
        try:
            logger.info(f'Convirtiendo imágnees {",".join(self.img_file_paths)}')
            images = [
                Image.open(f)
                for f in self.img_file_paths
            ]
            images[0].save(
                self.pdf_name, "PDF", resolution=100.0, save_all=True, append_images=images[1:]
            )
            logger.info(f'Imágenes convertidas con éxito {self.pdf_name}')
        except Exception as e:
            print(e)
            self.ok = False
            logger.error(f'Error al convertir {",".join(self.img_file_paths)}')
        return self.ok
