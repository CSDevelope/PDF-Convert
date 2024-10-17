import os
import time
import threading
import tempfile  
import json
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from docx import Document
from fpdf import FPDF
from PIL import Image
import win32com.client
from pywintypes import com_error
import pythoncom

# Load configuration from config.json
with open('config.json', 'r') as config_file:
    config = json.load(config_file)

PDF_FOLDER = os.path.abspath(config['pdf_folder'])
WATCH_FOLDER = os.path.abspath(config['watch_folder'])

os.makedirs(PDF_FOLDER, exist_ok=True)
os.makedirs(WATCH_FOLDER, exist_ok=True)
os.makedirs('./fonts', exist_ok=True)

class FileHandler(FileSystemEventHandler):
    def on_created(self, event):
        if not event.is_directory:
            filepath = event.src_path
            file_extension = os.path.splitext(filepath)[1].lower()
            pdf_filename = os.path.basename(filepath).rsplit('.', 1)[0] + '.pdf'
            pdf_filepath = os.path.join(PDF_FOLDER, pdf_filename)
            
            try:
                if file_extension == '.docx':
                    convert_docx_to_pdf(filepath, pdf_filepath)
                elif file_extension in ['.png', '.jpg', '.jpeg']:
                    convert_image_to_pdf(filepath, pdf_filepath)
                elif file_extension in ['.xls', '.xlsx']:
                    convert_excel_to_pdf(filepath, pdf_filepath)
                else:
                    print(f"Unsupported file type: {file_extension}. Deleting file.")
                    time.sleep(1)  # Delay to ensure file is fully written before deletion
                    os.remove(filepath)
            except Exception as e:
                print(f"Failed to convert {filepath}: {e}")

# Conversion functions
def convert_docx_to_pdf(docx_path, pdf_path):
    doc = Document(docx_path)
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.add_font('DejaVu', '', 'fonts/DejaVuSans.ttf')
    pdf.set_font('DejaVu', size=10)

    for para in doc.paragraphs:
        if para.text.strip():
            pdf.multi_cell(190, 10, para.text)
            pdf.ln(5)

    pdf.output(pdf_path)

def convert_image_to_pdf(image_path, pdf_path):
    image = Image.open(image_path)
    pdf = FPDF()
    pdf.add_page()
    pdf.add_font('DejaVu', '', 'fonts/DejaVuSans.ttf')
    pdf.set_font('DejaVu', size=10)
    pdf.image(image_path, x=10, y=10, w=190)
    pdf.output(pdf_path)

def convert_excel_to_pdf(excel_path, pdf_path):
    excel_path = os.path.abspath(excel_path)  # Get the absolute path for the Excel file
    pdf_path = os.path.abspath(pdf_path)      # Get the absolute path for the PDF output
    
    pythoncom.CoInitialize()  
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    try:
        wb = excel.Workbooks.Open(excel_path)
        temp_pdf_path = os.path.join(tempfile.gettempdir(), os.path.basename(pdf_path))
        wb.WorkSheets.Select()
        wb.ActiveSheet.ExportAsFixedFormat(0, temp_pdf_path)
        
        if os.path.exists(pdf_path):
            os.remove(pdf_path)
        os.rename(temp_pdf_path, pdf_path)
    except com_error as e:
        print('Excel to PDF conversion failed:', e)
    finally:
        if 'wb' in locals():  
            wb.Close(SaveChanges=False)
        excel.Quit()
        pythoncom.CoUninitialize()  

# Start monitoring the folder
def start_monitoring():
    event_handler = FileHandler()
    observer = Observer()
    observer.schedule(event_handler, WATCH_FOLDER, recursive=False)
    observer.start()
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()

if __name__ == "__main__":
    monitoring_thread = threading.Thread(target=start_monitoring)
    monitoring_thread.start()
