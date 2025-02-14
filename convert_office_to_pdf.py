##
## purposes: convert Excel, Word and PowerPoint documents in a specified folder into PDF, including files in child folder
## usage: py convert_office_to_pdf.py INPUT_FOLDER OUTPUT_FOLDER PROCESSED_FOLDER LOG_FOLDER
##
import os
import sys
import shutil
import win32com.client
import logging
from datetime import datetime
import psutil
import argparse

def print_info(info_message):
    print(f"⚠️ {info_message}")    
    logging.info(info_message)

def print_error(error_message):
    print(f"❌ {error_message}")
    logging.error(error_message)

def print_success(success_message):
    print(f"✅ {success_message}")
    logging.info(success_message)

def force_quit_word():
    # Find and kill all running Word processes
    for process in psutil.process_iter(['pid', 'name']):
        if process.info['name'].lower() == 'winword.exe':  # winword.exe is the process for Microsoft Word
            process.terminate()  # Terminate the Word process
            print_info("Terminate Word")

def move_to_processed(file_path, rel_path):
    """Move processed files to the 'Processed' folder while maintaining the directory structure."""
    processed_path = os.path.join(processed_base_folder, rel_path)
    os.makedirs(os.path.dirname(processed_path), exist_ok=True)
    try:
        shutil.move(file_path, processed_path)
        print_info(f"Moved to Processed: {rel_path}")        
    except Exception as e:
        print_error(f"Error moving file {file_path}: {str(e)}")

def copy_to_output(file_path, rel_path):
    """Copy not supported files to the 'Output' folder while maintaining the directory structure."""
    processed_path = os.path.join(processed_base_folder, rel_path)
    os.makedirs(os.path.dirname(processed_path), exist_ok=True)
    try:
        shutil.copy(file_path, processed_path)        
        print_info(f"Copy to Output: {rel_path}")        
    except Exception as e:
        print_error(f"Error copying file {file_path}: {str(e)}")

def convert_word_to_pdf(input_path, output_path, rel_path):
    """Convert a Word document (doc, docx) to PDF."""
    # Try 3 times
    for attempt in range (3):
        try:
            word = win32com.client.Dispatch("Word.Application")
            #word.Visible = False
            doc = word.Documents.Open(input_path)
            doc.SaveAs(output_path, FileFormat=17)  # 17 = wdFormatPDF
            doc.Close(False)
            word.Quit()
            print_success(f"Word converted: {rel_path} -> {os.path.basename(output_path)}")
            move_to_processed(input_path, rel_path)
            break
        except Exception as e:
            print_error(f"Error converting Word file {rel_path}: {str(e)}")
            print_info (f"Retrying attempt #{attempt+1}/3")
            if attempt > 3:
                print_error(f"Failed to convert Word file {rel_path}: {str(e)}")
                raise

def convert_excel_to_pdf(input_path, output_path, rel_path):
    """Convert an Excel spreadsheet (xls, xlsx) to PDF."""
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        #excel.Visible = False
        wb = excel.Workbooks.Open(input_path)
        wb.ExportAsFixedFormat(0, output_path)  # 0 = xlTypePDF
        wb.Close(False)
        excel.Quit()
        print_success(f"Excel converted: {rel_path} -> {os.path.basename(output_path)}")
        move_to_processed(input_path, rel_path)
    except Exception as e:
        print_error(f"Error converting Excel file {rel_path}: {str(e)}")

def convert_powerpoint_to_pdf(input_path, output_path, rel_path):
    """Convert a PowerPoint presentation (ppt, pptx) to PDF."""
    try:
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        presentation = powerpoint.Presentations.Open(input_path, WithWindow=False)
        presentation.SaveAs(output_path, 32)  # 32 = ppSaveAsPDF
        presentation.Close()
        powerpoint.Quit()
        print_success(f"PowerPoint converted: {rel_path} -> {os.path.basename(output_path)}")
        move_to_processed(input_path, rel_path)
    except Exception as e:
        print_error(f"Error converting PowerPoint file {rel_path}: {str(e)}")

def convert_office_files_recursively(input_folder):
    """Find and convert all Office documents in a folder and its subfolders to PDF."""
    for root, _, files in os.walk(input_folder):
        for filename in files:
            file_path = os.path.join(root, filename)
            rel_path = os.path.relpath(file_path, input_folder)  # Preserve directory structure
            output_pdf_path = os.path.join(output_base_folder, os.path.splitext(rel_path)[0] + ".pdf")
            os.makedirs(os.path.dirname(output_pdf_path), exist_ok=True)

            # Skip if PDF already exists
            if os.path.exists(output_pdf_path):
                print_info(f"Skipping (already converted): {rel_path}")
                move_to_processed(file_path, rel_path)
                continue

            file_ext = os.path.splitext(filename)[1].lower()
            if file_ext in [".doc", ".docx"]:
                convert_word_to_pdf(file_path, output_pdf_path, rel_path)
            elif file_ext in [".xls", ".xlsx"]:
                convert_excel_to_pdf(file_path, output_pdf_path, rel_path)
            elif file_ext in [".ppt", ".pptx"]:
                convert_powerpoint_to_pdf(file_path, output_pdf_path, rel_path)
            else:
                print_info(f"Skipping (unsupported file): {rel_path}")
                copy_to_output(file_path, output_pdf_path)
                move_to_processed(file_path, rel_path)

# Only run if executed directly
if __name__ != "__main__":
    sys.exit()

# Get folders from arguments
parser = argparse.ArgumentParser(description="Convert Ms Office files in a folder to PDF")
parser.add_argument("--input-folder", required=True, help="Name of input folder")
parser.add_argument("--output-folder", required=True, help="Name of output folder")
parser.add_argument("--processed-folder", required=True, help="Name of folder to archive processed files")
parser.add_argument("--log-folder", required=True, help="Name of folder to store conversion log file")

args = parser.parse_args()
input_folder = args.input_folder
output_base_folder = args.output_folder
processed_base_folder = args.processed_folder
log_folder = args.log_folder

# Ensure base output and processed directories exist
os.makedirs(output_base_folder, exist_ok=True)
os.makedirs(processed_base_folder, exist_ok=True)

# Ensure log folder exists
os.makedirs(log_folder, exist_ok=True)

# Configure logging
logging.basicConfig(
    filename=os.path.join(log_folder, datetime.now().strftime("process_%Y-%m-%d_%H-%M-%S.log")),  # Log file name
    level=logging.INFO,       # Logging level (INFO, DEBUG, WARNING, ERROR, CRITICAL)
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# 🔹 Run the conversion recursively

print_info("Start processing")
convert_office_files_recursively(input_folder)
print_success("Finish processing")

