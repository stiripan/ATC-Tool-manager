import os
import time
import win32com.client
import re

import logging
logging.basicConfig(filename='script.log', level=logging.DEBUG)

class ExcelHandler:
    def __init__(self, filename, macro_name):
        self.excel = win32com.client.Dispatch("Excel.Application")
        self.excel.Visible = True
        self.workbook = self.excel.Workbooks.Open(os.path.abspath(filename))
        self.macro_name = macro_name

    def run_macro(self):
        try:
            self.excel.Application.Run(self.macro_name)
        except Exception as e:
            logging.error("Failed to run macro: {}".format(str(e)))

    def close_workbook(self):
        self.workbook.Close(SaveChanges=0)
        self.excel.Quit()

def determine_excel_file(machine):
    base_path = r"C:\Users\Blue Mill\Desktop\MACHINE TOOL LIST"
    excel_files = {
        'VF-3': 'VF-3.xlsm',
        'VF-4': 'VF-4.xlsm',
        'VF-15': 'VF-15.xlsm',
        'VF-25': 'VF-25.xlsm'
    }

    if machine not in excel_files:
        logging.error("Unknown machine: {}".format(machine))
        return None

    excel_file = os.path.join(base_path, excel_files[machine])
    if not os.path.exists(excel_file):
        logging.error("Excel file not found for machine: {}".format(machine))
        return None

    return excel_file

def extract_machine_from_nc_file(nc_file_path):
    with open(nc_file_path, 'r') as file:
        contents = file.read()
        pattern = r"vendor:\s+Haas\s+(VF-\d+)"
        match = re.search(pattern, contents)
        if match:
            machine_name = match.group(1)
            return machine_name
        else:
            return None

def process_nc_files(folder):
    processed_files = set()

    while True:
        time.sleep(1)
        files = os.listdir(folder)
        for file in files:
            if file.endswith(".nc"):
                nc_file_path = os.path.join(folder, file)
                machine_name = extract_machine_from_nc_file(nc_file_path)
                if machine_name:
                    excel_file = determine_excel_file(machine_name)
                    if excel_file:
                        try:
                            excel_handler = ExcelHandler(excel_file, "ReadNCFile")
                            excel_handler.run_macro()
                            processed_files.add(file)
                            logging.info("Processed NC file: {}".format(file))
                        except Exception as e:
                            logging.exception("Error occurred")

watch_folder = r'C:\Users\Blue Mill\Desktop\fusion NC'

print(f"Watching folder {watch_folder} for changes to NC files...")
process_nc_files(watch_folder)
