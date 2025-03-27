import time 
from docx2pdf import convert
import win32com.client
from fpdf import FPDF
import os
from termcolor import colored
import random
from colorama import init
import comtypes.client

#ASCII title for viewing pleasure
title = r"""
   ______                           __     ___      ____      ______
  / ____/___  ____ _   _____  _____/ /_   |__ \    / __ \____/ / __/
 / /   / __ \/ __ \ | / / _ \/ ___/ __/   __/ /   / /_/ / __  / /_  
/ /___/ /_/ / / / / |/ /  __/ /  / /_    / __/   / ____/ /_/ / __/  
\____/\____/_/ /_/|___/\___/_/   \__/   /____/  /_/    \__,_/_/     
                                                                                 
                             
"""
print(colored(title, 'light_blue'))

init(strip=False)

#prompt user for folder path
print(colored("Enter path to folder with target files:", 'light_blue'))
path = input()

#function to convert all .doc/.docx files
def convert_docs_in_folder(folder_path):
    word = win32com.client.Dispatch('Word.Application')
    for doc_name in os.listdir(folder_path):
        if doc_name.endswith(".doc") or doc_name.endswith(".docx"):
            colors = ['blue', 'cyan', 'green', 'magenta', 'white', 'yellow', 'red', 'light_blue', 'light_cyan', 'light_green', 'light_magenta', 'light_red', 'light_yellow']
            doc_path = os.path.join(folder_path, doc_name)
            pdf_path = os.path.join(folder_path, doc_name.rsplit(".", 1)[0] + ".pdf")
            doc = word.Documents.Open(doc_path)
            doc.SaveAs(pdf_path, FileFormat=17)  # 17 is the constant for wdFormatPDF
            doc.Close()
            color_choice = random.choice(colors)
            filenamecolored = colored(f'Converted {str(doc_name)} to PDF', color_choice)
            print(f'{filenamecolored}')
        else:
            continue
    word.Quit()

#funtion to convert all excel files

def convert_xl_in_folder(folder_path):
    xl = win32com.client.Dispatch('Excel.Application')
    xl.Visible = False
    for file_name in os.listdir(folder_path):
        if file_name.endswith(".xlsx"):
            colors = ['blue', 'cyan', 'green', 'magenta', 'white', 'yellow', 'red', 'light_blue', 'light_cyan', 'light_green', 'light_magenta', 'light_red', 'light_yellow']
            doc_path = os.path.join(folder_path, file_name)
            pdf_path = os.path.join(folder_path, file_name.rsplit(".", 1)[0] + ".pdf")
            workbook = xl.Workbooks.Open(doc_path)
            workbook.ExportAsFixedFormat(0, pdf_path)  # 17 is the constant for wdFormatPDF
            workbook.Close()
            color_choice = random.choice(colors)
            filenamecolored = colored(str(file_name), color_choice)
            print(f"Converted {str(filenamecolored)} to PDF")

    xl.Quit()

# function to convert all .ppt and .pptx
def convert_pptx_in_folder(folder_path):
    powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
    powerpoint.WindowState = 2
    for file_name in os.listdir(folder_path):
        if file_name.endswith(".pptx"):
            colors = ['blue', 'cyan', 'green', 'magenta', 'white', 'yellow', 'red', 'light_blue', 'light_cyan', 'light_green', 'light_magenta', 'light_red', 'light_yellow']
            ppt_path = os.path.join(folder_path, file_name)
            pdf_path = os.path.join(folder_path, file_name.rsplit(".", 1)[0] + ".pdf")
                        # Open the PowerPoint file
            presentation = powerpoint.Presentations.Open(ppt_path, WithWindow=False)
            
            # Save as PDF
            presentation.SaveAs(pdf_path, 32)  # 32 corresponds to the PDF format

            # Close the presentation
            presentation.Close()
            color_choice = random.choice(colors)
            filenamecolored = colored(str(file_name), color_choice)
            print(f"Converted {str(filenamecolored)} to PDF")
        else:
            continue
    powerpoint.Quit()

folder_path = path

#calling funtions
convert_docs_in_folder(folder_path)
print("     ")
print("All .doc and .docx files in the folder have been converted to PDFs!")
print("     ")
convert_xl_in_folder(folder_path)
print("     ")
print("All .xlsx files have been converted!")
print("     ")
time.sleep(1)
convert_pptx_in_folder(folder_path)
print("     ")
print("All .ppt and .pptx files have been converted!")
print("     ")

time.sleep(1)
print("All done!")


