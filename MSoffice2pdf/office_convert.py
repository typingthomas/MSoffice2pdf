import time 
from docx2pdf import convert
import win32com.client
from fpdf import FPDF
import os
from termcolor import colored
from colorama import init
import comtypes.client
import tkinter as tk
from tkinter import filedialog
import sys

# About info
def print_about():
    
    print(colored("MSoffice2pdf v1.2", 'light_cyan'))
    print(colored("https://github.com/typingthomas/MSoffice2pdf", 'light_cyan'))
if __name__ == "__main__":
    if "--about" in sys.argv:
        print_about()
        sys.exit() 

def is_odd(x):
    return x % 2 != 0

#Uses tkinter to Folder picker
def open_file_explorer():
    file_path = filedialog.askdirectory()  
    if file_path:
        clean_path = os.path.normpath(file_path)
        print(colored(f'Selected folder {clean_path}', 'light_yellow'))
        return clean_path
    else:
        print(colored('No valid path selected, Ending program', 'light_red'))
        time.sleep(2)
        exit()

#Defining functions to do the converting, this mainly uses pywin32

def convert_docs_in_folder(folder_path):
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False

    # Get a fixed list of files (only .doc and .docx)
    doc_files = [f for f in os.listdir(folder_path) if f.endswith('.doc') or f.endswith('.docx')]

    for filenum, doc_name in enumerate(doc_files):  # enumerate gives (index, filename)
        doc_path = os.path.join(folder_path, doc_name)
        pdf_path = os.path.join(folder_path, doc_name.rsplit(".", 1)[0] + ".pdf")
        try:
            doc = word.Documents.Open(doc_path)
            doc.SaveAs(pdf_path, FileFormat=17)  # Save as PDF
            doc.Close()

            if is_odd(filenum):
                filenamecolored = colored(f'Converted {doc_name} to PDF', colors[0])
            else:
                filenamecolored = colored(f'Converted {doc_name} to PDF', colors[1])

            print(filenamecolored)
            print(colored('-', 'white') * 50)

        except Exception as e:
            print(colored(f'could not convert {doc_name}: {e} File may be corrupted', 'light_red'))

    word.Quit()


def convert_xl_in_folder(folder_path):
    xl = win32com.client.Dispatch('Excel.Application')
    xl.Visible = False

    xl_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]
    
    # Get a fixed list of files (only .xlsx)
    for filenum, xl_name in enumerate(xl_files):    
        doc_path = os.path.join(folder_path, xl_name)
        pdf_path = os.path.join(folder_path, xl_name.rsplit(".", 1)[0] + ".pdf")
        try:
            workbook = xl.Workbooks.Open(doc_path)
            workbook.ExportAsFixedFormat(0, pdf_path)  
            workbook.Close()
            if is_odd(filenum):
                filenamecolored = colored(f'Converted {xl_name} to PDF', colors[0])
            else:
                filenamecolored = colored(f'Converted {xl_name} to PDF', colors[1])
            print(filenamecolored)
            print(colored('-', 'white') * 50)
        except:
            print(colored(f"Couldnt convert {str(xl_name)} to PDF, XL file may be corrupted", 'light_red'))
    xl.Quit()

def convert_pptx_in_folder(folder_path):
    powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
    powerpoint.WindowState = 2

    powerpoint_files = [f for f in os.listdir(folder_path) if f.endswith('.ppt') or f.endswith('.pptx')]

    for filenum, ppt_file in enumerate(powerpoint_files):
        ppt_path = os.path.join(folder_path, ppt_file)
        pdf_path = os.path.join(folder_path, ppt_file.rsplit(".", 1)[0] + ".pdf")
                    # Open the PowerPoint file
        try:
            presentation = powerpoint.Presentations.Open(ppt_path, WithWindow=False)
            
            # Save as PDF
            presentation.SaveAs(pdf_path, 32)  # 32 corresponds to the PDF format

            # Close the presentation
            presentation.Close()

            if is_odd(filenum):
                filenamecolored = colored(f'Converted {ppt_file} to PDF', colors[0])
            else:
                filenamecolored = colored(f'Converted {ppt_file} to PDF', colors[1])

            print(filenamecolored)
            print(colored('-', 'white') * 50)

        except Exception as e:
            print(colored(f'could not convert {ppt_file}: PowerPoint File may be corrupted', 'light_red'))
    powerpoint.Quit()

#ASCII title for viewing pleasure
title = r"""
   ______                           __     ___      ____      ______
  / ____/___  ____ _   _____  _____/ /_   |__ \    / __ \____/ / __/
 / /   / __ \/ __ \ | / / _ \/ ___/ __/   __/ /   / /_/ / __  / /_  
/ /___/ /_/ / / / / |/ /  __/ /  / /_    / __/   / ____/ /_/ / __/  
\____/\____/_/ /_/|___/\___/_/   \__/   /____/  /_/    \__,_/_/     
                                                                                 
                             
"""
print(colored(title, 'light_blue'))

colors = ['green', 'light_green']
path = open_file_explorer()
init(strip=False)

folder_path = path

#calling funtions
convert_docs_in_folder(folder_path)
print("     ")
print(colored("All .doc and .docx files in the folder have been converted to PDFs!",'light_yellow'))
print("     ")
convert_xl_in_folder(folder_path)
print("     ")
print(colored("All .xlsx files have been converted!", 'light_yellow'))
print("     ")
time.sleep(1)
convert_pptx_in_folder(folder_path)
print("     ")
print(colored("All .ppt and .pptx files have been converted!", 'light_yellow'))
print("     ")

time.sleep(1)
print("All done!")

