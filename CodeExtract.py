# Author: Jacob Durham
# GitHub username: JDurham95
# Date: 11/15/2024
# Description: Opens PDFs and scans for a code of the format, C or E + 9 digits where the first two digits are the
# sample year , the next two digits are the sample month and the last 5 digits are the sample number, ex E2410000301.
# The PDFs are renamed to that code.  This will not work for samples greater than 999.
from random import sample

import fitz
import os
import re
import tkinter as tk
from tkinter import messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD

def extract_code_from_pdf(filepath):
    """
    Extracts the sample code from the PDF.
    """

    #opens the PDF, if it exists
    try:
        doc= fitz.open(filepath)
    except FileNotFoundError:
        print(f"File not found: {filepath}")
        return None

    sample_code = None
    code_pattern = r'\b[EC]\d{9}(?:\d{3})?\b'

    #scans the doc for the code
    for page in doc:
        text = page.get_text()
        matches = re.findall(code_pattern, text)
        if matches:
            sample_code = matches[0]
            break

    doc.close()


    #changes the sample code digits 5 -7 to 0, if they are not zero
    if sample_code:
        code_corrected = False
        edited_code = sample_code
        for index in range(5,7):
            if sample_code[index] != "0":
                code_corrected = True
                edited_code = sample_code[:index] + "0" + sample_code[index+1:]

        print(f"Extracted Code: {edited_code}")
        return edited_code if code_corrected else sample_code
    return None


def rename_file_with_code(sample_code, filepath):
    """
    renames the file in place with the sample code
    """

    if sample_code:
        base_name = sample_code
        counter = 1
        new_file_name = f"{sample_code}.pdf"
        new_file_path = os.path.join(os.path.dirname(filepath), new_file_name)

        #if a file of that name already exits, modifies the file name to make it unique
        while os.path.exists(new_file_path):
            new_file_name = f"{base_name}_{counter}.pdf"
            new_file_path = os.path.join(os.path.dirname(filepath), new_file_name)
            counter += 1

        os.rename(filepath, new_file_path)
        print(f"Renamed file to {new_file_name}")
        return new_file_name
    print("Code not found in the document.")
    return False


def workflow1(filepath):
    """
    Helper function used to control the workflow
    """
    sample_code = extract_code_from_pdf(filepath)

    return rename_file_with_code(sample_code, filepath)

def handle_dropped_files(event):
    """
    Allows the program to hand multiple drag and drop files
    """

    filepaths= root.tk.splitlist(event.data)
    success_count = 0
    error_count = 0

    for filepath in filepaths:
        filepath = filepath.strip("{}")
        if filepath.endswith(".pdf"):
            new_file_name = workflow1(filepath)
            if new_file_name:
                success_count += 1
            else:
                error_count += 1

    messagebox.showinfo("Result",f"Processed {len(filepaths)} files:\nSuccess: {success_count}\nFailed: {error_count}")

#GUI
root = TkinterDnD.Tk()
root.title("Batch PDF Renamer")
root.geometry("500x200")

label = tk.Label(root, text = "Drag and Drop PDF files here to rename :)")
label.pack(pady=20)

#Enables Drag and Drop
root.drop_target_register(DND_FILES)
root.dnd_bind("<<Drop>>", handle_dropped_files)

root.mainloop()



