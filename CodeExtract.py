# Author: Jacob Durham
# GitHub username: JDurham95
# Date:
# Description:
from random import sample

import fitz
import ocrmypdf
import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD

class CodeNotFoundError(Exception):
    """
    error raised if a sample code is not found in a PDF file
    """
    pass

def ocr_pdf(filepath):
    """
    ocrs the odf if needed
    """
    ocrmypdf.ocr(filepath,filepath)


def extract_code_from_pdf(filepath):
    """
    extracts the sample code from the pdf
    """

    doc = fitz.open(filepath)
    sample_code = None

    code_pattern = r'\b[EC]\d{9}(?:-\d{3})?\b'

    for page in doc:
        text = page.get_text()
        matches = re.findall(code_pattern, text)
        if matches:
            sample_code = matches[0]
            break
        if not matches:
            return False

    doc.close()


    code_corrected = False
    edited_code = sample_code
    for index in range(5,7):
        if sample_code[index] != "0":
            code_corrected = True
            edited_code = sample_code[:index] + "0" + sample_code[index+1:]


    print(edited_code)
    if code_corrected:
        return edited_code
    else:
        return sample_code

def rename_file_with_code(sample_code, file_path):
    """
    renames the file with the sample code
    """

    if sample_code:
        new_file_name = f"{sample_code}.pdf"
        new_file_path = os.path.join(os.path.dirname(file_path), new_file_name)
        os.rename(file_path, new_file_path)
        print(f"Renamed file to {new_file_name}")
        return sample_code
    else:
        print("Code not found in the document.")
        return False

def workflow1(file_path):
    """
    controls the flow of methods
    """
    sample_code = extract_code_from_pdf(file_path)
    if not sample_code:
        ocr_pdf(file_path)
        sample_code = extract_code_from_pdf(file_path)
        if not sample_code:
            raise CodeNotFoundError

    rename_file_with_code(sample_code, file_path)



def handle_dropped_files(event):
    """
    Processes multiple files from drag-and-drop.
    """
    file_paths = root.tk.splitlist(event.data)
    success_count = 0
    error_count = 0

    for file_path in file_paths:
        if file_path.endswith(".pdf"):
            new_file_name = workflow1(file_path)
            if new_file_name:
                success_count += 1
            else:
                error_count += 1

    messagebox.showinfo(
        "Result",
        f"Processed {len(file_paths)} files:\nSuccess: {success_count}\nFailed: {error_count}"
    )


# Set up the TkinterDnD GUI
root = TkinterDnD.Tk()
root.title("Multi PDF Renamer")
root.geometry("400x200")

label = tk.Label(root, text="Drag and drop PDF files here to rename:")
label.pack(pady=20)

# Enable drag-and-drop
root.drop_target_register(DND_FILES)
root.dnd_bind('<<Drop>>', handle_dropped_files)

root.mainloop()


