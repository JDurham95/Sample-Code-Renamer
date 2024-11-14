# Author: Jacob Durham
# GitHub username: JDurham95
# Date:
# Description:

import fitz
import ocrmypdf
import os
import re
import tkinter as tk
from tkinter import messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD

# Persistent set to track unique sample codes
id_set = set()


def ocr_pdf(filepath):
    """
    OCRs the PDF if needed
    """
    try:
        ocrmypdf.ocr(filepath, filepath)
    except Exception as e:
        print(f"Error in OCR processing: {e}")
        return False
    return True


def extract_code_from_pdf(filepath):
    """
    Extracts the sample code from the PDF.
    """
    try:
        doc = fitz.open(filepath)
    except FileNotFoundError:
        print(f"File not found: {filepath}")
        return None

    sample_code = None
    code_pattern = r'\b[EC]\d{9}(?:-\d{3})?\b'

    for page in doc:
        text = page.get_text()
        matches = re.findall(code_pattern, text)
        if matches:
            sample_code = matches[0]
            break

    doc.close()

    if sample_code:
        code_corrected = False
        edited_code = sample_code
        for index in range(5, 7):
            if sample_code[index] != "0":
                code_corrected = True
                edited_code = sample_code[:index] + "0" + sample_code[index + 1:]

        print(f"Extracted Code: {edited_code}")
        return edited_code if code_corrected else sample_code
    return None


def rename_file_with_code(sample_code, file_path):
    """
    Renames the file with the sample code, ensuring unique filenames.
    """
    if sample_code:
        base_name = sample_code
        count = 1
        new_file_name = f"{sample_code}.pdf"
        new_file_path = os.path.join(os.path.dirname(file_path), new_file_name)

        # Ensure the filename is unique by adding a suffix if needed
        while os.path.exists(new_file_path):
            new_file_name = f"{base_name}_{count}.pdf"
            new_file_path = os.path.join(os.path.dirname(file_path), new_file_name)
            count += 1

        os.rename(file_path, new_file_path)
        print(f"Renamed file to {new_file_name}")
        return new_file_name
    print("Code not found in the document.")
    return False


def workflow1(file_path):
    """
    Controls the flow of methods.
    """

    sample_code = extract_code_from_pdf(file_path)
    if not sample_code:
        ocr_success = ocr_pdf(file_path)
        if ocr_success:
            sample_code = extract_code_from_pdf(file_path)

    # If the sample_code is already in the set, mark it as an error and make it unique
    if sample_code in id_set:
        sample_code = f"Error_{sample_code}"

    id_set.add(sample_code)
    print(sample_code)

    return rename_file_with_code(sample_code, file_path)


def handle_dropped_files(event):
    """
    Processes multiple files from drag-and-drop.
    """
    file_paths = root.tk.splitlist(event.data)
    success_count = 0
    error_count = 0

    for file_path in file_paths:
        file_path = file_path.strip("{}")  # Remove braces around paths with spaces
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




