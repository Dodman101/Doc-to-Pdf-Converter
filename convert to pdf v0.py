import os
from tkinter import Tk, filedialog, messagebox
from docx import Document
from docx2pdf import convert


def convert_docx_to_pdf(input_path, output_path):
    # Convert the Word document to PDF using docx2pdf
    convert(input_path, output_path)


def select_word_documents():
    # Open a file dialog for selecting multiple Word documents
    root = Tk()
    root.withdraw()  # Hide the main window
    file_paths = filedialog.askopenfilenames(filetypes=[("Word Documents", "*.docx")])
    root.destroy()
    return file_paths


def select_output_folder():
    # Open a folder dialog for selecting the output folder
    root = Tk()
    root.withdraw()  # Hide the main window
    folder_path = filedialog.askdirectory()
    root.destroy()
    return folder_path


# Main program
try:
    # Select input Word documents
    messagebox.showinfo("Select Word Documents", "Please select the Word documents to convert.")
    input_files = select_word_documents()
    if not input_files:
        raise ValueError("No input files selected.")

    # Select output folder
    messagebox.showinfo("Select Output Folder", "Please select the folder to save the converted PDF files.")
    output_folder = select_output_folder()
    if not output_folder:
        raise ValueError("No output folder selected.")

    # Convert each selected Word document to PDF
    for input_file in input_files:
        file_name = os.path.basename(input_file)
        output_file = os.path.join(output_folder, os.path.splitext(file_name)[0] + ".pdf")
        convert_docx_to_pdf(input_file, output_file)
        messagebox.showinfo("Conversion Complete", f"Successfully converted '{file_name}' to PDF.")

except Exception as e:
    messagebox.showerror("Error", str(e))
