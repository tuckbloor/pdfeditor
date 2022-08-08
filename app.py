from pdf2docx import parse
from docx2pdf import convert
import tkinter as tk
from PIL import Image, ImageTk
from tkinter.filedialog import askopenfile
from pathlib import Path


root = tk.Tk()
bgc = "#E6E8FA"
root.configure(background=bgc)
root.geometry("900x1200")
root.title("Document Converter")
logo = Image.open('logo.jpg')

width, height = logo.size
logo = logo.resize((round(362/height*width), round(362)))
logo = ImageTk.PhotoImage(logo)
logo_label = tk.Label(image=logo)
logo_label.image = logo
logo_label.grid(column=1, row=0)


def open_file_pdf():

    pdf_file = askopenfile(parent=root, mode="rb", title="Select File", filetypes=[("Pdf file", "*.pdf")])

    if pdf_file:
        pdf_filename_to_convert = pdf_file.name.split('/')[-1].rsplit('.', maxsplit=1)[0]
        downloads_path = str(Path.home() / "Downloads")
        output = downloads_path + '/' + pdf_filename_to_convert + '.docx'
        parse(pdf_file, output)


def open_file_docx():

    docx_file2 = askopenfile(parent=root, mode="rb", title="Select File", filetypes=[("Docx file", "*.docx")])

    if docx_file2:
        docx_filename_to_convert = docx_file2.name.split('/')[-1].rsplit('.', maxsplit=1)[0]
        downloads_path = str(Path.home() / "Downloads")
        output = downloads_path + '/' + docx_filename_to_convert + '.pdf'
        convert(docx_file2.name, output)


text = tk.Label(root, text="Convert PDF TO Docx For Editing", bg=bgc)
text.grid(column=1, row=1)
browse_button_select_pdf = tk.Button(root, text="Browse", highlightbackground="#5D478B", fg="#000000", height=2, width=15, command=lambda: open_file_pdf())
browse_button_select_pdf.grid(column=1, row=2)

text = tk.Label(root, text="Convert Docx TO PDF For Editing", bg=bgc)
text.grid(column=1, row=3)
browse_button_select_docx = tk.Button(root, text="Browse", highlightbackground="#5D478B", fg="#000000", height=2, width=15, command=lambda: open_file_docx())
browse_button_select_docx.grid(column=1, row=4)

root.mainloop()