from docx2pdf import convert
import fitz  # PyMuPDF
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
from tkinter import messagebox
from PIL import Image, ImageTk

def convert_to_pdf(docx_path):
    pdf_path = docx_path.replace(".docx", ".pdf")
    convert(docx_path, pdf_path)
    return pdf_path

def open_pdf(file_path):
    try:
        pdf_document = fitz.open(file_path)
        return pdf_document
    except Exception as e:
        messagebox.showerror("Error", str(e))
        return None

def show_page(pdf_document, page_number):
    page = pdf_document.load_page(page_number)
    pix = page.get_pixmap()
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    return img

def show_pdf_in_tkinter(filepath):
    if filepath.endswith(".docx"):
        pdf_path = convert_to_pdf(filepath)
    else:
        pdf_path = filepath

    pdf_document = open_pdf(pdf_path)

    if pdf_document is None:
        return

    root = tk.Tk()
    root.title("PDF Viewer")

    canvas = tk.Canvas(root, width=800, height=1000)
    canvas.pack()

    page_number = 0
    img = show_page(pdf_document, page_number)
    photo = ImageTk.PhotoImage(img)

    img_label = tk.Label(root, image=photo)
    img_label.photo = photo
    img_label.pack()

    def next_page(event=None):
        nonlocal page_number
        page_number = (page_number + 1) % pdf_document.page_count
        img = show_page(pdf_document, page_number)
        photo = ImageTk.PhotoImage(img)
        img_label.configure(image=photo)
        img_label.photo = photo

    def prev_page(event=None):
        nonlocal page_number
        page_number = (page_number - 1) % pdf_document.page_count
        img = show_page(pdf_document, page_number)
        photo = ImageTk.PhotoImage(img)
        img_label.configure(image=photo)
        img_label.photo = photo

    root.bind("<Right>", next_page)
    root.bind("<Left>", prev_page)

    root.mainloop()

file_path = input("Vui lòng nhập đường dẫn tới tệp Excel: ")
if not file_path:
    file_path_tk = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx"), ("PDF files", "*.pdf")])
    if file_path_tk:
        show_pdf_in_tkinter(file_path_tk)
# file_path = "./80 - BCGE.docx"
# Open file dialog to select a file
# file_path = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx"), ("PDF files", "*.pdf")])
else:
    show_pdf_in_tkinter(file_path)
