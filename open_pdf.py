from docx2pdf import convert
import fitz  # PyMuPDF
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import os

def convert_to_pdf(docx_path):
    pdf_path = docx_path.replace(".docx", ".pdf")
    convert(docx_path, pdf_path)
    return pdf_path

def open_file(canvas, interior):
    file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")])
    if file_path:
        if file_path.endswith(".docx"):
            pdf_path = convert_to_pdf(file_path)
        else:
            pdf_path = file_path
        show_pdf(pdf_path, canvas, interior)

def show_pdf(file_path, canvas, interior):
    try:
        doc = fitz.open(file_path)
        num_pages = doc.page_count

        # Xóa trước khi hiển thị trang mới
        for widget in interior.winfo_children():
            widget.destroy()

        # Hiển thị từng trang dưới dạng ảnh và cuộn xuống
        y_position = 0
        for i in range(num_pages):
            page = doc.load_page(i)
            pix = page.get_pixmap(alpha=False)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            photo = ImageTk.PhotoImage(img)

            label = tk.Label(interior, image=photo)
            label.photo = photo  # Giữ tham chiếu để tránh mất ảnh
            label.pack()

            canvas.create_window(0, y_position, anchor="nw", window=label)
            y_position += pix.height

        # Cập nhật kích thước nội dung của canvas
        canvas.config(scrollregion=canvas.bbox("all"))

    except Exception as e:
        print("Error", str(e))
        messagebox.showerror("Error", str(e))

def show_pdf_in_tkinter(filepath=None):
    # Kiểm tra phần mềm PDF viewer trên máy
    if os.system("which acroread") == 0:  # Sử dụng Adobe Reader nếu có
        os.system(f"acroread {filepath}")
        return
    elif os.system("which evince") == 0:  # Sử dụng Evince nếu có
        os.system(f"evince {filepath}")
        return

    # Tạo cửa sổ GUI
    root = tk.Tk()
    root.title("PDF Viewer")

    # Tạo khung chứa PDF
    frame = tk.Frame(root)
    frame.pack(fill='both', expand=True)

    # Tạo canvas để cuộn
    canvas = tk.Canvas(frame)
    canvas.pack(side="left", fill="both", expand=True)

    # Thêm thanh cuộn dọc
    scrollbar = tk.Scrollbar(frame, orient="vertical", command=canvas.yview)
    scrollbar.pack(side="right", fill="y")

    # Cấu hình canvas với scrollbar
    canvas.config(yscrollcommand=scrollbar.set)

    # Tạo khung nội dung bên trong canvas
    interior = tk.Frame(canvas)
    interior_id = canvas.create_window(0, 0, window=interior, anchor="nw")

    # Cập nhật kích thước nội dung của canvas khi thay đổi kích thước
    def on_configure(event):
        canvas.config(scrollregion=canvas.bbox("all"))

    interior.bind("<Configure>", on_configure)

    # Tạo menu
    menu = tk.Menu(root)
    root.config(menu=menu)

    # Thêm tùy chọn "File" vào menu
    file_menu = tk.Menu(menu, tearoff=0)
    menu.add_cascade(label="File", menu=file_menu)
    file_menu.add_command(label="Open", command=lambda: open_file(canvas, interior))
    file_menu.add_separator()
    file_menu.add_command(label="Exit", command=root.quit)

    # Mở tệp PDF nếu filepath được cung cấp
    if filepath:
        show_pdf(filepath, canvas, interior)

    # Chạy ứng dụng GUI
    root.mainloop()

# Example usage
filepath = "./80 - BCGE.pdf"
show_pdf_in_tkinter(filepath)
