import tkinter as tk
from PIL import Image, ImageTk
import webbrowser

def close_notification(root):
    root.destroy()

def open_url(url):
    webbrowser.open(url)

def show_notification_popup(title, message, image_path, url):
    root = tk.Tk()
    root.overrideredirect(True)  # Bỏ thanh taskbar
    root.attributes("-topmost", True)  # Luôn hiển thị trên cùng

    frame = tk.Frame(root, borderwidth=0, relief="solid")
    frame.pack(fill="both", expand=True)

    # Thêm hình ảnh vào thông báo
    image = Image.open(image_path)
    image = image.resize((40, 40), Image.LANCZOS)
    photo = ImageTk.PhotoImage(image)
    image_label = tk.Label(frame, image=photo)
    image_label.pack(side="left", padx=10, pady=10)

    # Tạo nút đóng
    close_button = tk.Button(frame, text="X", command=lambda: close_notification(root), font=("Helvetica", 12), borderwidth=0, highlightthickness=0)
    close_button.pack(side="top", anchor="ne", padx=0, pady=5)

    # Thêm tiêu đề và thông điệp
    title_label = tk.Label(frame, text=title, font=("Helvetica", 14, "bold"))
    title_label.pack(side="top", anchor="w", padx=0, pady=(0, 0))

    message_label = tk.Label(frame, text=message, font=("Helvetica", 12))
    message_label.pack(side="top", anchor="w", padx=10, pady=(0, 10))

    # Mở URL khi nhấp vào thông báo
    frame.bind("<Button-1>", lambda e: open_url(url))

    # Vị trí hiển thị thông báo góc phải dưới cùng
    window_width = 300
    window_height = 100
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x_coordinate = screen_width - window_width - 10
    y_coordinate = screen_height - window_height - 50
    root.geometry(f"{window_width}x{window_height}+{x_coordinate}+{y_coordinate}")

    root.after(10000, root.destroy)  # Tự đóng cửa sổ sau 10 giây
    root.mainloop()

# Ví dụ sử dụng
show_notification_popup(
    title="Thông báo",
    message="Đây là nội dung của thông báo.",
    image_path="./Data/Img/noti.png",
    url="https://www.example.com"
)
