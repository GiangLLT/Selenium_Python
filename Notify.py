import tkinter as tk
import requests
from PIL import Image, ImageTk
from io import BytesIO

# root = tk.Tk()

def close_notification(root):
    root.destroy()

def show_notification_popup(title, message, width=100, height=100):
    root = tk.Tk()
    root.title("Custom Notification")

    # Remove title bar and border
    root.overrideredirect(True)
    root.attributes("-topmost", True)  # Ensure the window stays on top

    # Download the image from URL
    image_url = "https://static.vecteezy.com/system/resources/previews/016/017/077/large_2x/notification-icon-free-png.png"
    response = requests.get(image_url)

    if response.status_code == 200:
        try:
            # Open image using PIL
            image = Image.open(BytesIO(response.content))
            # Resize the image to desired dimensions
            resized_image = image.resize((width, height), Image.LANCZOS)
            photo = ImageTk.PhotoImage(resized_image)

            # Create a frame for the notification
            frame = tk.Frame(root, bg="white")
            frame.pack(fill="both", expand=True)

            # Create a frame for the text to align vertically
            text_frame = tk.Frame(frame, bg="white")
            text_frame.pack(side="left", fill="both", expand=True, padx=10, pady=10)

            # Create a title label
            title_label = tk.Label(text_frame, text=title, font=("Helvetica", 14, "bold"), bg="white")
            title_label.pack(anchor="w")

            # Create a message label
            message_label = tk.Label(text_frame, text=message, bg="white")
            message_label.pack(anchor="w")

            # Create an image label
            image_label = tk.Label(frame, image=photo, bg="white")
            image_label.pack(side="right", padx=10, pady=10)

            # Create a close button
            close_button = tk.Button(frame, text="X", command=lambda: close_notification(root), font=("Abadi Extra Light", 12), borderwidth=0, highlightthickness=0, bg="white")
            close_button.place(relx=1, x=-8, y=5, anchor="ne")

            # Calculate coordinates for bottom right corner
            window_width = 300
            window_height = 100
            screen_width = root.winfo_screenwidth()
            screen_height = root.winfo_screenheight()
            x_coordinate = screen_width - window_width - 10  # 10 pixels padding from the right edge
            y_coordinate = screen_height - window_height - 50  # 50 pixels padding from the bottom edge

            # Set the window geometry
            root.geometry(f"{window_width}x{window_height}+{x_coordinate}+{y_coordinate}")

            # Close the window after 5 seconds
            root.after(5000, root.destroy)

            root.mainloop()

        except Exception as e:
            print(f"Error opening image: {e}")

    else:
        print(f"Failed to download image. Status code: {response.status_code}")

# Example usage
title = "Notification Title"
message = "This is a custom notification message."
show_notification_popup(title, message, width=80, height=80)