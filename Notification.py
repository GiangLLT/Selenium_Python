import tkinter as tk
import requests
from PIL import Image, ImageTk
from io import BytesIO
import webbrowser


def close_notification(root):
    root.destroy()

def open_url(url):
    webbrowser.open(url)

def show_notification_popup(title, message, url, width=60, height=60):
    root = tk.Tk()
    root.title("Custom Notification")

    # Remove title bar and border
    root.overrideredirect(True)
    
    # # Load an image (example)   
    try:
            # Open image using PIL
            # image = Image.open(BytesIO(response.content))
            # Resize the image to desired dimensions
            # resized_image = image.resize((width, height), Image.ANTIALIAS)
            image_path = "./Data/Img/noti.png"
            image = Image.open(image_path)
            resized_image = image.resize((width, height), Image.LANCZOS)
            photo = ImageTk.PhotoImage(resized_image)

            # Create a frame for the notification
            frame = tk.Frame(root, bg="white")
            frame.pack(fill="both", expand=True)
            # Create an image label
            image_label = tk.Label(frame, image=photo, bg="white")
            image_label.pack(side="left", padx=10, pady=10)

            # Create a frame for the text to align vertically
            text_frame = tk.Frame(frame, bg="white")
            text_frame.pack(side="left", fill="both", expand=True, padx=10, pady=10)
            

            # Create a title label
            title_label = tk.Label(text_frame, text=title, font=("Helvetica", 12, "bold"), bg="white")
            title_label.pack(anchor="w")

            # Create a message label
            message_label = tk.Label(text_frame, text=message, bg="white")
            message_label.pack(anchor="w")


            # Create a close button
            close_button = tk.Button(frame, text="X", command=lambda: close_notification(root), font=("Abadi Extra Light", 12), borderwidth=0, highlightthickness=0, bg="white")
            close_button.place(relx=1, x=-8, y=5, anchor="ne")

            # Mở URL khi nhấp vào thông báo
            frame.bind("<Button-1>", lambda e: open_url(url))

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

# Example usage
# url="https://www.google.com.vn"
# title = "Notification"
# message = "This is a custom notification message."
# show_notification_popup(title, message, url, width=60, height=60)



    # photo = ImageTk.PhotoImage(image)

    # # Create a label with image and message
    # label = tk.Label(root, text=message, padx=20, pady=20, compound=tk.LEFT, image=photo)
    # label.pack()
    
    # # Center the window on screen
    # window_width = 300
    # window_height = 100
    # screen_width = root.winfo_screenwidth()
    # screen_height = root.winfo_screenheight()
    # x_coordinate = (screen_width - window_width) // 2
    # y_coordinate = (screen_height - window_height) // 2
    # root.geometry(f"{window_width}x{window_height}+{x_coordinate}+{y_coordinate}")
    
    # # Close the window after 5 seconds
    # root.after(5000, root.destroy)
    
    # root.mainloop()



    # # Download the image from URL
    # image_url = "https://static.vecteezy.com/system/resources/previews/016/017/077/large_2x/notification-icon-free-png.png"
    # response = requests.get(image_url)
    
    # if response.status_code == 200: 
    #     try:
    #         # Open image using PIL
    #         image = Image.open(BytesIO(response.content))
    #         # Resize the image to desired dimensions
    #         # resized_image = image.resize((width, height), Image.ANTIALIAS)
    #         resized_image = image.resize((width, height), Image.LANCZOS)
    #         photo = ImageTk.PhotoImage(resized_image)
            
    #         # Create a label with image and message
    #         label = tk.Label(root, text=message, padx=20, pady=20, compound=tk.LEFT, image=photo)
    #         label.pack()

    #         # Create a close button
    #         # close_button = tk.Button(root, text="X", command=close_notification, font=("Arial", 12, "bold"),foreground="red", borderwidth=0, highlightthickness=0 )
    #         close_button = tk.Button(root, text="X", command=close_notification, font=("Abadi Extra Light", 12), borderwidth=0, highlightthickness=0 )
    #         close_button.place(relx=1, x=-8, y=5, anchor="ne")
            
    #         # Center the window on screen
    #         window_width = 300
    #         window_height = 80
    #         screen_width = root.winfo_screenwidth()
    #         screen_height = root.winfo_screenheight()

    #         # x_coordinate = (screen_width - window_width) // 2
    #         # y_coordinate = (screen_height - window_height) // 2
    #         # root.geometry(f"{window_width}x{window_height}+{x_coordinate}+{y_coordinate}")

    #         # Calculate coordinates for bottom right corner
    #         x_coordinate = screen_width - window_width - 10  # 10 pixels padding from the right edge
    #         y_coordinate = screen_height - window_height - 80  # 10 pixels padding from the bottom edge
    #         # Set the window geometry
    #         root.geometry(f"{window_width}x{window_height}+{x_coordinate}+{y_coordinate}")
            
    #         # Close the window after 5 seconds
    #         root.after(5000, root.destroy)
            
    #         root.mainloop()
        
    #     except Exception as e:
    #         print(f"Error opening image: {e}")
    
    # else:
    #     print(f"Failed to download image. Status code: {response.status_code}")

