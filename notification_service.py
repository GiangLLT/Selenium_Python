from flask_socketio import SocketIO, Client
from win10toast import ToastNotifier
import requests
import time

socketio = Client()
toaster = ToastNotifier()

# Hàm hiển thị thông báo hệ thống Windows
def show_windows_notification(title, message):
    toaster.show_toast(title, message, threaded=True, icon_path=None, duration=10)

# Kết nối đến Flask-SocketIO server
def connect_to_server():
    socketio.connect('http://localhost:5000')
    user_id = input("Enter your user ID: ")
    
    # Đăng ký người dùng với server
    socketio.emit('register', {'user_id': user_id, 'socket_id': socketio.sid})

def receive_notifications():
    @socketio.on('notification')
    def handle_notification(data):
        title = data['title']
        message = data['message']
        show_windows_notification(title, message)

    socketio.wait()

if __name__ == "__main__":
    connect_to_server()
    receive_notifications()
