from flask import Flask, request, jsonify
from flask_socketio import SocketIO, emit
from flask_cors import CORS
import tkinter as tk
from Notification import show_notification_popup

app = Flask(__name__)
# socketio = SocketIO(app)
CORS(app)  # Thêm dòng này để cho phép CORS
socketio = SocketIO(app, cors_allowed_origins="*")  # Thêm đối số cors_allowed_origins

connected_users = {}

#============================================================#
# WEBHOOK SERVER
#============================================================#
@app.route('/webhook', methods=['POST'])
def webhook():
    data = request.json
    userid = data["userID"]
    message = data["messageSocket"]
    print(f"Received webhook data: {data}")
    if userid:
        send_notification_to_user(userid, message)
        return jsonify({'status': 'success'}), 200
    else:
        return jsonify({'status': 'error', 'message': 'Missing userID in request'}), 400
#============================================================#
# WEBHOOK SERVER
#============================================================#



#============================================================#
# WEBSOCKET SERVERpy webhook_server.py
#============================================================#
# Tạo cửa sổ popup thông báo
def show_notification_popup_Local(message):
    root = tk.Tk()
    root.title("Notification")
    
    label = tk.Label(root, text=message, padx=20, pady=20)
    label.pack()
    
    root.after(10000, root.destroy)  # Tự đóng cửa sổ sau 3 giây
    root.mainloop()

# Hàm gửi thông báo tới user thông qua WebSocket
def send_notification_to_user(user_id, message):
    if user_id in connected_users:
        socket_id = connected_users[user_id]
        socketio.emit('notification', {'message': message}, room=socket_id)
        print(f"Sent notification to user {user_id} via WebSocket")
        # Hiển thị cửa sổ thông báo cho người dùng
        # show_notification_popup(message)
        show_notification_popup("Automation Read File",message, "https://www.google.com.vn")
    else:
        print(f"User {user_id} is not connected or not in connected_users")


@socketio.on('connect')
def handle_connect():
    print(f"Client connected: {request.sid}")

@socketio.on('disconnect')
def handle_disconnect():
    for user_id, sid in connected_users.items():
        if sid == request.sid:
            del connected_users[user_id]
            break
    print(f"Client disconnected: {request.sid}")

@socketio.on('login')
def handle_login(data):
    user_id = data['user_id']
    socket_id = request.sid
    connected_users[user_id] = socket_id
    print(f"User {user_id} connected with socket id {socket_id}")
#============================================================#
# WEBSOCKET SERVER
#============================================================#

if __name__ == '__main__':
    # app.run(port=5000, debug=True)
    socketio.run(app, port=5000, debug=True)
