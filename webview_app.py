import os
from pathlib import Path
import threading
import webview
import time
import socket

# Thiết lập cờ báo hiệu đây là bản Desktop App (để tắt Captcha)
os.environ["DESKTOP_MODE"] = "1"

# Khởi tạo server Flask từ web_app
from web_app import app

def get_free_port():
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    sock.bind(('127.0.0.1', 0))
    port = sock.getsockname()[1]
    sock.close()
    return port

def start_server(port):
    app.run(host='127.0.0.1', port=port, debug=False, threaded=True, use_reloader=False)

if __name__ == '__main__':
    port = get_free_port()
    
    # Chạy Flask ở thread nền
    t = threading.Thread(target=start_server, args=(port,), daemon=True)
    t.start()
    
    # Đợi server lên
    time.sleep(1)
    
    # Mở cửa sổ Desktop bằng pywebview
    webview.create_window(
        'CheckForm KLTN v3.0 (3D Edition)', 
        f'http://127.0.0.1:{port}', 
        width=1280, 
        height=800,
        min_size=(900, 600),
        background_color='#050811'
    )
    
    webview.start()
