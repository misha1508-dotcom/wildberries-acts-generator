import os
import sys
import threading
import webview
from app import app

def start_flask():
    """Запускает Flask сервер в отдельном потоке"""
    app.run(host='127.0.0.1', port=5000, debug=False, use_reloader=False)

def main():
    # Запускаем Flask в отдельном потоке
    flask_thread = threading.Thread(target=start_flask, daemon=True)
    flask_thread.start()

    # Ждем немного, чтобы Flask успел запуститься
    import time
    time.sleep(1)

    # Создаем окно приложения
    window = webview.create_window(
        title='Генератор актов Wildberries',
        url='http://127.0.0.1:5000',
        width=800,
        height=900,
        resizable=True,
        fullscreen=False,
        min_size=(600, 700),
    )

    # Запускаем графическое окно
    webview.start()

if __name__ == '__main__':
    main()
