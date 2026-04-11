"""
Простой лаунчер для запуска FastAPI сервера прямо из VS Code.
Просто запустите этот файл (Ctrl+F5 или кнопка Run в VS Code).

Сервер запустится на http://127.0.0.1:8000
и автоматически откроется в браузере.
"""

import uvicorn
import webbrowser
import time
import os
from pathlib import Path


def run_server():
    """Запуск FastAPI сервера с автоматическим открытием браузера."""
    
    # Параметры сервера
    HOST = "127.0.0.1"
    PORT = 8000
    URL = f"http://{HOST}:{PORT}"
    
    # Информация о запуске
    print("\n" + "="*60)
    print("🚀 ETM API SERVER STARTING")
    print("="*60)
    print(f"📍 Server URL: {URL}")
    print(f"🌐 Web Interface: {URL}")
    print(f"📚 API Docs: {URL}/docs")
    print(f"🔧 ReDoc: {URL}/redoc")
    print("\n💡 Нажмите Ctrl+C для остановки сервера\n")
    print("="*60 + "\n")
    
    try:
        # Открываем браузер с небольшой задержкой
        # (даём серверу время на инициализацию)
        def open_browser():
            time.sleep(2)  # Ждём 2 секунды
            try:
                print(f"🌐 Открываю браузер на {URL}...")
                webbrowser.open(URL)
            except Exception as e:
                print(f"⚠️  Не удалось открыть браузер автоматически: {e}")
                print(f"Откройте вручную: {URL}")
        
        # Запускаем открытие браузера в отдельном потоке
        import threading
        browser_thread = threading.Thread(target=open_browser, daemon=True)
        browser_thread.start()
        
        # Запускаем uvicorn сервер
        uvicorn.run(
            "main:app",
            host=HOST,
            port=PORT,
            reload=True,  # автоматическая перезагрузка при изменении кода
            log_level="info",
            access_log=True
        )
        
    except OSError as e:
        if "Address already in use" in str(e):
            print(f"\n❌ ОШИБКА: Порт {PORT} уже используется!")
            print(f"   Пожалуйста, закройте другое приложение на этом порту")
            print(f"   или используйте другой порт.\n")
        else:
            print(f"\n❌ ОШИБКА запуска сервера: {e}\n")
    except KeyboardInterrupt:
        print("\n\n✅ Сервер остановлен (Ctrl+C)")
        print("="*60 + "\n")
    except Exception as e:
        print(f"\n❌ Неожиданная ошибка: {e}\n")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    # Проверяем, что мы в правильной директории
    if not Path("main.py").exists():
        print("❌ ОШИБКА: main.py не найден в текущей директории!")
        print("   Пожалуйста, запустите run.py из корневой директории проекта")
        exit(1)
    
    run_server()
