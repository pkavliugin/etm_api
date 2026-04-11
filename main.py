"""
Минимальный FastAPI backend для загрузки и обработки Excel файлов.
Endpoint POST /upload принимает xlsx файл и возвращает структурированные данные.
"""

from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import JSONResponse, StreamingResponse, HTMLResponse
import pandas as pd
from io import BytesIO
import logging
import requests
from typing import Dict, Any
import re

# Настройка логирования
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Константы конфигурации ETM API
# Документация: https://ipro.etm.ru/api/v1
ETM_API_BASE_URL = "https://ipro.etm.ru/api"  # Production API
# ETM_API_BASE_URL = "https://itest2.etm.ru/api"  # Test API

ETM_LOGIN = "your_login"  # Замените на ваши учетные данные
ETM_PASSWORD = "your_password"
ETM_SESSION_KEY = None  # Будет заполнен при авторизации

# Параметры запроса
REQUEST_TIMEOUT = 10  # секунды

# Ограничения по частоте запросов ETM:
# - Авторизация: 1 запрос в 2 минуты
# - Характеристики (/goods/{id}): 1 запрос в секунду
# - Цены (/goods/{id}/price): 1 запрос в секунду
# - Остатки (/goods/{id}/remains): 1 запрос в секунду

# Кэш для результатов поиска
cache = {}

# Инициализация приложения
app = FastAPI(
    title="Excel Upload Service",
    description="API для загрузки и обработки Excel файлов",
    version="1.0.0"
)


# ========================
# Утилиты и вспомогательные функции
# ========================

def get_etm_session() -> str:
    """
    Получение session ключа для работы с ETM API.
    
    Ключ действует 8 часов. При его истечении необходимо получить новый.
    
    Returns:
        str: Session ключ для использования в запросах
    
    Raises:
        Exception: Если авторизация не удалась
    """
    global ETM_SESSION_KEY
    
    if ETM_SESSION_KEY:
        logger.info("Использование существующего session ключа")
        return ETM_SESSION_KEY
    
    try:
        login_url = f"{ETM_API_BASE_URL}/v1/user/login"
        params = {"log": ETM_LOGIN, "pwd": ETM_PASSWORD}
        
        logger.info("Запрос авторизации в ETM API")
        response = requests.post(login_url, params=params, timeout=REQUEST_TIMEOUT)
        response.raise_for_status()
        
        data = response.json()
        if data.get("status", {}).get("code") == 200:
            ETM_SESSION_KEY = data.get("data", {}).get("session")
            logger.info("Успешная авторизация в ETM API")
            return ETM_SESSION_KEY
        else:
            raise Exception(f"Ошибка авторизации: {data.get('status', {}).get('message')}")
    
    except Exception as e:
        logger.error(f"Ошибка при авторизации в ETM: {str(e)}")
        raise


def normalize_name(name: str) -> str:
    """
    Нормализация названия товара для поиска.
    
    - Приводит строку к нижнему регистру
    - Убирает лишние пробелы
    - Заменяет кириллицу 'х' на 'x'
    - Удаляет лишние символы (скобки, запятые)
    
    Args:
        name: Исходное название
    
    Returns:
        str: Нормализованное название
    """
    # Приведение к нижнему регистру
    name = name.lower()
    
    # Замена кириллицы 'х' на 'x'
    name = name.replace('х', 'x')
    
    # Удаление лишних символов (скобки, запятые и подобные)
    name = re.sub(r'[()[\]{},;:"]', '', name)
    
    # Удаление дополнительных пробелов
    name = re.sub(r'\s+', ' ', name).strip()
    
    return name


def search_etm(name: str, code_type: str = "etm") -> Dict[str, Any]:
    """
    Функция для поиска товара в ETM через API (GET /v2/goods/{id}).
    Результаты кэшируются для оптимизации.
    
    Args:
        name: Название или код товара для поиска
        code_type: Тип кода (etm, cli, mnf). По умолчанию 'etm'
            - etm: коды ЭТМ (только цифры без префикса ETM)
            - cli: коды клиента (требует сопоставления с менеджером)
            - mnf: артикулы производителя
    
    Returns:
        Dict: Результат поиска от API или словарь с ошибкой
    """
    # Нормализация названия перед поиском
    normalized_name = normalize_name(name)
    
    # Проверка кэша
    if normalized_name in cache:
        logger.info(f"Результат найден в кэше для: {normalized_name}")
        return cache[normalized_name]
    
    try:
        # Получение session ключа для авторизации
        session_key = get_etm_session()
        
        # Формирование параметров GET запроса к API характеристик
        # Документация: GET /v2/goods/{id}
        goods_url = f"{ETM_API_BASE_URL}/v2/goods/{normalized_name}"
        params = {
            "type": code_type,
            "session-id": session_key
        }
        
        logger.info(f"Запрос характеристик товара: {name} ({code_type}={normalized_name})")
        response = requests.get(
            goods_url,
            params=params,
            timeout=REQUEST_TIMEOUT
        )
        
        # Проверка статуса ответа
        response.raise_for_status()
        
        # Структура ответа согласно документации ETM API
        data = response.json()
        if data.get("status", {}).get("code") != 200:
            logger.warning(f"ETM API вернул ошибку: {data.get('status', {}).get('message')}")
            return {"error": "API_ERROR", "message": data.get('status', {}).get('message')}
        
        logger.info(f"Успешный ответ от ETM API для: {normalized_name}")
        
        # Возврат структурированного результата
        result = {
            "status": "success",
            "data": data.get("data", {})
        }
        
        # Сохранение результата в кэш
        cache[normalized_name] = result
        
        return result
    
    except requests.exceptions.Timeout:
        logger.error(f"Timeout при запросе к ETM API для: {normalized_name}")
        return {
            "error": "Timeout",
            "message": "Запрос к ETM API превысил время ожидания"
        }
    
    except requests.exceptions.ConnectionError:
        logger.error(f"Ошибка соединения с ETM API")
        return {
            "error": "ConnectionError",
            "message": "Не удалось подключиться к ETM API"
        }
    
    except requests.exceptions.HTTPError as e:
        logger.error(f"HTTP ошибка при запросе к ETM API: {e.response.status_code}")
        return {
            "error": "HTTPError",
            "message": f"HTTP {e.response.status_code}: {e.response.reason}"
        }
    
    except Exception as e:
        logger.error(f"Неожиданная ошибка при поиске в ETM: {str(e)}")
        return {
            "error": "UnexpectedError",
            "message": f"Неожиданная ошибка: {str(e)}"
        }


# ========================
# HTML форма - веб-интерфейс
# ========================

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="ru">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>ETM API - Excel Processor</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            display: flex;
            align-items: center;
            justify-content: center;
            padding: 20px;
        }
        
        .container {
            background: white;
            border-radius: 12px;
            box-shadow: 0 20px 60px rgba(0, 0, 0, 0.3);
            padding: 40px;
            max-width: 500px;
            width: 100%;
        }
        
        h1 {
            color: #333;
            margin-bottom: 10px;
            font-size: 28px;
        }
        
        .subtitle {
            color: #666;
            font-size: 14px;
            margin-bottom: 30px;
        }
        
        .upload-area {
            border: 2px dashed #667eea;
            border-radius: 8px;
            padding: 30px;
            text-align: center;
            cursor: pointer;
            transition: all 0.3s ease;
            background-color: #f8f9fa;
        }
        
        .upload-area:hover {
            border-color: #764ba2;
            background-color: #f0f0f0;
        }
        
        .upload-area.dragover {
            border-color: #764ba2;
            background-color: #e8e4f3;
            transform: scale(1.02);
        }
        
        .upload-icon {
            font-size: 48px;
            margin-bottom: 15px;
        }
        
        .upload-text {
            color: #666;
            font-size: 16px;
            margin-bottom: 10px;
        }
        
        .upload-hint {
            color: #999;
            font-size: 13px;
        }
        
        input[type="file"] {
            display: none;
        }
        
        .button-group {
            display: flex;
            gap: 10px;
            margin-top: 30px;
        }
        
        button {
            flex: 1;
            padding: 12px 24px;
            border: none;
            border-radius: 6px;
            font-size: 16px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.3s ease;
        }
        
        .btn-process {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
        }
        
        .btn-process:hover {
            transform: translateY(-2px);
            box-shadow: 0 10px 20px rgba(102, 126, 234, 0.4);
        }
        
        .btn-process:disabled {
            opacity: 0.5;
            cursor: not-allowed;
            transform: none;
        }
        
        .btn-clear {
            background: #e0e0e0;
            color: #333;
        }
        
        .btn-clear:hover {
            background: #d0d0d0;
        }
        
        .file-info {
            margin-top: 20px;
            padding: 15px;
            background: #f0f9ff;
            border-left: 4px solid #667eea;
            border-radius: 4px;
            display: none;
        }
        
        .file-info.show {
            display: block;
        }
        
        .file-name {
            color: #333;
            font-weight: 600;
            word-break: break-all;
        }
        
        .loading {
            display: none;
            text-align: center;
            margin-top: 20px;
        }
        
        .loading.show {
            display: block;
        }
        
        .spinner {
            border: 4px solid #f3f3f3;
            border-top: 4px solid #667eea;
            border-radius: 50%;
            width: 40px;
            height: 40px;
            animation: spin 1s linear infinite;
            margin: 0 auto 10px;
        }
        
        @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
        }
        
        .error {
            color: #d32f2f;
            padding: 15px;
            background: #ffebee;
            border-left: 4px solid #d32f2f;
            border-radius: 4px;
            margin-top: 20px;
            display: none;
        }
        
        .error.show {
            display: block;
        }
        
        .info-box {
            background: #e3f2fd;
            border-left: 4px solid #2196f3;
            padding: 15px;
            border-radius: 4px;
            margin-bottom: 20px;
            font-size: 14px;
            color: #1565c0;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>📊 ETM API Excel Processor</h1>
        <p class="subtitle">Загрузите Excel файл для обработки с интеграцией ETM</p>
        
        <div class="info-box">
            Файл должен содержать колонки: <strong>"Наименование"</strong> и <strong>"Количество"</strong>
        </div>
        
        <form id="uploadForm" enctype="multipart/form-data">
            <div class="upload-area" id="uploadArea" onclick="document.getElementById('fileInput').click()">
                <div class="upload-icon">📁</div>
                <div class="upload-text">Нажмите для выбора файла</div>
                <div class="upload-hint">или перетащите файл сюда (xlsx)</div>
            </div>
            <input type="file" id="fileInput" name="file" accept=".xlsx" />
            
            <div class="file-info" id="fileInfo">
                <div class="file-name" id="fileName"></div>
            </div>
            
            <div class="button-group">
                <button type="submit" class="btn-process" id="processBtn" disabled>
                    ▶ Обработать
                </button>
                <button type="button" class="btn-clear" id="clearBtn">
                    ↺ Очистить
                </button>
            </div>
            
            <div class="loading" id="loading">
                <div class="spinner"></div>
                <p>Обработка файла...</p>
            </div>
            
            <div class="error" id="error"></div>
        </form>
    </div>
    
    <script>
        const uploadArea = document.getElementById('uploadArea');
        const fileInput = document.getElementById('fileInput');
        const uploadForm = document.getElementById('uploadForm');
        const processBtn = document.getElementById('processBtn');
        const clearBtn = document.getElementById('clearBtn');
        const fileInfo = document.getElementById('fileInfo');
        const fileName = document.getElementById('fileName');
        const loading = document.getElementById('loading');
        const errorDiv = document.getElementById('error');
        
        // Drag and drop
        ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
            uploadArea.addEventListener(eventName, preventDefaults, false);
        });
        
        function preventDefaults(e) {
            e.preventDefault();
            e.stopPropagation();
        }
        
        ['dragenter', 'dragover'].forEach(eventName => {
            uploadArea.addEventListener(eventName, () => {
                uploadArea.classList.add('dragover');
            });
        });
        
        ['dragleave', 'drop'].forEach(eventName => {
            uploadArea.addEventListener(eventName, () => {
                uploadArea.classList.remove('dragover');
            });
        });
        
        uploadArea.addEventListener('drop', handleDrop);
        
        function handleDrop(e) {
            const dt = e.dataTransfer;
            const files = dt.files;
            fileInput.files = files;
            handleFileSelect();
        }
        
        // File input change
        fileInput.addEventListener('change', handleFileSelect);
        
        function handleFileSelect() {
            const file = fileInput.files[0];
            if (file) {
                fileName.textContent = '✓ ' + file.name;
                fileInfo.classList.add('show');
                processBtn.disabled = false;
                errorDiv.classList.remove('show');
            } else {
                fileInfo.classList.remove('show');
                processBtn.disabled = true;
            }
        }
        
        // Clear button
        clearBtn.addEventListener('click', () => {
            fileInput.value = '';
            fileInfo.classList.remove('show');
            processBtn.disabled = true;
            errorDiv.classList.remove('show');
        });
        
        // Form submit
        uploadForm.addEventListener('submit', async (e) => {
            e.preventDefault();
            
            const file = fileInput.files[0];
            if (!file) return;
            
            loading.classList.add('show');
            errorDiv.classList.remove('show');
            processBtn.disabled = true;
            
            const formData = new FormData();
            formData.append('file', file);
            
            try {
                const response = await fetch('/upload', {
                    method: 'POST',
                    body: formData
                });
                
                if (response.ok) {
                    const blob = await response.blob();
                    const url = window.URL.createObjectURL(blob);
                    const a = document.createElement('a');
                    a.href = url;
                    a.download = 'result.xlsx';
                    document.body.appendChild(a);
                    a.click();
                    window.URL.revokeObjectURL(url);
                    document.body.removeChild(a);
                    
                    // Clear form after success
                    fileInput.value = '';
                    fileInfo.classList.remove('show');
                    
                    errorDiv.classList.add('show');
                    errorDiv.style.borderLeft = '4px solid #4caf50';
                    errorDiv.style.background = '#e8f5e9';
                    errorDiv.style.color = '#2e7d32';
                    errorDiv.textContent = '✓ Файл успешно обработан и загружен!';
                } else {
                    const error = await response.json();
                    throw new Error(error.detail || 'Ошибка обработки файла');
                }
            } catch (error) {
                errorDiv.classList.add('show');
                errorDiv.textContent = '✗ Ошибка: ' + error.message;
            } finally {
                loading.classList.remove('show');
                processBtn.disabled = false;
            }
        });
    </script>
</body>
</html>
"""


# ========================
# API Endpoints
# ========================

@app.get("/", response_class=HTMLResponse)
async def root():
    """Главная страница с HTML формой загрузки."""
    return HTML_TEMPLATE


@app.post("/upload")
async def upload_excel(file: UploadFile = File(...)):
    """
    Endpoint для загрузки Excel файла.
    
    Ожидает файл с расширением .xlsx.
    Ожидаемые колонки: "Наименование", "Количество"
    
    Returns:
        List[Dict]: Список словарей с полями "name" и "quantity"
    """
    
    # Проверка типа файла
    if not file.filename.endswith('.xlsx'):
        logger.warning(f"Попытка загрузить неправильный формат файла: {file.filename}")
        raise HTTPException(
            status_code=400,
            detail="Файл должен быть в формате .xlsx"
        )
    
    try:
        # Чтение содержимого файла в памяти
        contents = await file.read()
        file_buffer = BytesIO(contents)
        
        # Загрузка Excel файла через pandas
        df = pd.read_excel(file_buffer, engine='openpyxl')
        
        # Проверка наличия необходимых колонок
        required_columns = ["Наименование", "Количество"]
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            logger.error(f"Отсутствуют колонки: {missing_columns}")
            raise HTTPException(
                status_code=400,
                detail=f"Файл должен содержать колонки: {required_columns}. Отсутствуют: {missing_columns}"
            )
        
        # Преобразование данных в требуемый формат и обогащение данными из ETM
        result = []
        for index, row in df.iterrows():
            name = str(row["Наименование"]).strip()
            quantity = int(row["Количество"])
            
            # Поиск товара в ETM API
            etm_result = search_etm(name)
            
            item = {
                "name": name,
                "quantity": quantity,
                "etm_result": etm_result
            }
            result.append(item)
        
        logger.info(f"Файл {file.filename} успешно обработан. Обработано строк: {len(result)}")
        
        # Создание Excel файла в памяти
        output_df = pd.DataFrame(result)
        output_buffer = BytesIO()
        output_df.to_excel(output_buffer, index=False, engine='openpyxl')
        output_buffer.seek(0)
        
        # Возврат файла через StreamingResponse
        return StreamingResponse(
            iter([output_buffer.getvalue()]),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=result.xlsx"}
        )
    
    except pd.errors.EmptyDataError:
        logger.error("Файл пуст")
        raise HTTPException(
            status_code=400,
            detail="Файл Excel пуст"
        )
    
    except ValueError as e:
        logger.error(f"Ошибка при обработке данных: {str(e)}")
        raise HTTPException(
            status_code=400,
            detail=f"Ошибка при обработке данных: {str(e)}"
        )
    
    except Exception as e:
        logger.error(f"Неожиданная ошибка: {str(e)}")
        raise HTTPException(
            status_code=500,
            detail="Внутренняя ошибка сервера при обработке файла"
        )


@app.get("/search/{product_name}")
async def search_product(product_name: str):
    """
    Endpoint для поиска товара в ETM по названию.
    
    Args:
        product_name: Название товара для поиска
    
    Returns:
        Dict: Результат поиска от ETM API
    """
    logger.info(f"Получен запрос поиска для: {product_name}")
    
    if not product_name or not product_name.strip():
        raise HTTPException(
            status_code=400,
            detail="Название товара не может быть пустым"
        )
    
    result = search_etm(product_name)
    
    # Если произошла ошибка, вернуть 503 Service Unavailable
    if "error" in result:
        raise HTTPException(
            status_code=503,
            detail=result.get("message", "Ошибка при поиске в ETM")
        )
    
    return {"query": product_name, "results": result}


@app.get("/health")
async def health_check():
    """Endpoint для проверки здоровья сервиса."""
    return {"status": "healthy"}


if __name__ == "__main__":
    import uvicorn
    # Запуск сервера на localhost:8000
    uvicorn.run(app, host="0.0.0.0", port=8000)
