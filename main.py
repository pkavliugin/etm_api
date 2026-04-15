"""
ETM Excel Processor - Профессиональная обработка спецификаций

Цель приложения:
- Обработка Excel спецификаций товаров
- Получение актуальных данных из API ЭТМ:
  * Цен на товары
  * Информации о наличии и сроках поставки
  * Артикулов товаров
- Формирование структурированного Excel файла

Результат используется для:
- Расчета коммерческих предложений
- Планирования закупок материалов
- Анализа ценовой политики поставщиков
"""

from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import JSONResponse, StreamingResponse, HTMLResponse
import pandas as pd
from io import BytesIO
import logging
import requests
from typing import Dict, Any, List, Optional
import re
import os
from cachetools import TTLCache
from rapidfuzz import fuzz
import json

# Настройка логирования
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Константы конфигурации ETM API
# Документация: https://ipro.etm.ru/api/v1
ETM_API_BASE_URL = "https://ipro.etm.ru/api"  # Production API
# ETM_API_BASE_URL = "https://itest2.etm.ru/api"  # Test API

# ⚠️ ВАЖНО: Установите свои учетные данные ETM для работы с API
# Получите их на https://ipro.etm.ru/
ETM_LOGIN = os.getenv("ETM_LOGIN", "test_user")
ETM_PASSWORD = os.getenv("ETM_PASSWORD", "test_password")
ETM_SESSION_KEY = None  # Будет заполнен при авторизации

# Параметры запроса
REQUEST_TIMEOUT = 10  # секунды

# Ограничения по частоте запросов ETM:
# - Авторизация: 1 запрос в 2 минуты
# - Характеристики (/goods/{id}): 1 запрос в секунду
# - Цены (/goods/{id}/price): 1 запрос в секунду
# - Остатки (/goods/{id}/remains): 1 запрос в секунду

# Кэш для результатов поиска с TTL 1 час (3600 секунд)
cache = TTLCache(maxsize=1000, ttl=3600)

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


def extract_etm_data(etm_result: Dict[str, Any], original_name: str) -> Dict[str, Any]:
    """
    Извлечение данных из ответа ETM API с выбором лучшего товара.
    
    Если API возвращает несколько товаров:
    1. Выбирает лучший матч по названию (fuzzy matching)
    2. Среди похожих выбирает товар с наличием, затем с минимальной ценой
    3. Сохраняет топ-3 варианта для анализа
    
    Args:
        etm_result: Результат от search_etm()
        original_name: Исходное название для fuzzy matching
    
    Returns:
        Dict с ключами: 'found_name', 'article', 'unit', 'unit_price', 'availability', 'status', 'alternatives'
    """
    result = {
        'found_name': 'не найдено',
        'article': 'не найден',
        'unit': 'не указано',
        'unit_price': None,
        'availability': 'неизвестно',
        'status': 'not_found',
        'alternatives': []
    }
    
    try:
        # Если была ошибка при запросе
        if etm_result.get('error'):
            result['status'] = 'error'
            return result
        
        # Получаем список товаров
        goods_list = etm_result.get('data', [])
        if not goods_list:
            return result
        
        # Если вернулся один товар (не список)
        if isinstance(goods_list, dict):
            goods_list = [goods_list]
        
        # Нормализуем исходное название для сравнения
        normalized_original = normalize_name(original_name)
        
        # Оцениваем каждый товар
        scored_goods = []
        for good in goods_list:
            name = good.get('name') or good.get('title') or ''
            normalized_name = normalize_name(name)
            
            # Fuzzy matching score
            score = fuzz.ratio(normalized_original, normalized_name)
            
            # Извлекаем данные
            article = good.get('article') or good.get('code') or good.get('id') or 'не найден'
            unit = good.get('unit') or good.get('units') or good.get('measure') or 'не указано'
            price = good.get('price') or good.get('unit_price') or good.get('price_per_unit')
            availability = good.get('availability') or good.get('stock') or good.get('remains') or 'неизвестно'
            
            # Преобразуем цену в float
            try:
                price = float(price) if price is not None else None
            except (ValueError, TypeError):
                price = None
            
            scored_goods.append({
                'good': good,
                'name': name,
                'article': article,
                'unit': unit,
                'price': price,
                'availability': availability,
                'score': score
            })
        
        # Сортируем по score (лучший матч первым)
        scored_goods.sort(key=lambda x: x['score'], reverse=True)
        
        # Выбираем лучший товар
        best_good = None
        if scored_goods:
            # Фильтруем товары с достаточным совпадением (score > 70)
            relevant_goods = [g for g in scored_goods if g['score'] > 70]
            if not relevant_goods:
                # Если нет релевантных, берем лучший по score
                relevant_goods = [scored_goods[0]]
            
            # Среди релевантных сначала ищем товары с наличием
            available_goods = [g for g in relevant_goods if str(g['availability']).lower() not in ['0', 'нет', 'неизвестно', '']]
            if available_goods:
                # Среди доступных выбираем с минимальной ценой
                best_good = min(available_goods, key=lambda x: x['price'] or float('inf'))
            else:
                # Если нет доступных, выбираем с минимальной ценой среди релевантных
                best_good = min(relevant_goods, key=lambda x: x['price'] or float('inf'))
        
        if best_good:
            result.update({
                'found_name': best_good['name'],
                'article': best_good['article'],
                'unit': best_good['unit'],
                'unit_price': best_good['price'],
                'availability': best_good['availability'],
                'status': 'ok'
            })
        
        # Сохраняем топ-3 альтернатив (кроме основного)
        alternatives = []
        for g in scored_goods[:4]:  # топ-4, исключая основной если он там
            if g != best_good:
                alternatives.append({
                    'name': g['name'],
                    'article': g['article'],
                    'price': g['price'],
                    'availability': g['availability']
                })
            if len(alternatives) >= 3:
                break
        result['alternatives'] = json.dumps(alternatives, ensure_ascii=False)
        
    except Exception as e:
        logger.error(f"Ошибка при извлечении данных из ETM: {str(e)}")
        result['status'] = 'error'
    
    return result


def search_etm(name: str) -> Dict[str, Any]:
    """
    Поиск товаров в ETM через API.
    Использует поисковый endpoint для получения списка товаров.
    Результаты кэшируются с TTL.
    
    Args:
        name: Название товара для поиска
    
    Returns:
        Dict: Результат поиска с списком товаров или ошибка
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
        
        # Используем поисковый endpoint для получения списка товаров
        # POST /v1/goods/search
        search_url = f"{ETM_API_BASE_URL}/v1/goods/search"
        params = {
            "session-id": session_key,
            "query": normalized_name,
            "limit": 10  # Ограничиваем до 10 результатов
        }
        
        logger.info(f"Поиск товаров в ETM API: {name}")
        response = requests.post(
            search_url,
            json=params,  # POST с JSON телом
            timeout=REQUEST_TIMEOUT
        )
        
        # Проверка статуса ответа
        response.raise_for_status()
        
        data = response.json()
        if data.get("status", {}).get("code") != 200:
            logger.warning(f"ETM API вернул ошибку: {data.get('status', {}).get('message')}")
            return {"error": "API_ERROR", "message": data.get('status', {}).get('message')}
        
        # Извлекаем список товаров
        goods = data.get("data", {}).get("goods", [])
        if not goods:
            logger.info(f"Товары не найдены для: {normalized_name}")
            result = {"status": "not_found", "data": []}
        else:
            logger.info(f"Найдено {len(goods)} товаров для: {normalized_name}")
            result = {"status": "success", "data": goods}
        
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
    <meta name="description" content="Professional Excel processing with ETM API integration">
    <title>ETM Excel Processor - Pro</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        :root {
            --color-primary: #1a1a1a;
            --color-secondary: #666666;
            --color-border: #e0e0e0;
            --color-bg: #ffffff;
            --color-bg-light: #f9fafb;
            --color-accent: #3b82f6;
            --color-accent-hover: #2563eb;
            --color-success: #10b981;
            --color-error: #ef4444;
            --shadow-sm: 0 1px 2px 0 rgba(0, 0, 0, 0.05);
            --shadow-md: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06);
            --shadow-lg: 0 10px 15px -3px rgba(0, 0, 0, 0.1), 0 4px 6px -2px rgba(0, 0, 0, 0.05);
            --shadow-xl: 0 20px 25px -5px rgba(0, 0, 0, 0.1), 0 10px 10px -5px rgba(0, 0, 0, 0.04);
            --transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
        }
        
        html, body {
            height: 100%;
        }
        
        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Helvetica Neue', sans-serif;
            background: var(--color-bg-light);
            color: var(--color-primary);
            line-height: 1.6;
            overflow-x: hidden;
        }
        
        .wrapper {
            min-height: 100vh;
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            padding: 24px;
        }
        
        .container {
            width: 100%;
            max-width: 600px;
            animation: fadeInUp 0.6s ease-out;
        }
        
        @keyframes fadeInUp {
            from {
                opacity: 0;
                transform: translateY(20px);
            }
            to {
                opacity: 1;
                transform: translateY(0);
            }
        }
        
        .header-section {
            text-align: center;
            margin-bottom: 48px;
        }
        
        .logo {
            display: inline-flex;
            align-items: center;
            gap: 8px;
            margin-bottom: 24px;
            font-weight: 600;
            font-size: 18px;
            color: var(--color-primary);
        }
        
        .logo-icon {
            font-size: 24px;
        }
        
        h1 {
            font-size: 32px;
            font-weight: 700;
            margin-bottom: 12px;
            letter-spacing: -0.5px;
        }
        
        .subtitle {
            font-size: 15px;
            color: var(--color-secondary);
            margin-bottom: 0;
            line-height: 1.5;
        }
        
        .main-section {
            background: var(--color-bg);
            border: 1px solid var(--color-border);
            border-radius: 14px;
            padding: 40px;
            box-shadow: var(--shadow-lg);
        }
        
        form {
            display: flex;
            flex-direction: column;
            gap: 24px;
        }
        
        .upload-wrapper {
            display: flex;
            flex-direction: column;
            gap: 12px;
        }
        
        .upload-label {
            font-size: 13px;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            color: var(--color-secondary);
        }
        
        .upload-area {
            border: 2px solid var(--color-border);
            border-radius: 12px;
            padding: 48px 24px;
            text-align: center;
            cursor: pointer;
            transition: var(--transition);
            background: var(--color-bg-light);
            position: relative;
            overflow: hidden;
        }
        
        .upload-area::before {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background: linear-gradient(135deg, rgba(59, 130, 246, 0.05) 0%, transparent 100%);
            opacity: 0;
            transition: opacity 0.3s ease;
            pointer-events: none;
        }
        
        .upload-area:hover {
            border-color: var(--color-accent);
            background: linear-gradient(135deg, rgba(59, 130, 246, 0.02) 0%, transparent 100%);
        }
        
        .upload-area:hover::before {
            opacity: 1;
        }
        
        .upload-area.dragover {
            border-color: var(--color-accent);
            background: rgba(59, 130, 246, 0.05);
            transform: scale(1.02);
        }
        
        .upload-icon {
            font-size: 56px;
            margin-bottom: 16px;
            display: block;
            filter: drop-shadow(0 0 0px transparent);
            transition: filter 0.3s ease;
        }
        
        .upload-area:hover .upload-icon {
            filter: drop-shadow(0 4px 8px rgba(59, 130, 246, 0.2));
        }
        
        .upload-text {
            color: var(--color-primary);
            font-size: 15px;
            font-weight: 600;
            margin-bottom: 6px;
        }
        
        .upload-hint {
            color: var(--color-secondary);
            font-size: 13px;
            margin-bottom: 0;
        }
        
        input[type="file"] {
            display: none;
        }
        
        .file-info {
            padding: 12px 16px;
            background: rgba(16, 185, 129, 0.05);
            border: 1px solid rgba(16, 185, 129, 0.2);
            border-radius: 8px;
            display: none;
            align-items: center;
            gap: 10px;
            animation: slideDown 0.3s ease-out;
        }
        
        .file-info.show {
            display: flex;
        }
        
        @keyframes slideDown {
            from {
                opacity: 0;
                transform: translateY(-10px);
            }
            to {
                opacity: 1;
                transform: translateY(0);
            }
        }
        
        .file-info-icon {
            font-size: 20px;
            color: var(--color-success);
        }
        
        .file-info-content {
            flex: 1;
        }
        
        .file-name {
            font-size: 13px;
            font-weight: 600;
            color: var(--color-primary);
            word-break: break-all;
        }
        
        .button-group {
            display: flex;
            gap: 12px;
        }
        
        button {
            flex: 1;
            padding: 12px 24px;
            border: none;
            border-radius: 10px;
            font-size: 15px;
            font-weight: 600;
            cursor: pointer;
            transition: var(--transition);
            letter-spacing: 0.3px;
        }
        
        .btn-process {
            background: var(--color-accent);
            color: white;
            box-shadow: var(--shadow-md);
        }
        
        .btn-process:hover:not(:disabled) {
            background: var(--color-accent-hover);
            box-shadow: var(--shadow-lg);
            transform: translateY(-1px);
        }
        
        .btn-process:active:not(:disabled) {
            transform: translateY(0);
        }
        
        .btn-process:disabled {
            opacity: 0.5;
            cursor: not-allowed;
            transform: none;
        }
        
        .btn-clear {
            background: var(--color-bg-light);
            color: var(--color-primary);
            border: 1px solid var(--color-border);
        }
        
        .btn-clear:hover {
            background: var(--color-border);
            border-color: var(--color-secondary);
        }
        
        .loading {
            display: none;
            text-align: center;
            padding: 24px;
        }
        
        .loading.show {
            display: block;
            animation: fadeIn 0.3s ease-out;
        }
        
        @keyframes fadeIn {
            from { opacity: 0; }
            to { opacity: 1; }
        }
        
        .spinner {
            width: 40px;
            height: 40px;
            margin: 0 auto 12px;
            border: 3px solid var(--color-border);
            border-top-color: var(--color-accent);
            border-radius: 50%;
            animation: spin 1s linear infinite;
        }
        
        @keyframes spin {
            to { transform: rotate(360deg); }
        }
        
        .loading-text {
            font-size: 14px;
            color: var(--color-secondary);
            font-weight: 500;
        }
        
        .alert {
            padding: 16px;
            border-radius: 10px;
            font-size: 14px;
            display: none;
            align-items: center;
            gap: 12px;
            animation: slideDown 0.3s ease-out;
        }
        
        .alert.show {
            display: flex;
        }
        
        .alert-icon {
            font-size: 20px;
            flex-shrink: 0;
        }
        
        .alert-content {
            flex: 1;
        }
        
        .alert-success {
            background: rgba(16, 185, 129, 0.1);
            border: 1px solid rgba(16, 185, 129, 0.2);
            color: #065f46;
        }
        
        .alert-success .alert-icon {
            color: var(--color-success);
        }
        
        .alert-error {
            background: rgba(239, 68, 68, 0.1);
            border: 1px solid rgba(239, 68, 68, 0.2);
            color: #7f1d1d;
        }
        
        .alert-error .alert-icon {
            color: var(--color-error);
        }
        
        .info-box {
            background: rgba(59, 130, 246, 0.05);
            border: 1px solid rgba(59, 130, 246, 0.2);
            border-radius: 10px;
            padding: 16px;
            font-size: 13px;
            color: #1e40af;
            margin-bottom: 24px;
            line-height: 1.6;
        }
        
        .info-box strong {
            font-weight: 600;
            color: #1e3a8a;
        }
        
        @media (max-width: 640px) {
            .main-section {
                padding: 24px;
            }
            
            h1 {
                font-size: 26px;
            }
            
            .upload-area {
                padding: 36px 20px;
            }
            
            .button-group {
                flex-direction: column;
            }
            
            button {
                width: 100%;
            }
        }
    </style>
</head>
<body>
    <div class="wrapper">
        <div class="container">
            <!-- Header -->
            <div class="header-section">
                <div class="logo">
                    <span class="logo-icon">📊</span>
                    <span>ETM Excel Pro</span>
                </div>
                <h1>Excel Processing</h1>
                <p class="subtitle">Upload and process your Excel files with ETM API integration</p>
            </div>

            <!-- Main Form -->
            <div class="main-section">
                <form id="uploadForm" enctype="multipart/form-data">
                    
                    <!-- Info Box -->
                    <div class="info-box">
                        Your file must contain:<br>
                        <strong>Required:</strong> "Наименование" (name) and "Количество" (quantity)<br>
                        <strong>Result includes:</strong> Found name, article, unit, price, availability, status
                    </div>

                    <!-- Upload Area -->
                    <div class="upload-wrapper">
                        <label class="upload-label">Upload Excel File</label>
                        <div class="upload-area" id="uploadArea" onclick="document.getElementById('fileInput').click()">
                            <span class="upload-icon">📄</span>
                            <div class="upload-text">Click to upload or drag & drop</div>
                            <div class="upload-hint">XLSX format, up to 10 MB</div>
                        </div>
                        <input type="file" id="fileInput" name="file" accept=".xlsx" />
                    </div>

                    <!-- File Info -->
                    <div class="file-info" id="fileInfo">
                        <span class="file-info-icon">✓</span>
                        <div class="file-info-content">
                            <div class="file-name" id="fileName"></div>
                        </div>
                    </div>

                    <!-- Action Buttons -->
                    <div class="button-group">
                        <button type="submit" class="btn-process" id="processBtn" disabled>
                            Process File
                        </button>
                        <button type="button" class="btn-clear" id="clearBtn">
                            Clear
                        </button>
                    </div>

                    <!-- Loading State -->
                    <div class="loading" id="loading">
                        <div class="spinner"></div>
                        <p class="loading-text">Processing your file...</p>
                    </div>

                    <!-- Alerts -->
                    <div class="alert alert-success" id="successAlert">
                        <span class="alert-icon">✓</span>
                        <div class="alert-content" id="successMessage"></div>
                    </div>
                    <div class="alert alert-error" id="errorAlert">
                        <span class="alert-icon">✕</span>
                        <div class="alert-content" id="errorMessage"></div>
                    </div>
                </form>
            </div>
        </div>
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
        const successAlert = document.getElementById('successAlert');
        const errorAlert = document.getElementById('errorAlert');
        const successMessage = document.getElementById('successMessage');
        const errorMessage = document.getElementById('errorMessage');

        // Prevent default drag behaviors
        ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
            uploadArea.addEventListener(eventName, preventDefaults, false);
        });

        function preventDefaults(e) {
            e.preventDefault();
            e.stopPropagation();
        }

        // Highlight drop area
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
            hideAlerts();
            if (file) {
                fileName.textContent = file.name;
                fileInfo.classList.add('show');
                processBtn.disabled = false;
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
            hideAlerts();
        });

        // Form submission
        uploadForm.addEventListener('submit', async (e) => {
            e.preventDefault();

            const file = fileInput.files[0];
            if (!file) return;

            loading.classList.add('show');
            hideAlerts();
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

                    loading.classList.remove('show');
                    fileInput.value = '';
                    fileInfo.classList.remove('show');

                    successMessage.textContent = 'File processed successfully! Your result has been downloaded.';
                    successAlert.classList.add('show');
                } else {
                    const error = await response.json();
                    throw new Error(error.detail || 'Failed to process file');
                }
            } catch (error) {
                loading.classList.remove('show');
                errorMessage.textContent = error.message;
                errorAlert.classList.add('show');
                processBtn.disabled = false;
            }
        });

        function hideAlerts() {
            successAlert.classList.remove('show');
            errorAlert.classList.remove('show');
        }
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
    
    Возвращает Excel файл с колонками:
    - Исходное название
    - Найденное название  
    - Артикул
    - Единица измерения
    - Цена за единицу
    - Наличие / срок
    - Статус
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
        
        # СПОСОБ 1: Пробуем читать с заголовками
        df = pd.read_excel(file_buffer, engine='openpyxl')
        
        # СПОСОБ 2: Если первая колонка - это число (индекс), значит файл БЕЗ заголовков
        # или если колонки это просто "Unnamed" или значения из данных
        has_proper_headers = False
        
        # Проверяем есть ли разумные названия колонок
        for col in df.columns:
            col_str = str(col).lower()
            # Если есть "Наименование", "Name", "Product" и т.д. - это заголовки
            if any(x in col_str for x in ['наименование', 'название', 'qty', 'quantity', 'кол', 'name', 'product', 'item']):
                has_proper_headers = True
                break
        
        logger.info(f"Попытка 1: Колонки={list(df.columns)}, Заголовки OK={has_proper_headers}")
        
        # Если заголовки не выглядят правильно, пробуем без заголовков
        if not has_proper_headers:
            logger.info("Колонки не выглядят как заголовки - переоткрываю с header=None")
            file_buffer = BytesIO(contents)
            df = pd.read_excel(file_buffer, engine='openpyxl', header=None)
            
            # Даем стандартные имена первым двум колонкам (название и количество)
            if len(df.columns) >= 2:
                col_names = {0: 'Наименование', 1: 'Количество'}
                for i in range(2, len(df.columns)):
                    col_names[i] = f'Доп_колонка_{i}'
                df = df.rename(columns=col_names)
                logger.info(f"Переименованы колонки: {list(df.columns)}")
        
        logger.info(f"Итоговые колонки: {list(df.columns)}")
        logger.info(f"Строк данных: {len(df)}")
        
        # Поиск колонок с совпадением (игнорируя регистр)
        def find_column(df, patterns):
            for col in df.columns:
                col_lower = str(col).lower().strip()
                for pattern in patterns:
                    if col_lower == pattern.lower() or pattern.lower() in col_lower:
                        return col
            return None
        
        # Ищем колонки
        name_col = find_column(df, ["наименование", "название", "product", "name", "товар", "item"])
        qty_col = find_column(df, ["количество", "quantity", "кол", "qty", "кол-во"])
        
        logger.info(f"Найдены: name_col='{name_col}', qty_col='{qty_col}'")
        
        if not name_col or not qty_col:
            cols = ", ".join(str(c) for c in df.columns[:5])
            detail_msg = f"Не найдены необходимые колонки. Доступны: {cols}"
            logger.error(detail_msg)
            raise HTTPException(status_code=400, detail=f"❌ {detail_msg}")
        
        logger.info(f"Используем колонки: '{name_col}' для названия, '{qty_col}' для количества")
        
        # Преобразование данных в требуемый формат и обогащение данными из ETM
        result = []
        for index, row in df.iterrows():
            try:
                original_name = str(row[name_col]).strip()
                quantity = int(row[qty_col])
            except (ValueError, TypeError) as e:
                logger.warning(f"Ошибка при обработке строки {index}: {e}")
                continue
            
            # Поиск товара в ETM API
            etm_result = search_etm(original_name)
            
            # Извлечение данных с выбором лучшего товара
            etm_data = extract_etm_data(etm_result, original_name)
            
            # Формирование элемента результата
            item = {
                "Исходное название": original_name,
                "Найденное название": etm_data['found_name'],
                "Артикул": etm_data['article'],
                "Единица измерения": etm_data['unit'],
                "Цена за единицу": etm_data['unit_price'],
                "Наличие / срок": etm_data['availability'],
                "Статус": etm_data['status']
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
        error_msg = str(e)
        logger.error(f"Ошибка при обработке файла: {error_msg}")
        
        # Если это ошибка авторизации ETM API
        if "авторизац" in error_msg.lower() or "login" in error_msg.lower() or "unauthorized" in error_msg.lower():
            raise HTTPException(
                status_code=401,
                detail="❌ Ошибка авторизации ETM API. Проверьте учетные данные в переменных окружения ETM_LOGIN и ETM_PASSWORD"
            )
        
        # Если это ошибка соединения
        if "connection" in error_msg.lower() or "timeout" in error_msg.lower():
            raise HTTPException(
                status_code=503,
                detail="❌ Ошибка соединения с ETM API. Проверьте интернет соединение и доступность API"
            )
        
        # Общая ошибка
        raise HTTPException(
            status_code=500,
            detail=f"❌ Ошибка обработки файла: {error_msg[:150]}"
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
