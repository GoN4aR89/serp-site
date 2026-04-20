@echo off
chcp 65001 >nul

:: ===== SERP Comparator - Setup Script for Windows =====

echo 🚀 Setting up SERP Comparator...

:: 1. Проверка и создание виртуального окружения
if not exist ".venv" (
    echo 📦 Creating virtual environment...
    python -m venv .venv
)

:: 2. Активация виртуального окружения
echo ⚙️ Activating virtual environment...
call .venv\Scripts\activate.bat

:: 3. Обновление pip
echo ⬆️ Upgrading pip...
python -m pip install --upgrade pip

:: 4. Установка зависимостей
echo 📚 Installing dependencies...
pip install -r requirements.txt

:: 5. Проверка .env файла
if not exist ".env" (
    echo ⚠️ .env file not found. Copying from .env.example...
    copy .env.example .env
    echo ⚠️ Please edit .env file with your configuration
)

:: 6. Инициализация БД
echo 🗄️ Initializing database...
python init_db.py

echo.
echo ✅ Setup complete!
echo.
echo 📝 To start the server, run:
echo    python app.py
echo.
echo 🌐 Then open: http://127.0.0.1:5000

pause
