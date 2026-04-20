#!/bin/bash

# ===== SERP Comparator - Setup Script =====

echo "🚀 Setting up SERP Comparator..."

# 1. Проверка и создание виртуального окружения
if [ ! -d ".venv" ]; then
    echo "📦 Creating virtual environment..."
    python3 -m venv .venv
fi

# 2. Активация виртуального окружения
echo "⚙️ Activating virtual environment..."
source .venv/bin/activate

# 3. Обновление pip
echo "⬆️ Upgrading pip..."
pip install --upgrade pip

# 4. Установка зависимостей
echo "📚 Installing dependencies..."
pip install -r requirements.txt

# 5. Проверка .env файла
if [ ! -f ".env" ]; then
    echo "⚠️ .env file not found. Copying from .env.example..."
    cp .env.example .env
    echo "⚠️ Please edit .env file with your configuration"
fi

# 6. Инициализация БД
echo "🗄️ Initializing database..."
python3 init_db.py

echo ""
echo "✅ Setup complete!"
echo ""
echo "📝 To start the server, run:"
echo "   python3 app.py"
echo ""
echo "🌐 Then open: http://127.0.0.1:5000"
