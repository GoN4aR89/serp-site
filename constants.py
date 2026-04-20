"""
Общие константы для SERP Comparator
"""

# Цвета тональности для Excel (русские названия)
SENTIMENT_COLORS_RU = {
    'Домашний сайт': 'fff2cc',
    'Негативная': 'f4cccc',
    'Нейтральная': 'c9daf8',
    'Нерелевантная': 'cccccc',
    'Позитивная': 'd9ead3'
}

# Цвета тональности для Excel (английские названия)
SENTIMENT_COLORS_EN = {
    'positive': 'D9EAD3',
    'negative': 'F4CCCC',
    'neutral': 'C9DAF8',
    'irrelevant': 'CCCCCC',
    'client_site': 'FFF2CC'
}

# Полный список цветов тональности для распознавания (из comparator.py)
SENTIMENT_COLOR_VARIANTS = {
    'positive': ['00FF00', '00C000', '92D050', '00B050', 'C6EFCE', 'C8F3C2', '02FF00', 'D9EAD3', 'B5D7A8', '93C47D', '6AA84F', '38761D', 'BDF5BD', '00FF00', 'B6D7A8'],
    'negative': ['FF0000', 'C00000', 'FF6B6B', 'FFC7CE', 'E74C3C', 'F3C0BF', '980000', 'FF0000', 'E6B8AF', 'F5CBCC', 'DD7E6B', 'EA9999', 'CD4025', 'E06666', 'A61D01', 'CC0100', '990001', '85210C', 'FDBEBE', 'FF0000'],
    'neutral': ['00FFFF', '00B0F0', '87CEEB', 'ADD8E6', '5DADE2', 'BDD6FB', '02FFFF', '4987E8', '0000FF', 'C9DAF8', 'D0E2F3', 'A0C5E8', '6D9EEB', '70A8DC', '3C78D8', '3D85C6', 'B6D7FF', 'B6D7FF', 'B6D7FF'],
    'irrelevant': ['808080', 'C0C0C0', 'D3D3D3', 'A9A9A9', '95A5A6', 'AEAEAE', '434343', '666666', '999999', 'B7B7B7', 'CCCCCC', 'D9D9D9', 'EFEFEF', 'F3F3F3', 'AEAEAE', 'D9D9D9', 'A6A6A6', 'BFBFBF'],
    'client_site': ['FFFF00', 'FFD700', 'FFF2CC', 'F1C40F', 'F7DC6F', 'FFDCA1', 'F9DDA8', 'FF9900', 'FCE5CD', 'F9CB9C', 'F6B26B', 'E69138', 'F1C231', 'FFD966', 'FFE59A', 'FFFF00', 'F1C231']
}

# Цвета для диаграмм
SENTIMENT_CHART_COLORS = {
    'positive': '#6aa84f',      # Позитивная – зелёный
    'negative': '#cc0000',      # Негативная – красный
    'neutral': '#6fa8dc',       # Нейтральная – голубой
    'irrelevant': '#b7b7b7',    # Нерелевантная – серый
    'client_site': '#ff9900',   # Домашний сайт – оранжевый
    'unknown': '#808080'        # Неопределенная – тёмно-серый
}

# Названия тональностей на русском языке
SENTIMENT_NAMES_RU = {
    'positive': 'Позитивная',
    'negative': 'Негативная',
    'neutral': 'Нейтральная',
    'irrelevant': 'Нерелевантная',
    'client_site': 'Домашний сайт',
    'unknown': 'Неопределенная'
}

# Эмодзи для отчетов
SENTIMENT_EMOJI = {
    'client_site': '🟡',  # Желтый кружок - домашний сайт
    'positive': '🟢',      # Зеленый кружок
    'negative': '🔴',      # Красный кружок
    'neutral': '🔵',       # Синий кружок
    'irrelevant': '⚪',    # Белый/серый кружок
    'unknown': '⚫'        # Черный кружок
}

# Цвета для заголовков Excel
HEADER_FILL_COLOR = "3d85c6"
HEADER_FONT_COLOR = "FFFFFF"

# Разрешенные расширения файлов
ALLOWED_EXTENSIONS = {'xlsx'}

# Настройки шрифтов
DEFAULT_FONT_NAME = 'Calibri'
