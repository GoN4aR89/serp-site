#!/usr/bin/env python3
"""
Скрипт миграции базы данных
Добавляет таблицу Project и колонку project_id в Comparison
"""

import sqlite3
import os

# Путь к базе данных
DB_PATH = os.path.join(os.path.dirname(__file__), 'instance', 'serp_comparator.db')

def migrate():
    if not os.path.exists(DB_PATH):
        print(f"База данных не найдена: {DB_PATH}")
        return
    
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    try:
        # 1. Создаем таблицу project
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS project (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                name VARCHAR(200) NOT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (user_id) REFERENCES user (id)
            )
        ''')
        print("✓ Таблица 'project' создана")
        
        # 2. Проверяем, есть ли уже колонка project_id в comparison
        cursor.execute("PRAGMA table_info(comparison)")
        columns = [col[1] for col in cursor.fetchall()]
        
        if 'project_id' not in columns:
            # 3. Добавляем колонку project_id
            cursor.execute('''
                ALTER TABLE comparison 
                ADD COLUMN project_id INTEGER 
                REFERENCES project (id)
            ''')
            print("✓ Колонка 'project_id' добавлена в таблицу 'comparison'")
        else:
            print("ℹ Колонка 'project_id' уже существует")
        
        # 4. Создаем индексы для производительности
        cursor.execute('''
            CREATE INDEX IF NOT EXISTS idx_comparison_project_id 
            ON comparison (project_id)
        ''')
        print("✓ Индекс 'idx_comparison_project_id' создан")
        
        cursor.execute('''
            CREATE INDEX IF NOT EXISTS idx_project_user_id 
            ON project (user_id)
        ''')
        print("✓ Индекс 'idx_project_user_id' создан")
        
        # 5. Добавляем колонки для стартовых метрик (5 тональностей × 4 группы + тип ввода + total_urls)
        cursor.execute("PRAGMA table_info(project)")
        project_columns = [col[1] for col in cursor.fetchall()]
        
        # 28 полей: для каждой группы (Яндекс/Гугл × ТОП-10/ТОП-20) по 7 полей
        baseline_columns = [
            # Яндекс ТОП-20
            ('y20_client_site', 'REAL'), ('y20_positive', 'REAL'), ('y20_neutral', 'REAL'),
            ('y20_negative', 'REAL'), ('y20_irrelevant', 'REAL'),
            ('y20_input_type', 'VARCHAR(20)'), ('y20_total_urls', 'INTEGER'),
            # Яндекс ТОП-10  
            ('y10_client_site', 'REAL'), ('y10_positive', 'REAL'), ('y10_neutral', 'REAL'),
            ('y10_negative', 'REAL'), ('y10_irrelevant', 'REAL'),
            ('y10_input_type', 'VARCHAR(20)'), ('y10_total_urls', 'INTEGER'),
            # Google ТОП-20
            ('g20_client_site', 'REAL'), ('g20_positive', 'REAL'), ('g20_neutral', 'REAL'),
            ('g20_negative', 'REAL'), ('g20_irrelevant', 'REAL'),
            ('g20_input_type', 'VARCHAR(20)'), ('g20_total_urls', 'INTEGER'),
            # Google ТОП-10
            ('g10_client_site', 'REAL'), ('g10_positive', 'REAL'), ('g10_neutral', 'REAL'),
            ('g10_negative', 'REAL'), ('g10_irrelevant', 'REAL'),
            ('g10_input_type', 'VARCHAR(20)'), ('g10_total_urls', 'INTEGER'),
        ]
        
        for col_name, col_type in baseline_columns:
            if col_name not in project_columns:
                try:
                    cursor.execute(f'''
                        ALTER TABLE project
                        ADD COLUMN {col_name} {col_type}
                    ''')
                    print(f"✓ Колонка '{col_name}' добавлена")
                except Exception as e:
                    print(f"⚠ Колонка '{col_name}' возможно уже существует: {e}")
            else:
                print(f"ℹ Колонка '{col_name}' уже существует")

        # 6. Добавляем колонки для PPTX диаграмм в comparison
        cursor.execute("PRAGMA table_info(comparison)")
        comparison_columns = [col[1] for col in cursor.fetchall()]

        pptx_columns = [
            ('chart_y20_pptx_path', 'VARCHAR(500)'),
            ('chart_g20_pptx_path', 'VARCHAR(500)'),
            ('chart_total20_pptx_path', 'VARCHAR(500)'),
            ('chart_y10_pptx_path', 'VARCHAR(500)'),
            ('chart_g10_pptx_path', 'VARCHAR(500)'),
            ('chart_total10_pptx_path', 'VARCHAR(500)')
        ]

        for col_name, col_type in pptx_columns:
            if col_name not in comparison_columns:
                try:
                    cursor.execute(f'''
                        ALTER TABLE comparison
                        ADD COLUMN {col_name} {col_type}
                    ''')
                    print(f"✓ Колонка '{col_name}' добавлена в comparison")
                except Exception as e:
                    print(f"⚠ Колонка '{col_name}' возможно уже существует: {e}")
            else:
                print(f"ℹ Колонка '{col_name}' уже существует в comparison")
        
        conn.commit()
        print("\n✅ Миграция завершена успешно!")
        
    except Exception as e:
        conn.rollback()
        print(f"\n❌ Ошибка миграции: {e}")
        raise
    finally:
        conn.close()

if __name__ == '__main__':
    migrate()
