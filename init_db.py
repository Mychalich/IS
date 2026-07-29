# Модуль для автоматического создания базы данных и таблиц при первом запуске

# Библиотека для работы с PostgreSQL (подключение, выполнение запросов)
import psycopg2
# Класс ошибки подключения и функция sql для безопасного формирования имён БД
from psycopg2 import OperationalError, sql
# Конфигурация подключения к БД (хост, порт, имя БД, пользователь, пароль)
from db_config import DB_CONFIG
# Библиотека для хэширования паролей
import bcrypt


# Функция хэширования пароля с помощью bcrypt
def simple_hash_password(password: str) -> str:
    # Генерируем соль с 12 раундами (стандартная сложность)
    salt = bcrypt.gensalt(rounds=12)
    # Хэшируем пароль (сначала преобразуем строку в байты UTF-8)
    hashed = bcrypt.hashpw(password.encode('utf-8'), salt)
    # Возвращаем хэш в виде строки для сохранения в БД
    return hashed.decode('utf-8')


# Создание базы данных, если она ещё не существует
def create_database():
    # Копируем конфигурацию и меняем базу данных на системную 'postgres'
    conn_config = DB_CONFIG.copy()
    conn_config['database'] = 'postgres'

    try:
        # Подключаемся к системной БД для создания новой
        conn = psycopg2.connect(**conn_config)
        conn.autocommit = True  # Каждая команда выполняется сразу, без ожидания подтверждения)
        cursor = conn.cursor()

        # Проверяем существование БД с нужным именем
        cursor.execute("SELECT 1 FROM pg_database WHERE datname = %s", (DB_CONFIG['database'],))
        if not cursor.fetchone():
            # БД не существует — создаём
            cursor.execute(sql.SQL("CREATE DATABASE {}").format(
                sql.Identifier(DB_CONFIG['database'])))
            print(f"База данных '{DB_CONFIG['database']}' создана")
        else:
            print(f"База данных '{DB_CONFIG['database']}' уже существует")

        cursor.close()
        conn.close()
        return True
    except OperationalError as e:
        print(f"Ошибка подключения к PostgreSQL: {e}")
        return False


# Создание таблиц в базе данных
def create_tables():
    try:
        # Подключаемся к целевой базе данных
        conn = psycopg2.connect(**DB_CONFIG)
        cur = conn.cursor()

        # Список SQL-запросов на создание таблиц
        tables_sql = [
            # 1. Таблица Статус
            """
            CREATE TABLE IF NOT EXISTS Статус (
                id_статуса SERIAL PRIMARY KEY,
                наименование VARCHAR(12) NOT NULL UNIQUE
            )
            """,
            # 2. Таблица Сотрудники
            """
            CREATE TABLE IF NOT EXISTS Сотрудники (
                id_сотрудника SERIAL PRIMARY KEY,
                фамилия VARCHAR(12) NOT NULL,
                имя VARCHAR(12) NOT NULL,
                отчество VARCHAR(12),
                должность VARCHAR(50) NOT NULL,
                отдел VARCHAR(20) NOT NULL,
                телефон VARCHAR(16)
            )
            """,
            # 3. Таблица Техника
            """
            CREATE TABLE IF NOT EXISTS Техника (
                id_техники SERIAL PRIMARY KEY,
                инвентарный_номер VARCHAR(20) NOT NULL UNIQUE,
                наименование VARCHAR(50) NOT NULL,
                модель VARCHAR(30),
                серийный_номер VARCHAR(20),
                дата_поступления DATE NOT NULL,
                id_статуса INTEGER NOT NULL REFERENCES Статус(id_статуса)
            )
            """,
            # 4. Таблица Перемещение
            """
            CREATE TABLE IF NOT EXISTS Перемещение (
                id_перемещения SERIAL PRIMARY KEY,
                id_сотрудника INTEGER NOT NULL REFERENCES Сотрудники(id_сотрудника),
                id_техники INTEGER NOT NULL REFERENCES Техника(id_техники),
                дата DATE NOT NULL
            )
            """,
            # 5. Таблица Ремонт
            """
            CREATE TABLE IF NOT EXISTS Ремонт (
                id_ремонта SERIAL PRIMARY KEY,
                id_техники INTEGER NOT NULL REFERENCES Техника(id_техники),
                дата_заявки DATE NOT NULL,
                дата_ремонта DATE,
                описание VARCHAR(60),
                стоимость NUMERIC(10,2) NOT NULL DEFAULT 0,
                статус VARCHAR(12) NOT NULL
            )
            """,
            # 6. Таблица Пользователи
            """
            CREATE TABLE IF NOT EXISTS Пользователи (
                id_пользователя SERIAL PRIMARY KEY,
                логин VARCHAR(15) NOT NULL UNIQUE,
                пароль VARCHAR(60) NOT NULL,  
                роль VARCHAR(20) NOT NULL CHECK (роль IN ('admin', 'technician')),
                id_сотрудника INTEGER NOT NULL REFERENCES Сотрудники(id_сотрудника)
            )
            """
        ]

        # Последовательно выполняем каждый SQL-запрос
        for sql_query in tables_sql:
            cur.execute(sql_query)

        # Фиксируем изменения в базе данных
        conn.commit()
        print("Все таблицы успешно созданы")
        cur.close()
        conn.close()
        return True
    except OperationalError as e:
        print(f"Ошибка при создании таблиц: {e}")
        return False


# Заполнение таблиц начальными (тестовыми) данными
def insert_initial_data():
    try:
        conn = psycopg2.connect(**DB_CONFIG)
        cur = conn.cursor()

        # --- Статусы техники ---
        cur.execute("SELECT COUNT(*) FROM Статус")
        if cur.fetchone()[0] == 0:
            statuses = [('на складе',), ('в работе',), ('в ремонте',), ('списана',)]
            cur.executemany("INSERT INTO Статус (наименование) VALUES (%s)", statuses)
            print("Добавлены статусы техники")

        # --- Сотрудники ---
        cur.execute("SELECT COUNT(*) FROM Сотрудники")
        if cur.fetchone()[0] == 0:
            employees = [
                ('Иванов', 'Иван', 'Иванович', 'Инженер', 'Технический отдел', '+7-999-123-45-67'),
                ('Петров', 'Петр', 'Петрович', 'Сотрудник', 'Отдел продаж', '+7-999-123-45-68'),
                ('Сидорова', 'Анна', 'Сергеевна', 'Бухгалтер', 'Бухгалтерия', '+7-999-123-45-69'),
                ('Козлов', 'Дмитрий', 'Алексеевич', 'Руководитель отдела', 'Технический отдел', '+7-999-123-45-70'),
                ('Смирнова', 'Екатерина', 'Владимировна', 'Менеджер', 'Отдел продаж', '+7-999-123-45-71')
            ]
            cur.executemany("""
                INSERT INTO Сотрудники (фамилия, имя, отчество, должность, отдел, телефон) 
                VALUES (%s, %s, %s, %s, %s, %s)
            """, employees)
            print("Добавлены тестовые сотрудники")

        # --- Пользователи (учётные записи) ---
        cur.execute("SELECT COUNT(*) FROM Пользователи")
        if cur.fetchone()[0] == 0:
            users = [
                ('admin', simple_hash_password('admin123'), 'admin', 1),
                ('technician', simple_hash_password('tech123'), 'technician', 2),
                ('buch', simple_hash_password('buch123'), 'technician', 3)
            ]
            cur.executemany("""
                INSERT INTO Пользователи (логин, пароль, роль, id_сотрудника)
                VALUES (%s, %s, %s, %s)
            """, users)
            print("Добавлены учетные записи пользователей (пароли хэшированы bcrypt)")

        # --- Техника ---
        cur.execute("SELECT COUNT(*) FROM Техника")
        if cur.fetchone()[0] == 0:
            equipment = [
                ('IT-001', 'Моноблок HP 200 G3 AiO', 'HP 200 G3', 'SN001', '2024-01-15', 2),
                ('IT-002', 'Ноутбук Lenovo ThinkPad', 'ThinkPad T14', 'SN002', '2024-02-10', 2),
                ('IT-003', 'Принтер HP LaserJet', 'LaserJet Pro M402', 'SN003', '2024-01-20', 1),
                ('IT-004', 'Моноблок HP 200 G3 AiO', 'HP 200 G3', 'SN004', '2024-03-01', 2),
                ('IT-005', 'Сервер HP ProLiant', 'DL380 Gen10', 'SN005', '2023-11-15', 1),
                ('IT-006', 'Коммутатор Cisco', 'Cisco 2960', 'SN006', '2024-01-05', 2),
                ('IT-008', 'МФУ Kyocera', 'ECOSYS M2635', 'SN008', '2024-03-10', 1)
            ]
            cur.executemany("""
                INSERT INTO Техника (инвентарный_номер, наименование, модель, 
                                    серийный_номер, дата_поступления, id_статуса) 
                VALUES (%s, %s, %s, %s, %s, %s)
            """, equipment)
            print("Добавлены тестовые единицы техники")

            # --- Перемещения (история выдачи техники) ---
            movements = [
                (1, 1, '2024-01-15'), (2, 2, '2024-02-10'), (4, 4, '2024-03-01'),
                (6, 1, '2024-01-05'), (1, 4, '2024-03-15'), (4, 1, '2024-04-01')
            ]
            cur.executemany("""
                INSERT INTO Перемещение (id_техники, id_сотрудника, дата) 
                VALUES (%s, %s, %s)
            """, movements)
            print("Добавлена история перемещений")

            # --- Ремонты ---
            repairs = [
                (7, '2024-02-20', '2024-02-25', 'Не включается', 5000.00, 'выполнен'),
                (2, '2024-03-10', '2024-03-15', 'Замена клавиатуры', 3000.00, 'выполнен'),
                (1, '2024-04-01', None, 'Проблемы с загрузкой ОС', 0, 'в работе')
            ]
            cur.executemany("""
                INSERT INTO Ремонт (id_техники, дата_заявки, дата_ремонта, описание, стоимость, статус) 
                VALUES (%s, %s, %s, %s, %s, %s)
            """, repairs)
            print("Добавлена история ремонтов")

        # Фиксируем все изменения и закрываем соединение
        conn.commit()
        cur.close()
        conn.close()
        return True
    except OperationalError as e:
        print(f"Ошибка при добавлении начальных данных: {e}")
        return False


# Полная инициализация базы данных (создание БД, таблиц, наполнение данными)
def init_database():
    # Шаг 1: создаём базу данных, если её нет
    if not create_database():
        return False

    # Шаг 2: создаём все таблицы
    if not create_tables():
        return False

    # Шаг 3: заполняем таблицы начальными данными
    if not insert_initial_data():
        return False

    # Инициализация прошла успешно
    return True


# Если Данный файл запущен напрямую, выполняем инициализацию
if __name__ == "__main__":
    init_database()