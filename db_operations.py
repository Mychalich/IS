# db_operations.py
# Функции для работы с базой данных (CRUD операции, аутентификация, отчёты)

# Библиотека для работы с PostgreSQL
import psycopg2
# Класс ошибки подключения
from psycopg2 import OperationalError
# Конфигурация подключения к БД (хост, порт, имя БД, пользователь, пароль)
from db_config import DB_CONFIG
# Функция инициализации БД (создание таблиц)
from init_db import init_database
# Модуль для работы с датами
from datetime import date
# Библиотека для хэширования паролей (алгоритм bcrypt)
import bcrypt
# Модуль для работы с регулярными выражениями (проверка сложности пароля)
import re


# Флаг для отслеживания инициализации БД (чтобы не инициализировать повторно)
_db_initialized = False

# Конфигурация требований к паролю
PASSWORD_POLICY = {
    'min_length': 12,          # Минимальная длина пароля
    'require_upper': True,     # Требовать заглавные буквы
    'require_lower': True,     # Требовать строчные буквы
    'require_digits': True,    # Требовать цифры
    'require_special': True,   # Требовать специальные символы
}


# Хеширование пароля с помощью bcrypt
def hash_password(password: str) -> str:
    # Генерируем соль с 12 раундами (стандартная сложность)
    salt = bcrypt.gensalt(rounds=12)
    # Преобразуем пароль в байты и хэшируем
    hashed = bcrypt.hashpw(password.encode('utf-8'), salt)
    # Возвращаем строку для хранения в БД
    return hashed.decode('utf-8')


# Проверка пароля (сравнение введённого с хэшем из БД)
def verify_password(password: str, hashed: str) -> bool:
    try:
        return bcrypt.checkpw(password.encode('utf-8'), hashed.encode('utf-8'))
    except Exception:
        return False


# Проверка сложности пароля
def validate_password_strength(password: str) -> tuple:
    # Проверка длины
    if len(password) < PASSWORD_POLICY['min_length']:
        return False, f"Пароль должен содержать не менее {PASSWORD_POLICY['min_length']} символов"

    # Проверка наличия заглавных букв
    if PASSWORD_POLICY['require_upper'] and not re.search(r'[A-ZА-Я]', password):
        return False, "Пароль должен содержать хотя бы одну заглавную букву"

    # Проверка наличия строчных букв
    if PASSWORD_POLICY['require_lower'] and not re.search(r'[a-zа-я]', password):
        return False, "Пароль должен содержать хотя бы одну строчную букву"

    # Проверка наличия цифр
    if PASSWORD_POLICY['require_digits'] and not re.search(r'\d', password):
        return False, "Пароль должен содержать хотя бы одну цифру"

    # Проверка наличия специальных символов
    if PASSWORD_POLICY['require_special'] and not re.search(r'[!@#$%^&*(),.?":{}|<>]', password):
        return False, "Пароль должен содержать хотя бы один специальный символ"

    # Проверка на слишком простые (словарные) пароли
    simple_passwords = ['password', '12345678', 'qwerty123', 'admin123', 'password123', 'qwerty', '12345']
    if password.lower() in simple_passwords:
        return False, "Пароль слишком простой. Выберите более сложный пароль"

    return True, "OK"


# Проверяет и инициализирует БД при первом запуске (создаёт таблицы)
def ensure_db_initialized():
    global _db_initialized
    if not _db_initialized:
        if init_database():
            _db_initialized = True
        else:
            raise Exception("Не удалось инициализировать базу данных")


# Установка соединения с БД (с автоматической инициализацией при первом вызове)
def get_connection():
    ensure_db_initialized()
    try:
        conn = psycopg2.connect(**DB_CONFIG)
        return conn   # канал для общения с сервером СУБД
    except OperationalError as e:
        print(f"Ошибка подключения к БД: {e}")
        return None


# Универсальная функция выполнения SQL-запроса с защитой от SQL-инъекций
def execute_query(query, params=None, fetch=False):
    # params: кортеж (или список) значений для подстановки вместо %s
    # fetch: если True, возвращает результат SELECT; если False, возвращает количество затронутых строк
    conn = get_connection()
    if not conn:
        return None if not fetch else []
    try:
        cursor = conn.cursor()
        cursor.execute(query, params or ())
        if fetch:
            result = cursor.fetchall()
            conn.commit()          # Фиксируем транзакцию после SELECT
            return result
        conn.commit()              # Фиксируем изменения (INSERT, UPDATE, DELETE)
        return cursor.rowcount     # Количество затронутых строк
    except Exception as e:
        conn.rollback()            # Откатываем изменения при ошибке
        print(f"Ошибка выполнения запроса: {e}")
        return None if not fetch else []
    finally:
        cursor.close()
        conn.close()


# Возвращает текущую дату в формате ГГГГ-ММ-ДД (для подстановки в поля по умолчанию)
def get_current_date():
    return str(date.today())


# ==================== РАБОТА С ТЕХНИКОЙ ====================

# Получить список техники с возможной фильтрацией по статусу
def get_equipment(status_filter=None):
    # Формируем запрос: выбираем поля техники и название статуса
    query = """
        SELECT t.id_техники, t.инвентарный_номер, t.наименование, 
               t.модель, t.серийный_номер, t.дата_поступления,
               s.наименование as статус
        FROM Техника t
        JOIN Статус s ON t.id_статуса = s.id_статуса
        WHERE 1=1
    """
    params = []
    if status_filter:
        query += " AND s.наименование = %s"
        params.append(status_filter)
    query += " ORDER BY t.дата_поступления DESC"
    return execute_query(query, params, fetch=True)


# Добавление новой единицы техники (без привязки к сотруднику, только основные характеристики)
def add_equipment(inv_number, name, model, serial, date_val, status_id):
    query = """
        INSERT INTO Техника (инвентарный_номер, наименование, модель, 
                            серийный_номер, дата_поступления, id_статуса)
        VALUES (%s, %s, %s, %s, %s, %s)
    """
    return execute_query(query, (inv_number, name, model, serial, date_val, status_id))


# Обновление статуса техники по ID
def update_equipment_status(equipment_id, status_id):
    query = "UPDATE Техника SET id_статуса = %s WHERE id_техники = %s"
    return execute_query(query, (status_id, equipment_id))

# Подсчёт общего количества единиц техники в БД
def get_equipment_count():
    query = "SELECT COUNT(*) FROM Техника"
    result = execute_query(query, fetch=True)
    return result[0][0] if result else 0


# ==================== РАБОТА С СОТРУДНИКАМИ ====================

# Получение списка всех сотрудников (сортировка по фамилии)
def get_employees():
    query = """
        SELECT id_сотрудника, фамилия, имя, отчество, должность, отдел, телефон
        FROM Сотрудники ORDER BY фамилия
    """
    return execute_query(query, fetch=True)


# Добавление нового сотрудника
def add_employee(last_name, first_name, middle_name, position, department, phone):
    query = """
        INSERT INTO Сотрудники (фамилия, имя, отчество, должность, отдел, телефон)
        VALUES (%s, %s, %s, %s, %s, %s)
    """
    return execute_query(query, (last_name, first_name, middle_name, position, department, phone))


# Поиск сотрудника по его уникальному идентификатору
def get_employee_by_id(employee_id):
    query = "SELECT * FROM Сотрудники WHERE id_сотрудника = %s"
    result = execute_query(query, (employee_id,), fetch=True)
    return result[0] if result else None


# Подсчёт общего количества сотрудников
def get_employees_count():
    query = "SELECT COUNT(*) FROM Сотрудники"
    result = execute_query(query, fetch=True)
    return result[0][0] if result else 0


# ==================== РАБОТА С ПЕРЕМЕЩЕНИЯМИ ====================

# Получение истории перемещений (выдачи/передачи техники) с ограничением LIMIT
def get_movements(limit=100):
    # Используем конкатенацию для формирования ФИО и приведение даты к типу DATE
    query = """
        SELECT m.id_перемещения, 
               t.инвентарный_номер, t.наименование,
               e.фамилия || ' ' || e.имя as сотрудник,
               m.дата::date as дата
        FROM Перемещение m
        JOIN Техника t ON m.id_техники = t.id_техники
        JOIN Сотрудники e ON m.id_сотрудника = e.id_сотрудника
        ORDER BY m.дата DESC
        LIMIT %s
    """
    return execute_query(query, (limit,), fetch=True)


# Добавление записи о перемещении
def add_movement(equipment_id, employee_id, date_val):
    query = """
        INSERT INTO Перемещение (id_техники, id_сотрудника, дата)
        VALUES (%s, %s, %s)
    """
    return execute_query(query, (equipment_id, employee_id, date_val))


# Получение данных одного перемещения по его ID для сохранения акта
def get_movement_by_id(movement_id):
    query = """
        SELECT m.id_перемещения, 
               t.инвентарный_номер, t.наименование, t.модель, t.серийный_номер,
               e.фамилия, e.имя, e.отчество, e.должность, e.отдел,
               m.дата::date as дата
        FROM Перемещение m
        JOIN Техника t ON m.id_техники = t.id_техники
        JOIN Сотрудники e ON m.id_сотрудника = e.id_сотрудника
        WHERE m.id_перемещения = %s
    """
    result = execute_query(query, (movement_id,), fetch=True)
    return result[0] if result else None
# ==================== РАБОТА С РЕМОНТАМИ ====================

# Получение списка ремонтов с возможной фильтрацией по периоду дат заявок
def get_repairs(period_start=None, period_end=None):
    # Базовый запрос с JOIN для получения инвентарного номера и наименования техники
    query = """
        SELECT r.id_ремонта, t.инвентарный_номер, t.наименование,
               r.дата_заявки, r.дата_ремонта, r.описание, 
               r.стоимость, r.статус
        FROM Ремонт r
        JOIN Техника t ON r.id_техники = t.id_техники
        WHERE 1=1
    """
    params = []
    if period_start:
        query += " AND r.дата_заявки >= %s"
        params.append(period_start)
    if period_end:
        query += " AND r.дата_заявки <= %s"
        params.append(period_end)
    query += " ORDER BY r.дата_заявки DESC"
    return execute_query(query, params, fetch=True)


# Добавление записи о ремонте
def add_repair(equipment_id, request_date, repair_date, description, cost, status):
    query = """
        INSERT INTO Ремонт (id_техники, дата_заявки, дата_ремонта, 
                           описание, стоимость, статус)
        VALUES (%s, %s, %s, %s, %s, %s)
    """
    return execute_query(query, (equipment_id, request_date, repair_date, description, cost, status))


# Подсчёт общего количества ремонтов
def get_repairs_count():
    query = "SELECT COUNT(*) FROM Ремонт"
    result = execute_query(query, fetch=True)
    return result[0][0] if result else 0


# ==================== ПОЛУЧЕНИЕ СПИСКА СТАТУСОВ ====================

def get_statuses():
    # Возвращает список статусов: (id_статуса, наименование)
    query = "SELECT id_статуса, наименование FROM Статус ORDER BY id_статуса"
    return execute_query(query, fetch=True)


# ==================== АУТЕНТИФИКАЦИЯ ====================

# Проверка учётных данных пользователя (логин + пароль)
def authenticate_user(login, password):
    # Запрос: выбираем данные пользователя и связанного сотрудника
    query = """
        SELECT u.id_пользователя, u.роль, u.id_сотрудника, 
               e.фамилия, e.имя, u.пароль
        FROM Пользователи u
        JOIN Сотрудники e ON u.id_сотрудника = e.id_сотрудника
        WHERE u.логин = %s
    """
    result = execute_query(query, (login,), fetch=True)
    if not result:
        return None
    user = result[0]
    stored_hash = user[5]                # Хранимый хэш пароля
    if verify_password(password, stored_hash):
        # Возвращаем кортеж: id_пользователя, роль, id_сотрудника, фамилия, имя
        return (user[0], user[1], user[2], user[3], user[4])
    else:
        return None


# ==================== УПРАВЛЕНИЕ ПОЛЬЗОВАТЕЛЯМИ ====================

# Получить список всех зарегистрированных пользователей с их ролями и ФИО сотрудников
def get_all_users():
    query = """
        SELECT u.id_пользователя, u.логин, u.роль, 
               e.фамилия || ' ' || e.имя as сотрудник
        FROM Пользователи u
        JOIN Сотрудники e ON u.id_сотрудника = e.id_сотрудника
    """
    return execute_query(query, fetch=True)


# Добавление нового пользователя (с проверкой логина, сложности пароля и хэшированием)
def add_user(username, password, role, employee_id):
    # Проверка уникальности логина
    check_query = "SELECT id_пользователя FROM Пользователи WHERE логин = %s"
    existing = execute_query(check_query, (username,), fetch=True)
    if existing:
        return False, "Пользователь с таким логином уже существует"

    # Проверка сложности пароля
    is_valid, error_msg = validate_password_strength(password)
    if not is_valid:
        return False, error_msg

    # Хэширование пароля
    password_hash = hash_password(password)

    # Вставка новой записи
    query = """
        INSERT INTO Пользователи (логин, пароль, роль, id_сотрудника)
        VALUES (%s, %s, %s, %s)
    """
    result = execute_query(query, (username, password_hash, role, employee_id))
    if result:
        return True, "Пользователь успешно добавлен"
    return False, "Ошибка при создании пользователя"

# Получение данных пользователя по ID
def get_user_by_id(user_id):
    query = """
        SELECT u.id_пользователя, u.логин, u.роль, u.id_сотрудника,
               e.фамилия, e.имя, e.отчество, e.должность, e.отдел
        FROM Пользователи u
        JOIN Сотрудники e ON u.id_сотрудника = e.id_сотрудника
        WHERE u.id_пользователя = %s
    """
    result = execute_query(query, (user_id,), fetch=True)
    return result[0] if result else None

# Обновление данных пользователя (логин, пароль, роль, сотрудник)
def update_user(user_id, login, password, role, employee_id):
    # Проверка сложности, если передан новый пароль
    if password:
        is_valid, msg = validate_password_strength(password)
        if not is_valid:
            return False, msg
        password_hash = hash_password(password)
        query = """
            UPDATE Пользователи
            SET логин = %s, пароль = %s, роль = %s, id_сотрудника = %s
            WHERE id_пользователя = %s
        """
        result = execute_query(query, (login, password_hash, role, employee_id, user_id))
    else:
        # Пароль не меняем
        query = """
            UPDATE Пользователи
            SET логин = %s, роль = %s, id_сотрудника = %s
            WHERE id_пользователя = %s
        """
        result = execute_query(query, (login, role, employee_id, user_id))
    return (True, "Пользователь обновлён") if result else (False, "Ошибка обновления пользователя")

# Удаление пользователя по ID
def delete_user(user_id):
    query = "DELETE FROM Пользователи WHERE id_пользователя = %s"
    result = execute_query(query, (user_id,))
    return result > 0


# ==================== ОТЧЕТЫ ====================

# Отчет о закреплении техники за сотрудниками (сколько единиц у каждого, актуальное состояние)
def get_report_equipment_by_employee():
    # Подзапрос: для каждой единицы техники берём последнее перемещение (дата DESC)
    query = """
        SELECT e.фамилия || ' ' || e.имя as сотрудник, 
               e.должность,
               COUNT(m.id_техники) as количество_техники
        FROM Сотрудники e
        LEFT JOIN (
            SELECT DISTINCT ON (id_техники) id_техники, id_сотрудника
            FROM Перемещение
            ORDER BY id_техники, дата DESC
        ) m ON e.id_сотрудника = m.id_сотрудника
        GROUP BY e.id_сотрудника, e.фамилия, e.имя, e.должность
        ORDER BY количество_техники DESC
    """
    return execute_query(query, fetch=True)