from dotenv import dotenv_values
import psycopg2

config = dotenv_values(".env")

def select_from_datauser(value: str):
    with psycopg2.connect(dbname = config["POSTGRES_DB"], user = config["POSTGRES_USER"], password = config["POSTGRES_PASSWORD"], host = "localhost", port = "5432") as con:
        with con.cursor() as cur:
            cur.execute('select * from workers where username = %s', (value,))
            result = cur.fetchall()
    return result
    
def check_permissions(telegram_ID: int) -> bool:
    with psycopg2.connect(dbname = config["POSTGRES_DB"], user = config["POSTGRES_USER"], password = config["POSTGRES_PASSWORD"], host = "localhost", port = "5432") as con:
        with con.cursor() as cur:
            cur.execute('select * from users where telegram_id = %s', (telegram_ID,))
            result = cur.fetchone()
    return bool(result)

def get_fio_from_user(telegram_ID: int) -> bool:
    with psycopg2.connect(dbname = config["POSTGRES_DB"], user = config["POSTGRES_USER"], password = config["POSTGRES_PASSWORD"], host = "localhost", port = "5432") as con:
        with con.cursor() as cur:
            cur.execute('select current_fio from users where telegram_id = %s', (telegram_ID,))
            result = cur.fetchone()
    return result

def check_users_by_fio(value: str):
    with psycopg2.connect(dbname = config["POSTGRES_DB"], user = config["POSTGRES_USER"], password = config["POSTGRES_PASSWORD"], host = "localhost", port = "5432") as con:
        with con.cursor() as cur:
            cur.execute('select * from workers where username = %s', (value,))
            result = cur.fetchall()
    return bool(result)

def add_fio(value: str, telegram_id):
    with psycopg2.connect(dbname = config["POSTGRES_DB"], user = config["POSTGRES_USER"], password = config["POSTGRES_PASSWORD"], host = "localhost", port = "5432") as con:
        with con.cursor() as cur:
            cur.execute('update users set current_fio = %s where telegram_id = %s', (value, telegram_id))
            con.commit()
