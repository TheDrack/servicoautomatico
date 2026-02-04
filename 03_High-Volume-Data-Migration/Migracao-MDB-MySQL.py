import os
import subprocess
import logging
import mysql.connector
from tqdm import tqdm
from contextlib import closing
from dotenv import load_dotenv


# ================= ENV =================
load_dotenv()

MDB_PATH = os.getenv("MDB_PATH")
TABLE_NAME = os.getenv("TABLE_NAME")
BATCH_SIZE = int(os.getenv("BATCH_SIZE", 1000))

DB_CONFIG = {
    "host": os.getenv("DB_HOST"),
    "user": os.getenv("DB_USER"),
    "password": os.getenv("DB_PASS"),
    "database": os.getenv("DB_NAME"),
    "autocommit": False
}


# ================= LOG =================
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)


# ================= MIGRATION =================
def run_migration():
    if not MDB_PATH or not TABLE_NAME:
        logging.critical("MDB_PATH ou TABLE_NAME não configurados no .env")
        return

    logging.info(
        f"Migração iniciada | Tabela: '{TABLE_NAME}' | Batch: {BATCH_SIZE}"
    )

    try:
        conn = mysql.connector.connect(**DB_CONFIG)
    except mysql.connector.Error as e:
        logging.error(f"Falha na conexão MySQL: {e}")
        return

    insert_count = 0
    batch_counter = 0

    with closing(conn), closing(conn.cursor()) as cursor:

        process = subprocess.Popen(
            ['mdb-export', '-I', 'mysql', MDB_PATH, TABLE_NAME],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            bufsize=1
        )

        try:
            for line in tqdm(
                process.stdout,
                desc="Pipeline MDB → MySQL",
                unit=" rows"
            ):
                query = line.strip()

                if not query.startswith("INSERT INTO"):
                    continue

                try:
                    cursor.execute(query)
                    insert_count += 1
                    batch_counter += 1

                    if batch_counter >= BATCH_SIZE:
                        conn.commit()
                        batch_counter = 0

                except mysql.connector.Error as e:
                    logging.warning(f"Registro ignorado (erro SQL): {e}")
                    conn.rollback()
                    batch_counter = 0

            conn.commit()

            stderr = process.stderr.read()
            if process.returncode not in (0, None):
                logging.error(f"Erro no mdb-export: {stderr.strip()}")

            logging.info(
                f"Migração concluída | Registros inseridos: {insert_count}"
            )

        except Exception as e:
            logging.critical(f"Erro crítico no pipeline: {e}")
            conn.rollback()

        finally:
            process.stdout.close()
            process.stderr.close()
            process.wait(timeout=10)


# ================= ENTRYPOINT =================
if __name__ == "__main__":
    run_migration()