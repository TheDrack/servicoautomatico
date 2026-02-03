import mysql.connector
import subprocess
import os
from tqdm import tqdm

# --- CONFIGURAÇÃO MANUAL ---
MDB_PATH = r'\\192.168.2.214\mis\--BASE DADOS--\Conv_SQL\base.mdb'
TABLE_NAME = 'minha_tabela'

DB_CONFIG = {
    "host": "127.0.0.1",
    "user": "root",
    "password": "",
    "database": "aros_nova",
    "autocommit": False
}

BATCH_SIZE = 1000  # Commit estratégico para performance e segurança

def run_migration():
    print(f"[*] Operação: Migração de Alta Volumetria | Tabela: '{TABLE_NAME}'")

    try:
        conn = mysql.connector.connect(**DB_CONFIG)
        cursor = conn.cursor()
    except Exception as e:
        print(f"[!] Falha na conexão MySQL: {e}"); return

    insert_count = 0
    
    # Popen com PIPE: O dado flui direto do MDB para o MySQL via memória
    process = subprocess.Popen(
        ['mdb-export', '-I', 'mysql', MDB_PATH, TABLE_NAME],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE, # Captura erros do mdb-tools
        text=True,
        bufsize=1
    )

    try:
        # Tqdm consome o generator do stdout
        for line in tqdm(process.stdout, desc="Processando Pipeline", unit=" rows"):
            query = line.strip()
            if not query.startswith("INSERT INTO"):
                continue

            try:
                cursor.execute(query)
                insert_count += 1
                
                if insert_count % BATCH_SIZE == 0:
                    conn.commit()
            except mysql.connector.Error as e:
                print(f"\n[!] Erro de Integridade/Sintaxe: {e}")
                continue

        conn.commit()
        
        # Verifica se o mdb-export reportou erro no stderr
        _, stderr = process.communicate()
        if process.returncode != 0:
            print(f"\n[!] Erro no mdb-export: {stderr}")

        print(f"\n[OK] Finalizado. {insert_count} registros processados.")

    except Exception as e:
        print(f"\n[!] Abortado por erro crítico: {e}")
        conn.rollback()
    finally:
        cursor.close()
        conn.close()

if __name__ == "__main__":
    run_migration()
