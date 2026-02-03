
import os
import json
import asyncio
import datetime
import hashlib
import logging
import requests
import websockets
import mysql.connector
from dotenv import load_dotenv
from contextlib import contextmanager
from mysql.connector import pooling

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


# ================= CONFIGURAÇÃO GLOBAL =================
load_dotenv()

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)

REQUEST_TIMEOUT = 15
WS_TIMEOUT = 20
RECONNECT_DELAY = 60


# ================= CLASSE PRINCIPAL =================
class WhatsAppBridge:

    def __init__(self):
        self.token: str | None = None

        self.db_pool = pooling.MySQLConnectionPool(
            pool_name="wa_pool",
            pool_size=5,
            host=os.getenv("DB_HOST"),
            user=os.getenv("DB_USER"),
            password=os.getenv("DB_PASS"),
            database=os.getenv("DB_NAME")
        )

    # ================= LGPD =================
    def _anonymize(self, data: str) -> str:
        salt = os.getenv("LGPD_SALT", "default_salt")
        return hashlib.sha256(f"{data}{salt}".encode()).hexdigest()[:12]

    @contextmanager
    def _get_db(self):
        conn = self.db_pool.get_connection()
        try:
            yield conn
        finally:
            conn.close()

    # ================= SELENIUM =================
    def get_auth_token(self) -> str:
        logging.info("Autenticando via Selenium...")

        options = Options()
        options.add_argument("--headless=new")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")

        driver = webdriver.Chrome(options=options)

        try:
            driver.get("https://app.botconversa.com.br/login")
            wait = WebDriverWait(driver, 30)

            wait.until(EC.presence_of_element_located((By.NAME, "email"))) \
                .send_keys(os.getenv("BOT_EMAIL"))

            driver.find_element(By.NAME, "password") \
                .send_keys(os.getenv("BOT_PASS"))

            driver.find_element(By.CSS_SELECTOR, "button[type='submit']").click()

            wait.until(lambda d: d.execute_script(
                "return localStorage.getItem('access_token')"
            ))

            token = driver.execute_script(
                "return localStorage.getItem('access_token')"
            )

            if not token:
                raise RuntimeError("Token não encontrado")

            return token

        finally:
            driver.quit()

    # ================= REQUESTS =================
    def fetch_companies(self) -> list[dict]:
        logging.info("Sincronizando companhias...")

        headers = {"Authorization": f"Bearer {self.token}"}
        url = (
            "https://newbackend.botconversa.com.br/api/v1/companies/"
            f"?fr_code={os.getenv('FRANCHISE_CODE')}"
        )

        r = requests.get(url, headers=headers, timeout=REQUEST_TIMEOUT)
        r.raise_for_status()
        return r.json()

    # ================= WEBSOCKET =================
    async def listen_socket(self, comp: dict):
        base_ws = comp["server_url"].replace("http", "ws")
        uri = (
            f"{base_ws}/socket.io/"
            f"?bot={comp['bot']}"
            f"&jwt={self.token}"
            f"&connection_id={comp['server_number']}"
            f"&instance_id={comp['instance_id']}"
            f"&EIO=4&transport=websocket"
        )

        anon_id = self._anonymize(comp["instance_id"])

        while True:
            try:
                async with websockets.connect(uri, ping_interval=20, timeout=WS_TIMEOUT) as ws:
                    logging.info(f"Canal ativo | {comp['name']} | {anon_id}")

                    async for msg in ws:
                        await self._route_message(ws, msg)

            except Exception as e:
                logging.warning(
                    f"Socket encerrado | {comp['name']} | {anon_id} | {self._anonymize(str(e))}"
                )
                await asyncio.sleep(RECONNECT_DELAY)

    async def _route_message(self, ws, raw_msg: str):
        # Heartbeat
        if raw_msg == "2":
            await ws.send("3")
            return

        # Handshake
        if raw_msg.startswith("0"):
            await ws.send("40/notifications")
            await ws.send("40/live_chat")
            return

        if '42/live_chat,["new-mensagem"' not in raw_msg:
            return

        try:
            payload = json.loads(raw_msg[raw_msg.find("["):])[1]
            self._persist_data(payload)
        except Exception:
            pass  # payload malformado → ignora silenciosamente

    # ================= BANCO =================
    def _persist_data(self, payload: dict):
        chat = payload.get("chat") or {}
        msg = payload.get("mensagem") or {}

        content = msg.get("mensagem") or {}
        text = (
            content.get("conversation")
            or content.get("extendedTextmensagem", {}).get("text")
            or "[Mídia/Outros]"
        )

        with self._get_db() as conn:
            with conn.cursor() as cursor:
                cursor.execute(
                    """
                    INSERT INTO conversas
                    (nome_consultor, nome_cliente, telefone_cliente, mensagem, data_mensagem)
                    VALUES (%s, %s, %s, %s, %s)
                    """,
                    (
                        chat.get("isBusyByName"),
                        chat.get("name"),
                        (chat.get("id") or "").split("@")[0],
                        text,
                        datetime.datetime.utcnow()
                    )
                )
                conn.commit()

    # ================= LOOP PRINCIPAL =================
    async def run(self):
        while True:
            try:
                self.token = self.get_auth_token()
                companies = self.fetch_companies()

                tasks = [
                    self.listen_socket(c)
                    for c in companies
                    if c.get("server_url")
                ]

                if tasks:
                    await asyncio.gather(*tasks)

            except Exception as e:
                logging.error(f"Erro crítico no ciclo principal: {e}")
                await asyncio.sleep(RECONNECT_DELAY)


# ================= ENTRYPOINT =================
if __name__ == "__main__":
    bridge = WhatsAppBridge()
    try:
        asyncio.run(bridge.run())
    except KeyboardInterrupt:
        logging.info("Bridge encerrada manualmente.")