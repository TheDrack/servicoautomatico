# 🤖 Engenharia de Automação & Data Pipeline

Repositório focado em soluções de automação de alto nível, transitando desde RPA tradicional até sistemas de migração de dados e IA com aprendizado por reforço.

---

## 📂 Estrutura do Projeto

O repositório está organizado por módulos de complexidade e finalidade:

* **`00_Protótipos_Historicos`**: Legado e evolução das primeiras automações.
* **`01_Gerenciador_Robos_Simples`**: Gerenciamento básico de scripts RPA.
* **`02_Gerenciador_Robos_Pro`**: Orquestrador avançado com interface **Tkinter** e multithreading.
* **`03_High-Volume-Data-Migration`**: Pipelines de ETL e migração massiva de dados com foco em performance.
* **`04_SnakeIA`**: Agente de IA treinado via Deep Q-Learning (**TensorFlow**).
* **`whatsapp_bridge.py`**: Bridge de comunicação em tempo real via WebSockets e Selenium.

---

## 🚀 Tecnologias em Destaque

### Real-Time WhatsApp Bridge
Arquitetura híbrida para integração de mensagens:
- **Selenium & Asyncio**: Automação assíncrona.
- **WebSockets**: Captura de eventos em tempo real.
- **LGPD Compliant**: Tratamento ético e seguro de dados.

### SnakeIA (RL)
Estudo de **Reinforcement Learning**:
- Implementação de rede neural para tomada de decisão autônoma.
- Framework: TensorFlow.

---

## 🛠️ Stack Tecnológica
- **Linguagem:** Python 3.9+
- **Bibliotecas:** TensorFlow, Selenium, Asyncio, Websockets, Pandas.
- **Interface:** Tkinter / CustomTkinter.
- **DevOps:** Docker, `.env` (gestão de secrets), Logging estruturado.

---

## 🔧 Configuração e Instalação

1. **Clone o repositório:**
   ```bash
   git clone [https://github.com/seu-usuario/seu-repo.git](https://github.com/seu-usuario/seu-repo.git)

2. **Configure o ambiente:**
   ```bash
   cp .env.example .env
# Edite o .env com suas credenciais

3. **Instale as dependências:**
   ```bash
   pip install -r requirements.txt

---

### 2. `requirements.txt` (Consolidado)
Este arquivo deve conter todas as bibliotecas necessárias para rodar tanto a IA quanto o Bridge e a Interface. 

```text
# Automação e Web
selenium
websockets
requests
asyncio

# Inteligência Artificial e Dados
tensorflow>=2.10.0
numpy
pandas

# Interface e Sistema
customtkinter
python-dotenv
mysql-connector-python

# Utilitários
logging