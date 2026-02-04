# 🤖 Serviço Automático — Portfólio Técnico

Este repositório reúne projetos reais, experimentais e evolutivos que desenvolvi ao longo da minha trajetória com **automação de processos, engenharia de sistemas, integrações assíncronas e inteligência artificial aplicada**.

Mais do que projetos isolados, este repositório representa uma **linha de evolução técnica**: da automação simples à orquestração de robôs, migração de dados em alto volume e exploração prática de IA.

> **Nota:** O foco aqui não é código de tutorial, mas soluções pensadas para problemas reais, limitações operacionais e cenários de produção.

---

## 🧠 Áreas de Atuação Exploradas

* **Automação de Processos:** RPA e scripts inteligentes.
* **Orquestração:** Gerenciamento centralizado de robôs.
* **Data Engineering:** Processamento e migração de dados em grande escala.
* **Sistemas Modernos:** Arquiteturas assíncronas e concorrentes.
* **Real-Time:** Integrações via WebSocket e mensageria.
* **IA Aplicada:** Reinforcement Learning (Aprendizado por Reforço).

---

## 📂 Estrutura do Repositório

### `00_Protótipos_Historicos`
Registro da evolução técnica e arquitetural das ideias. Demonstra o aprendizado incremental e como decisões arquiteturais foram refinadas ao longo do tempo antes da consolidação das soluções atuais.

### `02_Gerenciador_Robos_Pro`
Orquestrador avançado com interface gráfica (GUI) e suporte a multithreading para execução simultânea de agentes.

### `03_High-Volume-Data-Migration`
Pipeline focado em performance para movimentação massiva de dados com tratamento de integridade.

### `04_SnakeIA`
Estudo prático de **Reinforcement Learning** utilizando Deep Q-Networks para treinamento de agentes autônomos.

---

## 🚀 Projetos em Destaque

### Real-Time WhatsApp Bridge
Arquitetura híbrida para integração de mensagens:
* **Selenium & Asyncio:** Automação assíncrona.
* **WebSockets:** Captura de eventos em tempo real.
* **LGPD Compliant:** Tratamento ético e seguro de dados sensíveis.

---

## 🛠️ Stack Tecnológica
* **Linguagem:** Python 3.9+
* **IA & Dados:** TensorFlow, NumPy, Pandas.
* **Automação:** Selenium, Asyncio, Websockets.
* **Interface:** Tkinter / CustomTkinter.
* **DevOps:** Docker, `.env` (gestão de secrets), Logging estruturado.

---

## 🔧 Configuração e Instalação

### 🐳 Rodando com Docker (Recomendado)
Se você tem o Docker instalado, execute:
```bash
docker-compose up --build
```

###🐍 Instalação Manual
​Caso prefira rodar localmente:
​Clone o repositório:
​<!-- end list -->
   git clone [https://github.com/TheDrack/servicoautomatico.git](https://github.com/TheDrack/servicoautomatico.git)

2. **Configure o ambiente:**
   ```bash
   cp .env.example .env
 Edite o .env com suas credenciais

3. **Instale as dependências:**
   ```bash
   pip install -r requirements.txt

---
🎯 Filosofia do Repositório
Este repositório não foi criado para ser um produto final fechado, mas sim para documentar:
Processo de pensamento
Evolução técnica
Decisões arquiteturais
Soluções para problemas reais
Ele reflete um perfil voltado a engenharia, não apenas implementação.
👤 Autor
Projeto desenvolvido por Jesus Anhaia
Perfil focado em automação, engenharia de sistemas e IA aplicada.
“Código é só uma parte do sistema. Entender o problema é o que faz a diferença.”

___

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
