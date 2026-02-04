# 🤖 # Serviço Automático — Portfólio Técnico

Este repositório reúne projetos reais, experimentais e evolutivos que desenvolvi
ao longo da minha trajetória com **automação de processos, engenharia de sistemas,
integrações assíncronas e inteligência artificial aplicada**.

Mais do que projetos isolados, este repositório representa uma **linha de evolução técnica**:
da automação simples à orquestração de robôs, migração de dados em alto volume
e exploração prática de IA.

O foco aqui **não é código de tutorial**, mas soluções pensadas para
problemas reais, limitações operacionais e cenários de produção.

---

## 🧠 Áreas de Atuação Exploradas

- Automação de processos (RPA, scripts inteligentes)
- Orquestração e gerenciamento de robôs
- Processamento e migração de dados em grande escala
- Arquiteturas assíncronas e concorrentes
- Integrações em tempo real (WebSocket / mensageria)
- Inteligência Artificial aplicada (Reinforcement Learning)
- Design evolutivo de sistemas

---

## 📂 Estrutura do Repositório

### `00_Protótipos_Historicos`
Protótipos iniciais e versões antigas de soluções de automação.

**Objetivo**
- Registrar a evolução técnica e arquitetural das ideias
- Explorar abordagens diferentes antes da consolidação das soluções atuais

**Valor técnico**
- Demonstra aprendizado incremental
- Mostra decisões arquiteturais sendo refinadas ao longo do tempo

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
 Edite o .env com suas credenciais

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