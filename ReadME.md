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

### 📁 Projetos Principais

```text
📂 Projetos
├── 01_Gerenciador_Robos_Simples → RPA básico, execução sequencial
├── 02_Gerenciador_Robos_Pro → Orquestração, GUI, concorrência
├── 03_High-Volume-Data-Migration → ETL e migração massiva
├── 04_SnakeIA → Reinforcement Learning
```
Cada projeto possui seu próprio README.md, contendo explicações técnicas, arquitetura e instruções específicas de execução.
---

## 🚀 Projeto em Destaque

### Real-Time WhatsApp Bridge
Arquitetura híbrida para integração e captura de mensagens em tempo real.

**Características:**
- **Selenium & Asyncio:** Automação assíncrona e controle de fluxo.
- **WebSockets:** Captura e propagação de eventos em tempo real.
- **LGPD Compliant:** Tratamento ético e seguro de dados sensíveis.

Este projeto destaca-se pela complexidade arquitetural e pelos desafios reais envolvidos em integrações sensíveis e orientadas a eventos.

---

## 🛠️ Stack Tecnológica

- **Linguagem:** Python 3.9+
- **Automação & Sistemas:** Selenium, Asyncio
- **Arquiteturas & Concorrência:** WebSockets, processamento assíncrono
- **IA & Dados:** TensorFlow, NumPy, Pandas
- **Interface:** Tkinter / CustomTkinter
- **DevOps & Infra:** Docker, `.env` (gestão de secrets), logging estruturado

---

## 🔧 Configuração e Instalação

### 🐳 Rodando com Docker (Recomendado)

Com o Docker instalado, execute:

```bash
docker-compose up --build

### 🐍 Instalação Manual
​Caso prefira rodar localmente:
1. **​Clone o repositório:**
   ```bash
   git clone [https://github.com/TheDrack/servicoautomatico.git](https://github.com/TheDrack/servicoautomatico.git)


2. **Configure o ambiente:**
   ```bash
   cp .env.example .env
 Edite o .env com suas credenciais

3. **Instale as dependências:**
   ```bash
   pip install -r requirements.txt

---
---

## 🎯 Filosofia do Repositório

Este repositório não foi criado para ser um produto final fechado, mas para documentar:

- Processo de pensamento  
- Evolução técnica  
- Decisões arquiteturais  
- Soluções para problemas reais  

Ele reflete um perfil voltado à **engenharia de sistemas**, não apenas à implementação de código.

---

## 👤 Autor

Projeto desenvolvido por **Jesus Anhaia**  
Perfil focado em automação, engenharia de sistemas e IA aplicada.

> *“Código é só uma parte do sistema. Entender o problema é o que faz a diferença.”*