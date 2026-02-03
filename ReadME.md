​🤖 Engenharia de Automação & Data Pipeline
​Repositório focado em soluções de automação de alto nível, transitando desde RPA (Robotic Process Automation) tradicional até sistemas complexos de migração de dados e IA.
​📂 Estrutura do Repositório
​O projeto está organizado de forma evolutiva, cobrindo diferentes pilares da automação:
​00. Protótipos Históricos
​Registro da evolução das automações, contendo scripts iniciais e provas de conceito que fundamentaram os sistemas atuais.
​01. Gerenciador de Robôs (Simples)
​Versão inicial de orquestração de scripts, focada em execução sequencial e controle básico de logs.
​02. Gerenciador de Robôs Pro
​Interface gráfica (GUI) desenvolvida em Tkinter com orquestração de threads. Permite gerenciar, monitorar e executar múltiplos processos Python simultaneamente de forma assíncrona.
​03. High-Volume Data Migration
​Pipeline focado em performance para movimentação de grandes volumes de dados. Implementa estratégias de batch processing e tratamento de erros para garantir a integridade da migração.
​04. SnakeIA - Aprendizado por Reforço
​Estudo de Inteligência Artificial aplicada, utilizando TensorFlow e DQN (Deep Q-Network) para treinar um agente autônomo.
​🚀 Destaque: Real-Time WhatsApp Bridge
​Localizado na raiz (whatsapp_bridge.py), este é o motor de integração do repositório. Uma arquitetura híbrida que utiliza:
​Selenium: Para autenticação e bypass de segurança.
​Requests & WebSockets: Para captura e orquestração de metadados em tempo real com asyncio.
​LGPD Compliant: Implementação de pseudonimização e proteção de dados sensíveis.
​🛠️ Stack Tecnológica
​Linguagem: Python 3.9+
​IA & Dados: TensorFlow, NumPy, Pandas.
​Automação & Web: Selenium, Asyncio, Websockets.
​Interface: Tkinter (CustomTkinter).
​Infra: Docker, Environment Variables (.env), Logging Estruturado.