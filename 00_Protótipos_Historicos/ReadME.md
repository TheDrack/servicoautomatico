# 🏛️ Protótipo Histórico (Legacy - 2019)

**Status:** Arquivado / Demonstrativo de Evolução Técnica.

Este diretório contém o código de um assistente de automação desenvolvido em **2019**. Ele foi projetado para operar em um sistema legado interno, utilizando técnicas de RPA baseadas em visão computacional e comandos de voz primordiais.

## 🕰️ Contexto do Projeto
Este script representa o "ponto zero" da jornada em automação. Na época, o foco era resolver gargalos operacionais em um sistema de almoxarifado (4R) que não possuía API, exigindo que o robô:
1.  **Enxergasse** a tela (via `PyAutoGUI`).
2.  **Ouvisse** comandos simples (via `SpeechRecognition`).
3.  **Executasse** cliques e digitação humana de forma acelerada.

## 🚀 A Evolução: O Novo Ecossistema
Este protótipo é a base conceitual para o projeto atual, muito mais complexo e robusto, que integra tudo o que foi aprendido em Python ao longo dos anos. O novo sistema (em desenvolvimento nas pastas superiores) eleva o nível da automação para:

* **Agentes de IA:** Tomada de decisão inteligente em vez de scripts lineares.
* **Automação Híbrida:** Integração total entre ambiente **Desktop** e **WebBrowser**.
* **Comando de Voz Avançado:** Processamento de linguagem natural para interações fluidas.
* **Orquestração Pro:** Gerenciamento centralizado de múltiplos robôs com monitoramento em tempo real.

## 🛠️ Tecnologias Utilizadas (2019)
* `Tkinter`: Interface gráfica para operação humana.
* `PyAutoGUI`: Automação de interface (cliques e teclado).
* `SpeechRecognition` & `Pyttsx3`: Interface de voz (STT e TTS).
* `Pandas`: Manipulação de dados de Excel para o sistema legado.

---
> **Nota:** Este código é mantido aqui como registro histórico e prova de conceito da transição de automações simples para arquiteturas complexas de engenharia de software.
