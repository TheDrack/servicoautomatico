# 🤖 Gerenciador de Robôs Pro 2.0 (Arquitetura Desacoplada)

Este projeto representa a evolução de ferramentas de automação monolíticas para uma arquitetura modular, escalável e preparada para integração com sistemas de IA e Assistentes de Voz.

## 🏗️ Diferenciais da Arquitetura

Diferente da versão 1.0, este gerenciador separa a **Lógica de Controle (Core)** da **Interface de Usuário (UI)**, utilizando o padrão de design focado em robustez:

* **Motor (Supervisor):** Uma classe agnóstica que gerencia filas de execução, threads e controle de concorrência. Ele pode ser operado via GUI, Terminal ou por comandos de voz.
* **Isolamento de Processos:** Utiliza `subprocess` com pipes de comunicação (`stdout`) capturados em tempo real, garantindo que falhas em robôs externos não afetem a estabilidade do sistema principal.
* **Comunicação Thread-Safe:** Implementação de `queue.Queue` para garantir que a interface gráfica nunca trave durante a recepção de logs massivos.
* **Escalabilidade:** Limite configurável de *Workers* simultâneos para preservação de hardware (CPU/RAM).

## 📂 Organização do Projeto

* `/core`: O "Cérebro". Contém o Supervisor e a lógica de execução.
* `/ui`: O "Corpo". Interface gráfica moderna desenvolvida em Tkinter.
* `/robots`: Scripts de exemplo que podem ser carregados pelo sistema.
* `run_gui.py`: Ponto de entrada principal para uso com interface.

## 🚀 Próximos Passos
Esta estrutura foi desenhada para servir como o "Sistema Motor" de um **Assistente Virtual Cognitivo**, onde o assistente poderá despachar tarefas para o Supervisor via API interna.
