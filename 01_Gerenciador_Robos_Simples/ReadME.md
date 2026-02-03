Orquestrador de Processos (Multi-threading Engine)
​Este projeto é o motor que eu desenvolvi para rodar múltiplos robôs Python simultaneamente sem travar a interface e sem perder o controle do que cada um está fazendo. É o "chassi" que sustenta a execução de qualquer automação que eu crio.
​O que esse motor resolve:
​Paralelismo Real: Eu precisava rodar vários robôs ao mesmo tempo. Aqui, cada robô ganha sua própria Thread e um ID único (UUID), o que permite isolar as execuções.
​Captura de Logs em Tempo Real: Usei subprocess.PIPE para "sequestrar" a saída de cada script. Tudo o que o robô printa aparece instantaneamente no console centralizado.
​Interface Fluida: Para a interface não "congelar" enquanto os robôs trabalham, usei uma fila (queue.Queue) que faz a ponte entre os logs dos robôs e a tela do usuário.
​Diferenciais Técnicos:
​Gestão de Processos (psutil): Se você fechar um robô por aqui, ele não fica "pendurado" na memória. O código rastreia a árvore de processos e encerra até os processos filhos.
​Arquitetura Event-Driven: O consumo de logs não usa um loop infinito pesado; ele usa o sistema de agendamento do Tkinter (.after), o que mantém o uso de CPU baixo.
​Escalabilidade: Posso disparar 1, 5 ou 10 instâncias do mesmo robô apenas mudando um número na tela.
​Como eu uso no dia a dia:
​É a minha ferramenta de "stress test" e monitoramento. Eu seleciono meus scripts de automação, defino quantas instâncias quero e acompanho tudo pelo console verde (estilo terminal). Se algum robô se comportar mal, o encerramento é feito com dois cliques.
​Stack Técnica:
​Python (Core)
​Tkinter (Interface)
​Threading & Subprocess (Orquestração)
​Psutil (Controle de baixo nível do SO)