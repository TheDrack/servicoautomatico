SnakeIA - Ensinando uma Rede Neural a Jogar
Este projeto foi o meu "mergulho de cabeça" em Inteligência Artificial. Em vez de usar regras fixas (IF/ELSE), eu construí um agente de Deep Reinforcement Learning (Aprendizado por Reforço) que aprende a jogar Snake por tentativa e erro.
A Ideia do Projeto
Eu queria entender como uma rede neural toma decisões em tempo real. Usei o algoritmo DQN (Deep Q-Network), o mesmo conceito que a DeepMind usou para vencer jogos de Atari. A cobra não sabe o que é uma maçã ou uma parede; ela recebe pontos (recompensas) e ajusta os pesos da sua rede neural para maximizar essa pontuação.
O que eu apliquei de Engenharia:
Pipes de Dados Vetorizados: Refatorei o treino para usar o poder do TensorFlow e NumPy, processando "lotes" (batches) de 64 experiências de uma vez. Isso deixa o aprendizado muito mais rápido do que treinar frame por frame.
Memória de Experiência (Replay Buffer): Implementei uma fila circular que armazena as últimas 20 mil jogadas. A IA "estuda" esses momentos aleatoriamente para não ficar viciada apenas nas últimas ações.
Double-Network Strategy: Usei duas redes neurais (Policy e Target). Uma dita o que fazer, a outra serve de referência para o cálculo do erro. Isso evita que o aprendizado fique instável.
Como a IA "Enxerga":
Para manter o modelo leve e eficiente, não usei visão computacional pesada. Passei para a rede apenas a distância relativa entre a cabeça da cobra e a comida (dx, dy). É como se a cobra tivesse uma bússola interna que ela aprende a seguir conforme ganha recompensas.
Stack Técnica:
Python
TensorFlow / Keras (Cérebro da IA)
Pygame (Interface visual do jogo)
NumPy (Cálculos de matrizes e vetores)