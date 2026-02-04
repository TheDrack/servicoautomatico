import tensorflow as tf
import numpy as np
import random
import pygame
import gym
import sys
from collections import deque
from typing import Tuple, List


# --- CONFIGURAÇÕES E HIPERPARÂMETROS ---
ENV_SIZE = 10
STATE_SHAPE = (2,) # [dist_x, dist_y]
ACTIONS = 4
GAMMA = 0.99
LR = 0.001
BATCH_SIZE = 64
BUFFER_SIZE = 20000
EPSILON_START = 1.0
EPSILON_MIN = 0.1
EPSILON_DECAY = 0.995
TARGET_UPDATE_FREQ = 10 # episódios

# --- ARQUITETURA DA REDE (DQN) ---
def build_dqn(input_shape: Tuple[int, ...], action_space: int) -> tf.keras.Sequential:
    """
    Constrói uma Deep Q-Network (DQN) para aprendizado por reforço.
    
    Args:
        input_shape: Formato do estado de entrada.
        action_space: Número de ações possíveis.
        
    Returns:
        Modelo Keras compilado com otimizador Adam e loss MSE.
    """
    model = tf.keras.Sequential([
        tf.keras.layers.Dense(128, activation='relu', input_shape=input_shape),
        tf.keras.layers.Dense(64, activation='relu'),
        tf.keras.layers.Dense(action_space, activation='linear')
    ])
    model.compile(optimizer=tf.keras.optimizers.Adam(learning_rate=LR), loss='mse')
    return model

# --- AMBIENTE (LÓGICA DO JOGO) ---
class SnakeEnv:
    """
    Ambiente do jogo Snake para treinamento de RL.
    
    Implementa a lógica do jogo com estados normalizados para
    facilitar o aprendizado da rede neural.
    """
    def __init__(self, size: int = 10) -> None:
        self.size = size
        self.reset()

    def reset(self) -> np.ndarray:
        self.snake = [self.size//2, self.size//2]
        self.food = [random.randint(0, self.size-1), random.randint(0, self.size-1)]
        self.done = False
        return self._get_state()

    def _get_state(self) -> np.ndarray:
        # Distância relativa normalizada ajuda a rede a aprender mais rápido
        return np.array([
            self.food[0] - self.snake[0],
            self.food[1] - self.snake[1]
        ], dtype=np.float32)

    def step(self, action: int) -> Tuple[np.ndarray, int, bool]:
        # 0: Dir, 1: Esq, 2: Baixo, 3: Cima
        moves = {0: [0, 1], 1: [0, -1], 2: [1, 0], 3: [-1, 0]}
        move = moves[action]
        
        self.snake = [self.snake[0] + move[0], self.snake[1] + move[1]]

        # Verificação de colisão
        if (not (0 <= self.snake[0] < self.size)) or (not (0 <= self.snake[1] < self.size)):
            return self._get_state(), -10, True # Penalidade maior por morrer

        # Verificação de comida
        if self.snake == self.food:
            self.food = [random.randint(0, self.size-1), random.randint(0, self.size-1)]
            return self._get_state(), 10, False # Recompensa por comer
            
        return self._get_state(), 0, False

# --- AGENTE (INTELIGÊNCIA) ---
class DQNAgent:
    """
    Agente DQN (Deep Q-Network) que aprende a jogar Snake.
    
    Usa experience replay e target network para estabilizar
    o treinamento de aprendizado por reforço.
    """
    def __init__(self) -> None:
        self.memory = deque(maxlen=BUFFER_SIZE)
        self.epsilon = EPSILON_START
        self.model = build_dqn(STATE_SHAPE, ACTIONS)
        self.target_model = build_dqn(STATE_SHAPE, ACTIONS)
        self.update_target()

    def update_target(self) -> None:
        self.target_model.set_weights(self.model.get_weights())

    def act(self, state: np.ndarray) -> int:
        if np.random.rand() <= self.epsilon:
            return random.randrange(ACTIONS)
        act_values = self.model.predict(state.reshape(1, *STATE_SHAPE), verbose=0)
        return np.argmax(act_values[0])

    def train(self) -> None:
        if len(self.memory) < BATCH_SIZE:
            return

        minibatch = random.sample(self.memory, BATCH_SIZE)
        
        # Vetorização Sênior: Processando o batch de uma vez (Numpy magic)
        states = np.array([ex[0] for ex in minibatch])
        actions = np.array([ex[1] for ex in minibatch])
        rewards = np.array([ex[2] for ex in minibatch])
        next_states = np.array([ex[3] for ex in minibatch])
        dones = np.array([ex[4] for ex in minibatch])

        # Predição em lote (Mais rápido)
        targets = self.model.predict(states, verbose=0)
        next_q_values = self.target_model.predict(next_states, verbose=0)

        for i in range(BATCH_SIZE):
            if dones[i]:
                targets[i][actions[i]] = rewards[i]
            else:
                targets[i][actions[i]] = rewards[i] + GAMMA * np.amax(next_q_values[i])

        self.model.fit(states, targets, epochs=1, verbose=0)

        if self.epsilon > EPSILON_MIN:
            self.epsilon *= EPSILON_DECAY

# --- EXECUÇÃO E VISUALIZAÇÃO ---
def main() -> None:
    pygame.init()
    screen = pygame.display.set_mode((400, 400))
    clock = pygame.time.Clock()
    
    env = SnakeEnv(ENV_SIZE)
    agent = DQNAgent()
    episode = 0

    while True:
        state = env.reset()
        total_reward = 0
        
        while True:
            for event in pygame.event.get():
                if event.type == pygame.QUIT: sys.exit()

            action = agent.act(state)
            next_state, reward, done = env.step(action)
            
            agent.memory.append((state, action, reward, next_state, done))
            state = next_state
            total_reward += reward

            # Renderização leve
            screen.fill((0, 0, 0))
            pygame.draw.rect(screen, (255, 0, 0), (env.food[1]*40, env.food[0]*40, 38, 38))
            pygame.draw.rect(screen, (0, 255, 0), (env.snake[1]*40, env.snake[0]*40, 38, 38))
            pygame.display.flip()
            clock.tick(20) # Aumente para treinar mais rápido visualmente

            if done:
                episode += 1
                if episode % TARGET_UPDATE_FREQ == 0:
                    agent.update_target()
                print(f"Episódio: {episode} | Score: {total_reward} | Epsilon: {agent.epsilon:.2f}")
                break
        
        agent.train()

if __name__ == "__main__":
    main()
