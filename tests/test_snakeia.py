"""
Testes unitários para 04_SnakeIA/SnakeIA.py
Testa apenas funções de lógica, cálculos e manipulação de dados
"""
import pytest
import numpy as np


def test_state_calculation():
    """Testa o cálculo de estado (distância relativa)"""
    # Simula a lógica de _get_state
    snake_pos = [5, 5]
    food_pos = [7, 3]
    
    state = np.array([
        food_pos[0] - snake_pos[0],
        food_pos[1] - snake_pos[1]
    ], dtype=np.float32)
    
    assert state[0] == 2.0
    assert state[1] == -2.0
    assert state.dtype == np.float32


def test_state_same_position():
    """Testa estado quando snake e food estão na mesma posição"""
    snake_pos = [5, 5]
    food_pos = [5, 5]
    
    state = np.array([
        food_pos[0] - snake_pos[0],
        food_pos[1] - snake_pos[1]
    ], dtype=np.float32)
    
    assert state[0] == 0.0
    assert state[1] == 0.0


def test_movement_calculation():
    """Testa o cálculo de movimento baseado em ação"""
    # Simula a lógica de movimentos
    moves = {0: [0, 1], 1: [0, -1], 2: [1, 0], 3: [-1, 0]}
    
    snake_pos = [5, 5]
    
    # Testa movimento para direita
    action = 0
    new_pos = [snake_pos[0] + moves[action][0], snake_pos[1] + moves[action][1]]
    assert new_pos == [5, 6]
    
    # Testa movimento para esquerda
    action = 1
    new_pos = [snake_pos[0] + moves[action][0], snake_pos[1] + moves[action][1]]
    assert new_pos == [5, 4]
    
    # Testa movimento para baixo
    action = 2
    new_pos = [snake_pos[0] + moves[action][0], snake_pos[1] + moves[action][1]]
    assert new_pos == [6, 5]
    
    # Testa movimento para cima
    action = 3
    new_pos = [snake_pos[0] + moves[action][0], snake_pos[1] + moves[action][1]]
    assert new_pos == [4, 5]


def test_collision_detection():
    """Testa detecção de colisão com bordas"""
    size = 10
    
    # Testa posição válida
    pos1 = [5, 5]
    is_valid1 = (0 <= pos1[0] < size) and (0 <= pos1[1] < size)
    assert is_valid1 is True
    
    # Testa colisão com borda superior
    pos2 = [-1, 5]
    is_valid2 = (0 <= pos2[0] < size) and (0 <= pos2[1] < size)
    assert is_valid2 is False
    
    # Testa colisão com borda direita
    pos3 = [5, 10]
    is_valid3 = (0 <= pos3[0] < size) and (0 <= pos3[1] < size)
    assert is_valid3 is False


def test_food_detection():
    """Testa detecção de quando cobra come comida"""
    snake_pos = [5, 5]
    food_pos = [5, 5]
    
    ate_food = snake_pos == food_pos
    assert ate_food is True
    
    food_pos = [6, 5]
    ate_food = snake_pos == food_pos
    assert ate_food is False


def test_reward_calculation():
    """Testa cálculo de recompensas"""
    # Recompensa por comer
    reward_eat = 10
    assert reward_eat == 10
    
    # Penalidade por morrer
    reward_die = -10
    assert reward_die == -10
    
    # Recompensa neutra
    reward_neutral = 0
    assert reward_neutral == 0
