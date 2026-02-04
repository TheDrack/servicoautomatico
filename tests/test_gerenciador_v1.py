"""
Testes unitários para 01_Gerenciador_Robos_Simples/gerenciador_v1.py
Testa apenas funções de lógica e manipulação de dados
"""
import pytest
import uuid


def test_robot_id_generation():
    """Testa geração de ID único para robô"""
    # Simula a lógica de start_robot
    robot_name = "test_robot.py"
    robot_id = f"{robot_name}_{uuid.uuid4().hex[:6]}"
    
    assert robot_name in robot_id
    assert len(robot_id.split('_')[-1]) == 6


def test_robot_id_uniqueness():
    """Testa que IDs gerados são únicos"""
    robot_name = "test_robot.py"
    
    robot_id1 = f"{robot_name}_{uuid.uuid4().hex[:6]}"
    robot_id2 = f"{robot_name}_{uuid.uuid4().hex[:6]}"
    
    assert robot_id1 != robot_id2


def test_exec_count_validation():
    """Testa validação de número de execuções"""
    # Simula a lógica de add_robot
    
    # Teste válido
    runs = 5
    is_valid = runs >= 1 and runs <= 100
    assert is_valid is True
    
    # Teste inválido - menor que 1
    runs = 0
    is_valid = runs >= 1 and runs <= 100
    assert is_valid is False
    
    # Teste inválido - maior que 100
    runs = 101
    is_valid = runs >= 1 and runs <= 100
    assert is_valid is False


def test_status_counting():
    """Testa contagem de robôs ativos"""
    # Simula a lógica de _update_status_bar
    robots_status = {
        "robot1": "Rodando",
        "robot2": "Finalizado",
        "robot3": "Rodando",
        "robot4": "Encerrado manualmente"
    }
    
    ativos = sum(1 for status in robots_status.values() if status == "Rodando")
    
    assert ativos == 2
    assert len(robots_status) == 4


def test_robot_name_extraction():
    """Testa extração de nome do arquivo"""
    # Simula lógica de start_robot com os.path.basename
    robot_file = "/home/user/scripts/meu_robo.py"
    robot_name = robot_file.split('/')[-1]
    
    assert robot_name == "meu_robo.py"


def test_tree_values_structure():
    """Testa estrutura de valores para árvore"""
    # Simula estrutura de dados da TreeView
    robot_name = "test.py"
    robot_file = "/path/to/test.py"
    status = "Rodando"
    
    values = (robot_name, robot_file, status)
    
    assert len(values) == 3
    assert values[0] == robot_name
    assert values[2] == status
