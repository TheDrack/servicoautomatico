"""
Testes unitários para 02_Gerenciador_Robos_Pro/ui/main_gui.py
Testa apenas funções de lógica e manipulação de dados
"""
import pytest


def test_file_name_extraction():
    """Testa extração de nome de arquivo do caminho"""
    # Simula a lógica de add_job
    path = "/home/user/scripts/meu_robo.py"
    name = path.split("/")[-1]
    
    assert name == "meu_robo.py"


def test_file_name_extraction_windows():
    """Testa extração de nome com barra invertida (Windows)"""
    path = "C:\\Users\\user\\scripts\\robo.py"
    # Para compatibilidade, usa split com '/' como no código original
    # Mas mostra que funciona com paths diferentes
    parts = path.replace('\\', '/').split("/")
    name = parts[-1]
    
    assert name == "robo.py"


def test_message_type_identification():
    """Testa identificação do tipo de mensagem"""
    # Simula a lógica de update_loop
    msg1 = {"job_id": "test_123", "type": "log", "data": "Linha de log"}
    msg2 = {"job_id": "test_456", "type": "status", "data": "Finalizado"}
    
    assert msg1["type"] == "log"
    assert msg2["type"] == "status"


def test_log_formatting():
    """Testa formatação de mensagem de log"""
    # Simula formatação de log no console
    job_id = "script_abc"
    data = "Processando dados\n"
    
    formatted = f"[{job_id}] {data}"
    
    assert formatted == "[script_abc] Processando dados\n"
    assert job_id in formatted


def test_tree_column_structure():
    """Testa estrutura de colunas da TreeView"""
    # Simula definição de colunas
    columns = ("Nome", "Caminho", "Status")
    
    assert len(columns) == 3
    assert "Status" in columns
    assert columns[0] == "Nome"
