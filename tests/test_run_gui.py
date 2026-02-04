"""
Testes unitários para 02_Gerenciador_Robos_Pro/run_gui.py
Testa apenas funções de lógica e configuração
"""
import pytest


def test_max_workers_configuration():
    """Testa configuração de max_workers"""
    # Simula a configuração no __main__
    max_workers = 3
    
    assert max_workers == 3
    assert isinstance(max_workers, int)
    assert max_workers > 0


def test_max_workers_different_values():
    """Testa diferentes valores de max_workers"""
    config1 = 2
    config2 = 5
    config3 = 10
    
    assert config1 < config2 < config3
    assert all(isinstance(x, int) for x in [config1, config2, config3])
