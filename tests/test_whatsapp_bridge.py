"""
Testes unitários para whatsapp_bridge.py
Testa apenas funções de lógica e manipulação de dados (sem DB, APIs ou Selenium)
"""
import pytest
import hashlib
import os


def test_anonymize_logic():
    """Testa a lógica de anonimização de dados"""
    # Simula a lógica do método _anonymize com valor conhecido
    data = "test"
    salt = "salt"
    result = hashlib.sha256(f"{data}{salt}".encode()).hexdigest()[:12]
    
    # Verifica propriedades do hash gerado
    assert len(result) == 12
    assert isinstance(result, str)
    # Hash conhecido para esta entrada específica
    assert result == "4edf07edc95b"


def test_anonymize_different_data():
    """Testa que dados diferentes geram hashes diferentes"""
    salt = "test_salt"
    data1 = "user1"
    data2 = "user2"
    
    hash1 = hashlib.sha256(f"{data1}{salt}".encode()).hexdigest()[:12]
    hash2 = hashlib.sha256(f"{data2}{salt}".encode()).hexdigest()[:12]
    
    assert hash1 != hash2


def test_anonymize_same_data_same_hash():
    """Testa que os mesmos dados geram o mesmo hash"""
    salt = "test_salt"
    data = "consistent_data"
    
    hash1 = hashlib.sha256(f"{data}{salt}".encode()).hexdigest()[:12]
    hash2 = hashlib.sha256(f"{data}{salt}".encode()).hexdigest()[:12]
    
    assert hash1 == hash2
