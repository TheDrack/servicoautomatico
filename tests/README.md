# Testes Unitários - Serviço Automático

Este diretório contém testes unitários para todos os arquivos `.py` do repositório, utilizando Pytest.

## Estrutura dos Testes

Cada arquivo de teste corresponde a um módulo Python específico e testa apenas:
- Funções de lógica simples
- Cálculos e manipulação de dados
- Validações e formatações

**Não foram incluídos testes para:**
- Conexões com bancos de dados
- Chamadas a APIs externas
- Bibliotecas complexas (Selenium, Tkinter GUI, etc.)
- Processos assíncronos e threading

## Como Executar os Testes

### Instalar dependências
```bash
pip install pytest numpy
```

### Executar todos os testes
```bash
pytest tests/ -v
```

### Executar um arquivo específico
```bash
pytest tests/test_whatsapp_bridge.py -v
```

## Arquivos de Teste Criados

### 1. tests/test_whatsapp_bridge.py
Testa lógica de anonimização de dados do arquivo `whatsapp_bridge.py`
- `test_anonymize_logic()` - Verifica a lógica de hash SHA256
- `test_anonymize_different_data()` - Testa que dados diferentes geram hashes diferentes
- `test_anonymize_same_data_same_hash()` - Testa consistência do hash

### 2. tests/test_snakeia.py
Testa cálculos e lógica do jogo do arquivo `04_SnakeIA/SnakeIA.py`
- `test_state_calculation()` - Verifica cálculo de distância relativa
- `test_state_same_position()` - Testa estado quando snake e food estão juntos
- `test_movement_calculation()` - Valida movimentos (direita, esquerda, cima, baixo)
- `test_collision_detection()` - Testa detecção de colisão com bordas
- `test_food_detection()` - Verifica quando a cobra come a comida
- `test_reward_calculation()` - Valida valores de recompensas

### 3. tests/test_assistente_almoxarifado.py
Testa funções utilitárias do arquivo `00_Protótipos_Historicos/AssistenteDeAlmoxarifado.py`
- `test_material_code_formatting()` - Testa formatação de código de material
- `test_material_code_string_conversion()` - Verifica conversão numérico para string
- `test_text_uppercase_conversion()` - Testa conversão para maiúsculas
- `test_ac_prefix_formatting()` - Valida formatação do campo A/C
- `test_quantidade_calculation()` - Testa cálculo de ajuste de estoque
- `test_quantidade_calculation_positive()` - Testa ajuste positivo
- `test_tuple_unpacking_simulation()` - Verifica desempacotamento de tuplas
- `test_check_keyword_in_command()` - Testa detecção de palavras-chave

### 4. tests/test_gerenciador_v1.py
Testa lógica do gerenciador do arquivo `01_Gerenciador_Robos_Simples/gerenciador_v1.py`
- `test_robot_id_generation()` - Verifica geração de ID único
- `test_robot_id_uniqueness()` - Testa unicidade de IDs
- `test_exec_count_validation()` - Valida número de execuções
- `test_status_counting()` - Testa contagem de robôs ativos
- `test_robot_name_extraction()` - Verifica extração de nome de arquivo
- `test_tree_values_structure()` - Valida estrutura de dados da TreeView

### 5. tests/test_supervisor.py
Testa lógica do supervisor do arquivo `02_Gerenciador_Robos_Pro/core/supervisor.py`
- `test_job_id_generation()` - Verifica geração de job ID
- `test_job_id_uniqueness()` - Testa unicidade de job IDs
- `test_max_workers_limit()` - Valida limite de workers
- `test_result_queue_message_structure()` - Testa estrutura de mensagens
- `test_job_queue_tuple_structure()` - Verifica estrutura da fila de jobs

### 6. tests/test_main_gui.py
Testa funções utilitárias do arquivo `02_Gerenciador_Robos_Pro/ui/main_gui.py`
- `test_file_name_extraction()` - Verifica extração de nome de arquivo
- `test_file_name_extraction_windows()` - Testa com paths do Windows
- `test_message_type_identification()` - Valida identificação de tipo de mensagem
- `test_log_formatting()` - Testa formatação de logs
- `test_tree_column_structure()` - Verifica estrutura de colunas

### 7. tests/test_run_gui.py
Testa configuração do arquivo `02_Gerenciador_Robos_Pro/run_gui.py`
- `test_max_workers_configuration()` - Verifica configuração de max_workers
- `test_max_workers_different_values()` - Testa diferentes valores

### 8. tests/test_migracao_mdb_mysql.py
Testa validações do arquivo `03_High-Volume-Data-Migration/Migracao-MDB-MySQL.py`
- `test_batch_size_default_value()` - Verifica valor padrão de BATCH_SIZE
- `test_batch_size_custom_value()` - Testa valor customizado
- `test_insert_query_detection()` - Valida detecção de query INSERT
- `test_insert_counter_increment()` - Testa incremento de contadores
- `test_batch_commit_logic()` - Verifica lógica de commit em lote
- `test_batch_counter_reset()` - Testa reset do contador
- `test_config_validation()` - Valida configurações obrigatórias
- `test_query_strip()` - Testa remoção de espaços em branco

## Resultados

✅ **43 testes criados**  
✅ **100% de aprovação**  
✅ **0 dependências complexas** (apenas pytest e numpy)

## Importações

Todos os testes seguem a estrutura real do repositório:
- Não há importações diretas dos módulos principais
- Os testes simulam a lógica para evitar dependências complexas
- Apenas bibliotecas padrão (hashlib, uuid) e numpy são utilizadas
