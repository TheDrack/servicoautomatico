Migração de Dados (MDB para MySQL) - Alta Volumetria
Esse código eu criei para resolver um problema real: eu precisava mover milhares de registros de um servidor antigo (Access/MDB) para um MySQL novo. O arquivo era pesado, o sistema era instável e as ferramentas prontas travavam tudo ou estouravam a memória.
O Problema
Se eu tentasse abrir o arquivo inteiro ou exportar para um SQL gigante antes de importar, o PC travava. O arquivo .mdb não aceita desaforo quando fica muito grande.
A Minha Solução
Em vez de fazer o básico, eu usei uma estratégia de Streaming com Pipes.
Sem arquivo temporário: O código não cria um arquivo .sql no disco para depois ler. Ele usa o subprocess.PIPE para "sugar" os dados direto do mdb-export e já ir jogando para dentro do MySQL em tempo real.
Memória Baixa: Como o dado não para no disco e nem é carregado todo de uma vez no Python, o uso de memória RAM é quase zero, não importa se a tabela tem 1 milhão de linhas.
Commit em Lote (Batch): Eu configurei para dar commit a cada 1000 linhas. Se a rede cair ou der algum erro no meio do caminho, o que já foi processado está salvo no banco. É muito mais seguro para migrações que demoram horas.
Tecnologias
Python
MySQL Connector (com controle manual de transação)
mdbtools (roda por trás via subprocess)
tqdm (barra de progresso para não ficar no escuro)
Como eu uso
É um script utilitário. Eu abro o código, mudo o caminho do arquivo e o nome da tabela no cabeçalho e mando rodar.
Precisa ter o mdbtools instalado no Windows ou Linux.