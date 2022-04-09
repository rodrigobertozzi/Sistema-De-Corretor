### Sistema-De-Corretor

### Configurações da Base de Dados

1- Criar Base no localhost
    cn.Provider = "SQLOLEDB"
    cn.Properties("Data Source").Value = "localhost"
    cn.Properties("Initial Catalog").Value = "SistemaCorretor"
    cn.Properties("User ID").Value = "sa"
    cn.Properties("Password").Value = "1q2w3e4r@#$"
2- Executar a query CriacaoDataBaseETabelas.sql
3- Executar a query Adicao_UF.sql
4- Executar a query cidades_e_estados.sql

### Iniciar Projeto

Baixar o Executavel Sistema-De-Corretores.exe

### Dificuldades

1- Não consegui deletar o registro quando sai da grid
2- Não consegui filtrar os resultados, apenas todos os registros
