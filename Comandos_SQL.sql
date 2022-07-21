--**Comandos SQL focado em RPA**
---------------------------------------------------------------------------------------------------------------------
-- Comando para aplicar o SELECT em uma planilha Excel:
Provider=Microsoft.ACE.OLEDB.12.0;Data Source='String caminho Excel';Extended Properties="Excel 12.0 Xml;HDR=Yes";

-- Mostra todas as tabelas presentes no banco de dados:
SELECT * FROM SYSOBJECTS

-- Mostra todas as colunas presentes na tabela:
SELECT * FROM [NOME DA TABELA]

-- Mostra a coluna especificada presente na tabela:
SELECT [NOME DA COLUNA] FROM [NOME DA TABELA]

-- Apaga os dados presentes na tabela especificada:
DELETE FROM [NOME DA TABELA]

-- Apaga os dados da coluna especificada presentes na tabela:
DELETE FROM [NOME DA TABELA] WHERE [NOME DA COLUNA]

-- Traz uma procedures especificada:
EXEC sp_helptext [NOME DA PROCEDURES]

-- Seleciona o tanto de registros (linhas) que quiser trazer na consulta:
SELECT TOP (QUANTIDADE DE LINHAS QUE QUR TRAZE) [NOME DA TABELA]

-- Faz a inserção dos dados presentes na planilha Excel para as linhas das colunas da tabela:
INSERT INTO [NOME DA TABELA] 
([NOME DA COLUNA], [NOME DA COLUNA], [NOME DA COLUNA])
VALUES ('INSERIR VARIAVEL', 'INSERIR VARIAVEL', 'INSERIR VARIAVEL')

-- Junta 2 tabelas:
SELECT * FROM [NOME DA TABELA] 'INSERE UM NOME PARA TABELA 1 EX: a'
INNER JOIN [NOME DA TABELA] 'INSERE UM NOME PARA TABELA 2 EX: b'
ON a.'INSERIR COLUNA DE COMUNICAÇÃO' = b.'INSERIR COLUNA DE COMUNICAÇÃO'

-- Realiza o update utilizando alguma condição:
UPDATE [NOME DA TABELA]
SET [NOME DA COLUNA A SER REALIZADO O UPDATE] = 'VALOR A SER INSERIDO'
WHERE [NOME DA COLUNA PARAMETRO] LIKE 'VALOR PARAMETRO'

-- Tratamento para inserir NULL em células em branco:
CASE
WHEN 'INSERIR VARIAVEL' = ' ' THEN NULL
ELSE 'INSERIR VARIAVEL'
END

-- Tratamento que ignora erros apresentados:
BEGIN TRY
'INSERIR QUARY'
END TRY
BEGIN
END CATCH

-- Observações:
% = Coringa (ignora tudo o que vem na sequência ou anteriormente)
# = Utilizado para ignorar ponto final (.) quando existir no cabeçalho
null = Adiciona valor nulo (Em branco) ao campo
