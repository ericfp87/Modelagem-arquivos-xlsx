import pandas as pd
import shutil
import os
import csv

###### CÓPIA DO ARQUIVO EXCEL ######

# Caminho do arquivo original
caminho_arquivo_original = 'C:\\Users\\EricFerreira\\OneDrive - Bulbe\\Power Automate\\Atividades TI\\Atividades TI.xlsm'

# Caminho da pasta de destino
caminho_destino = 'C:\\Users\\EricFerreira\\OneDrive - Bulbe\\Power Automate\\Atividades TI\\BI'

# Extrair apenas o nome do arquivo
nome_arquivo = os.path.basename(caminho_arquivo_original)

# Caminho completo do arquivo de destino
caminho_arquivo_destino = os.path.join(caminho_destino, nome_arquivo)

# Verificar se o arquivo de destino já existe
if os.path.exists(caminho_arquivo_destino):
    # Remover o arquivo de destino existente
    os.remove(caminho_arquivo_destino)

# Copiar o arquivo para a pasta de destino
shutil.copy2(caminho_arquivo_original, caminho_arquivo_destino)


###### GERAÇÃO ARQUIVO CSV ######

# Ler o arquivo Excel
nome_arquivo_excel = 'C:\\Users\\EricFerreira\\OneDrive - Bulbe\\Power Automate\\Atividades TI\\BI\\Atividades TI.xlsm'
nome_aba = 'BI'
colunas = 'A:I'
df = pd.read_excel(nome_arquivo_excel, sheet_name=nome_aba, usecols=colunas)

# Escrever o DataFrame em um arquivo CSV delimitado por ponto e vírgula
atividades_csv = 'C:\\Users\\EricFerreira\\OneDrive - Bulbe\\Power Automate\\Atividades TI\\BI\\teste.csv'
df.to_csv(atividades_csv, sep=';', index=False)



###### CONECTAR AO SQL SERVER ######

import pyodbc

# Detalhes de conexão
server = '192.168.1.249'
database = 'ATIVIDADES_TI'
username = 'sa'
password = 'Bulb3#2023$'

# String de conexão
connection_string = f"DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}"

with open(atividades_csv, 'r', encoding='utf-8') as file:

    # Cria um leitor de CSV para ler o arquivo
    reader = csv.reader(file, delimiter=';')

    # Pula a primeira linha, que contém os cabeçalhos das colunas
    next(reader)

    
    # Estabelecer a conexão
    connection = pyodbc.connect(connection_string)

    
    # Criar um objeto cursor
    cursor = connection.cursor()
    
    for row in reader:
    # Insere a linha na tabela do banco de dados
        query = "INSERT INTO Atividades (Atividade, Setor, Colaborador, Inicio, Status, Conclusao, Tempo, Descricao) VALUES (?, ?, ?, ?, ?, ?, CAST(? AS time), ?)"
        cursor.execute(query, (row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8]))

    # Executar uma consulta SQL
    cursor.execute("SELECT * FROM atividades")

    # Obter os resultados da consulta
    for row in cursor:
        print(row)

    cursor.commit()
    cursor.close()

print("Job Concluído com Sucesso!")


