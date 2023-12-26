# -*- coding: utf-8 -*-
import pandas as pd
import shutil

diretorio_origem = "G:\\Drives compartilhados\\Suprimentos Bulbe\\09 - Veículos\\Controle de Frota.xlsx"
diretorio_destino = "C:\\Users\\Bulbe\\OneDrive - Bulbe\\Power Automate\\Controle de Frota\\Controle de Frota.xlsx"
shutil.copy2(diretorio_origem, diretorio_destino)


arquivo_excel = "C:\\Users\\Bulbe\\OneDrive - Bulbe\\Power Automate\\Controle de Frota\\Controle de Frota.xlsx"
arquivo_csv = "C:\\Users\\Bulbe\\OneDrive - Bulbe\\Power Automate\\Controle de Frota\\Controle de Frota.csv"

dados_total = []

df = pd.read_excel(arquivo_excel, sheet_name="Gastos", usecols="B:H", header=1, engine='openpyxl')
dados_total.append(df)


dados_concatenados = pd.concat(dados_total)
dados_concatenados.to_csv(arquivo_csv, sep=";", index=False, encoding="utf-8-sig", decimal=',')
