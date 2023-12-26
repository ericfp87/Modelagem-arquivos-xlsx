# -*- coding: utf-8 -*-
import pandas as pd
import shutil

diretorio_origem = "G:\\Drives compartilhados\\Suprimentos Bulbe\\26 - Adiantamentos\\Adiantamento 1º semestre.xlsx"
diretorio_destino = "C:\\Users\\Bulbe\\OneDrive - Bulbe\\Power Automate\\Adiantamentos\\Adiantamento 1º semestre.xlsx"
shutil.copy2(diretorio_origem, diretorio_destino)


arquivo_excel = "C:\\Users\\Bulbe\\OneDrive - Bulbe\\Power Automate\\Adiantamentos\\Adiantamento 1º semestre.xlsx"
arquivo_csv = "C:\\Users\\Bulbe\\OneDrive - Bulbe\\Power Automate\\Adiantamentos\\Adiantamentos.csv"

dados_total = []

df = pd.read_excel(arquivo_excel, sheet_name="Adiantamentos", usecols="A:M", header=0, engine='openpyxl')
dados_total.append(df)


dados_concatenados = pd.concat(dados_total)
dados_concatenados.to_csv(arquivo_csv, sep=";", index=False, encoding="utf-8-sig", decimal=',')
