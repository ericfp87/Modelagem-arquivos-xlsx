# -*- coding: utf-8 -*-
import pandas as pd
import shutil

arquivo_excel = "C:\\Users\\Bulbe\\OneDrive - Bulbe\\Power Automate\\Controle Suprimentos\\Base de dados BI\\SAVING.xlsm"
arquivo_csv = "C:\\Users\\Bulbe\\OneDrive - Bulbe\\Power Automate\\Controle Suprimentos\\Base de dados BI\\SAVING.csv"

dados_total = []

df = pd.read_excel(arquivo_excel, sheet_name="Controle de Contratos Atualizad", usecols="A,C,F,H,I,J,K,N,R", header=5, engine='openpyxl')
dados_total.append(df)


dados_concatenados = pd.concat(dados_total)
dados_concatenados.to_csv(arquivo_csv, sep=";", index=False, encoding="utf-8-sig", decimal=',')
