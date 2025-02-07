import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, NamedStyle
from openpyxl.utils.dataframe import dataframe_to_rows

# Caminho do arquivo original e do arquivo de saída
input_file = r"C:\Users\gustavo.telles\Desktop\Case Itaú\case_estagio.xlsx"
output_file = r"C:\Users\gustavo.telles\Desktop\Case Itaú\case_estagio_limpo.xlsx"

# Exclui o arquivo de saída se já existir
if os.path.exists(output_file):
    os.remove(output_file)

# Lê a planilha
df = pd.read_excel(input_file)

# Remove linhas com células vazias
df = df.dropna()

# Padroniza colunas de datas no formato datetime
for column in df.select_dtypes(include=['datetime']):
    df[column] = pd.to_datetime(df[column])  # Garante que as colunas sejam do tipo datetime

# Padroniza o nome dos vendedores (primeira letra maiúscula)
if 'Vendedor' in df.columns:  # Verifica se a coluna existe
    df['Vendedor'] = df['Vendedor'].str.title()

# Remove duplicatas
df = df.drop_duplicates()

# Cria um estilo para as células de data
date_style = NamedStyle(name="datetime", number_format="DD/MM/YYYY")

# Salva o DataFrame no Excel e aplica o estilo para datas
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    df.to_excel(writer, index=False, sheet_name="Dados Limpos")
    workbook = writer.book
    worksheet = writer.sheets["Dados Limpos"]

    # Aplica alinhamento e formatação de data
    for col_idx, column in enumerate(df.columns, start=1):
        if df[column].dtype == 'datetime64[ns]':  # Verifica se a coluna é do tipo datetime
            for row_idx in range(2, len(df) + 2):  # Aplica estilo às células da coluna de datas
                cell = worksheet.cell(row=row_idx, column=col_idx)
                cell.style = date_style
                cell.alignment = Alignment(horizontal='right')

    # Ativa filtros na planilha
    worksheet.auto_filter.ref = worksheet.dimensions

print(f"Arquivo processado salvo com datas formatadas corretamente como {output_file}.")
