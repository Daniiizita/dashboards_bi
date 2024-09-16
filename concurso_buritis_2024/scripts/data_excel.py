import pandas as pd
import os

# Caminho para o arquivo Excel
arquivo_excel = "buritis_concurso2024.xlsx"

# Carregar todas as planilhas do arquivo Excel
xls = pd.ExcelFile(arquivo_excel)

# Lista para armazenar os DataFrames de cada planilha
todas_planilhas = []

# Iterar por cada nome de planilha e ler o conteúdo
for nome_planilha in xls.sheet_names:
    df = pd.read_excel(xls, sheet_name=nome_planilha)
    todas_planilhas.append(df)

# Concatenar todas as planilhas em um único DataFrame
df_completo = pd.concat(todas_planilhas, ignore_index=True)

# Obter o nome do arquivo sem a extensão
nome_arquivo, extensao = os.path.splitext(os.path.basename(arquivo_excel))

# Novo nome do arquivo com o prefixo 'merged_'
novo_arquivo_excel = f"merged_{nome_arquivo}{extensao}"

# Obter o caminho absoluto da pasta onde o script está sendo executado
pasta_atual = os.path.abspath(os.getcwd())

# Caminho completo para o novo arquivo Excel
caminho_arquivo_final = os.path.join(pasta_atual, novo_arquivo_excel)

# Salvar o DataFrame completo em um novo arquivo Excel
with pd.ExcelWriter(caminho_arquivo_final, mode='w') as writer:
    df_completo.to_excel(writer, sheet_name='Planilha_Completa', index=False)

print(f"Todas as planilhas foram combinadas com sucesso em '{caminho_arquivo_final}'!")

