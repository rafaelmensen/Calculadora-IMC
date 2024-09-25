# Calculadora-IMC

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

def copiar_e_bloquear_planilha(caminho_origem, caminho_destino, intervalo_linhas=(2, 100)):
    try:
        # Carregar o arquivo de origem
        wb_origem = load_workbook(caminho_origem)
        ws_origem = wb_origem.active
        
        # Bloquear a planilha de origem
        ws_origem.protection.sheet = True
        ws_origem.protection.password = "senha_protegida"  # Defina sua senha de proteção
        wb_origem.save(caminho_origem)
        
        # Carregar os dados das linhas especificadas
        df = pd.read_excel(caminho_origem, skiprows=intervalo_linhas[0] - 1, nrows=intervalo_linhas[1] - intervalo_linhas[0] + 1)
        
        # Carregar o arquivo de destino, criando um novo se não existir
        try:
            wb_destino = load_workbook(caminho_destino)
        except FileNotFoundError:
            wb_destino = load_workbook(caminho_origem)  # Clonar a estrutura do arquivo de origem
            wb_destino.remove(wb_destino.active)  # Remove a planilha ativa
        
        ws_destino = wb_destino.create_sheet(title="Dados Copiados")
        
        # Adicionar os dados à planilha de destino
        for row in dataframe_to_rows(df, index=False, header=True):
            ws_destino.append(row)
        
        # Bloquear a planilha de destino
        ws_destino.protection.sheet = True
        ws_destino.protection.password = "senha_protegida"  # Defina sua senha de proteção
        wb_destino.save(caminho_destino)
        
        print(f"Dados copiados e planilhas bloqueadas com sucesso de {intervalo_linhas[0]} a {intervalo_linhas[1]}.")

    except Exception as e:
        print(f"Ocorreu um erro: {e}")

# Configuração dos caminhos e execução da função
caminho_origem = 'planilha_origem.xlsx'
caminho_destino = 'planilha_destino.xlsx'
copiar_e_bloquear_planilha(caminho_origem, caminho_destino)