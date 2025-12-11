import sys
import os
import pandas as pd
import csv
import re

# ================================
# Converte letras de coluna para índice numérico (A=0, B=1, ..., AA=26)
# ================================
def col_letter_to_index(col):
    idx = 0
    for c in col:
        idx = idx * 26 + (ord(c.upper()) - ord('A') + 1)
    return idx - 1

# ================================
# Função principal
# ================================
def processar_arquivo():
    pasta_input = "../input"
    pasta_output = "../output"

    # Localiza primeiro arquivo XLS/XLSX
    arquivo_excel = None
    for f in os.listdir(pasta_input):
        if f.lower().endswith((".xls", ".xlsx")):
            arquivo_excel = os.path.join(pasta_input, f)
            nome_entrada = f
            break

    if not arquivo_excel:
        print("Nenhum arquivo XLS/XLSX encontrado na pasta input/")
        return

    # Lê o Excel ignorando duas linhas de cabeçalho
    try:
        df = pd.read_excel(arquivo_excel, dtype=str, header=None, keep_default_na=False, skiprows=2)
    except Exception as e:
        print(f"Erro ao ler o Excel: {e}")
        return

    # Mapeamento de colunas por letra
    idx_B = col_letter_to_index("B")  # Para coluna 05
    idx_R = col_letter_to_index("R")  # Para coluna 03

    # Construção do dataframe final
        # Formatação coluna 03: manter somente números
    col3 = df[idx_R].astype(str).apply(lambda x: re.sub(r"[^0-9]", "", x))

    df_final = pd.DataFrame({
        1: "",
        2: "",
        3: col3,
        4: "",
        5: df[idx_B].astype(str)
    })

    # Garante que pasta output existe
    os.makedirs(pasta_output, exist_ok=True)

    # Define nome de saída mantendo nome original + "_formatado"
    nome_base = os.path.splitext(nome_entrada)[0]
    nome_saida = f"{nome_base}_formatado.txt"
    caminho_saida = os.path.join(pasta_output, nome_saida)

    # Grava CSV separado por vírgula
    try:
        df_final.to_csv(caminho_saida, sep=",", index=False, header=False,
                        quoting=csv.QUOTE_NONE, escapechar='\\', encoding='utf-8')
    except Exception as e:
        print(f"Erro ao escrever TXT: {e}")
        return

    print(f"Arquivo gerado com sucesso: {caminho_saida}")


if __name__ == "__main__":
    processar_arquivo()
