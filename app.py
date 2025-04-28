import pandas as pd
from tkinter import Tk, filedialog
from io import BytesIO
import os

# 1. Abrir janela para selecionar arquivos .xlsx
Tk().withdraw()  # Oculta a janela principal do Tkinter
file_paths = filedialog.askopenfilenames(
    title="Selecione os arquivos .xlsx",
    filetypes=[("Excel files", "*.xlsx")]
)

# 2. Entrada do nome a buscar
nome_busca = input("Digite o nome que deseja buscar: ").strip().lower()

# 3. Verificação do nome nos arquivos
planilhas_com_nome = []

for path in file_paths:
    filename = os.path.basename(path)
    try:
        xls = pd.ExcelFile(path)
        for sheet_name in xls.sheet_names:
            df = xls.parse(sheet_name)
            if df.astype(str).apply(lambda x: x.str.lower()).isin([nome_busca]).any().any():
                planilhas_com_nome.append(filename)
                break
    except Exception as e:
        print(f"Erro ao processar {filename}: {e}")

# 4. Resultado
if planilhas_com_nome:
    print("\nNome encontrado nos seguintes arquivos:")
    for p in planilhas_com_nome:
        print(f"- {p}")
else:
    print("\nNome não encontrado em nenhum dos arquivos enviados.")
