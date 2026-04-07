import os
from datetime import datetime
import pandas as pd
from docx import Document

pasta = r"C:\Users\CRC\Dropbox\2025\Saída de trabalhos"

resultados = []

def extrair_dados(doc):
    data_entrada = None
    data_envio = None
    ensaios = []

    for tabela in doc.tables:
        for row_idx, linha in enumerate(tabela.rows):
            for col_idx, celula in enumerate(linha.cells):

                texto = celula.text.strip().lower()

                # -------- DATAS --------
                if "data entrada" in texto:
                    if col_idx + 1 < len(linha.cells):
                        data_entrada = linha.cells[col_idx + 1].text.strip()

                if "resultados enviados em" in texto:
                    if col_idx + 1 < len(linha.cells):
                        data_envio = linha.cells[col_idx + 1].text.strip()

                # -------- COLUNA ENSAIO --------
                if "ensaio" in texto:

                    # Ler até 12 células abaixo na mesma coluna
                    for i in range(1, 13):
                        if row_idx + i < len(tabela.rows):
                            try:
                                valor = tabela.rows[row_idx + i].cells[col_idx].text.strip()
                                if valor:  # só adiciona se não estiver vazio
                                    ensaios.append(valor)
                            except:
                                pass

    return data_entrada, data_envio, ensaios


for arquivo in os.listdir(pasta):
    if arquivo.endswith(".docx"):
        caminho_arquivo = os.path.join(pasta, arquivo)
        doc = Document(caminho_arquivo)

        data_entrada_str, data_envio_str, ensaios = extrair_dados(doc)

        if data_entrada_str and data_envio_str:
            try:
                data_entrada = datetime.strptime(data_entrada_str, "%d/%m/%Y")
                data_envio = datetime.strptime(data_envio_str, "%d/%m/%Y")
                dias = (data_envio - data_entrada).days
            except:
                continue
        else:
            continue

        linha_resultado = {
            "Arquivo": arquivo,
            "Data Entrada": data_entrada,
            "Data Envio": data_envio,
            "Dias no Laboratório": dias
        }

        # Garantir 12 colunas fixas
        for i in range(12):
            linha_resultado[f"Ensaio {i+1}"] = ensaios[i] if i < len(ensaios) else ""

        resultados.append(linha_resultado)

df = pd.DataFrame(resultados)
df.to_excel("tempo_laboratorio.xlsx", index=False)

print("Planilha gerada com sucesso.")
