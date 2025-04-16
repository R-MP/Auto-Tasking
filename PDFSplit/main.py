import os
import glob
from PyPDF2 import PdfReader, PdfWriter

# Define as pastas de entrada e saída
pasta_input = "input"
pasta_output = "output"
os.makedirs(pasta_input, exist_ok=True)
os.makedirs(pasta_output, exist_ok=True)

# Procura por todos os arquivos PDF na pasta input
arquivos_pdf = glob.glob(os.path.join(pasta_input, "*.pdf"))

if not arquivos_pdf:
    print("Nenhum arquivo PDF encontrado na pasta 'input'.")
else:
    for arquivo in arquivos_pdf:
        # Obtém o nome base do arquivo (sem extensão)
        nome_pai = os.path.splitext(os.path.basename(arquivo))[0]
        # Cria uma subpasta na pasta output com o nome do arquivo
        pasta_destino = os.path.join(pasta_output, nome_pai)
        os.makedirs(pasta_destino, exist_ok=True)
        print(f"Processando o arquivo: {arquivo} em pasta: {pasta_destino}")

        # Abre o PDF para leitura
        leitor = PdfReader(arquivo)
        total_paginas = len(leitor.pages)
        
        # Para cada página, cria um novo PDF na subpasta
        for i, pagina in enumerate(leitor.pages, start=1):
            escritor = PdfWriter()
            escritor.add_page(pagina)
            
            # Define o nome do arquivo de saída, ex: "3b - 1.pdf", "3b - 2.pdf", ...
            nome_saida = os.path.join(pasta_destino, f"{nome_pai} - {i}.pdf")
            with open(nome_saida, "wb") as f:
                escritor.write(f)
            print(f"Gerado: {nome_saida}")
        
    print("Processamento concluído!")
