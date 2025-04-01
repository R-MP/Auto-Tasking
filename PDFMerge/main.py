import os
import glob
from PyPDF2 import PdfMerger

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
    merger = PdfMerger()
    
    # Ordena os arquivos para mesclar em ordem alfabética
    for pdf in sorted(arquivos_pdf):
        print(f"Adicionando {pdf}...")
        merger.append(pdf)
    
    # Define o nome do arquivo de saída
    pdf_saida = os.path.join(pasta_output, "merged.pdf")
    merger.write(pdf_saida)
    merger.close()
    
    print(f"PDF mesclado criado com sucesso: {pdf_saida}")
