import os
import glob
from pdf2image import convert_from_path

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
    # Se estiver no Windows, informe o caminho para o poppler
    # Exemplo: poppler_path=r"C:\caminho\para\poppler-0.68.0\bin"
    poppler_path = None  # ou defina o caminho, se necessário
    
    for pdf in sorted(arquivos_pdf):
        # Obtém o nome base do arquivo (sem extensão)
        nome_pai = os.path.splitext(os.path.basename(pdf))[0]
        print(f"Processando: {pdf}")
        
        # Converte as páginas do PDF para imagens
        if poppler_path:
            paginas = convert_from_path(pdf, poppler_path="C:\poppler-24.08.0\Library\bin")
        else:
            paginas = convert_from_path(pdf)
            
        # Salva cada página como um arquivo JPEG
        for i, pagina in enumerate(paginas, start=1):
            nome_imagem = os.path.join(pasta_output, f"{nome_pai} - {i}.jpeg")
            pagina.save(nome_imagem, "JPEG")
            print(f"Salvo: {nome_imagem}")
            
    print("Conversão concluída!")
