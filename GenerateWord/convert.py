import os
import re
from docx2pdf import convert
from PyPDF2 import PdfMerger

def convert_docx_to_pdf(docx_folder, pdf_folder):
    """
    Converte todos os arquivos DOCX da pasta 'docx_folder' para PDF,
    salvando-os na pasta 'pdf_folder'. Retorna uma lista com os caminhos dos PDFs.
    """
    os.makedirs(pdf_folder, exist_ok=True)
    pdf_files = []
    docx_files = [f for f in os.listdir(docx_folder) if f.lower().endswith(".docx")]
    
    for file in docx_files:
        docx_path = os.path.join(docx_folder, file)
        # Define o nome do PDF correspondente, substituindo a extensão
        pdf_filename = re.sub(r'\.docx$', '.pdf', file, flags=re.IGNORECASE)
        pdf_path = os.path.join(pdf_folder, pdf_filename)
        try:
            convert(docx_path, pdf_path)
            print(f"Convertido: {file} -> {pdf_filename}")
            pdf_files.append(pdf_path)
        except Exception as e:
            print(f"Erro na conversão de {file}: {e}")
    return pdf_files

def merge_pdfs(pdf_files, output_pdf):
    """
    Junta (merge) os arquivos PDF listados em 'pdf_files' em um único PDF,
    salvando-o como 'output_pdf'.
    """
    merger = PdfMerger()
    # Ordena os arquivos PDF pela numeração presente no nome, por exemplo "pagina 1.pdf", "pagina 2.pdf", ...
    pdf_files.sort(key=lambda x: int(re.search(r'\d+', os.path.basename(x)).group()))
    for pdf in pdf_files:
        merger.append(pdf)
    merger.write(output_pdf)
    merger.close()
    print(f"Arquivo final mesclado gerado: {output_pdf}")

def main():
    # Pasta onde estão os arquivos DOCX a serem mesclados (cada um com uma página)
    docx_folder = "output"
    # Pasta para armazenar os PDFs temporários
    pdf_folder = "temp_pdf"
    # Caminho do PDF final
    output_pdf = "final - .pdf"
    
    # Converte os DOCX para PDF
    pdf_files = convert_docx_to_pdf(docx_folder, pdf_folder)
    if not pdf_files:
        print("Nenhum PDF foi gerado a partir dos DOCX.")
        return
    
    # Mescla os PDFs em um único arquivo
    merge_pdfs(pdf_files, output_pdf)

if __name__ == "__main__":
    print("Iniciando conversão e merge dos DOCX em um único PDF...")
    main()
    print("Processamento concluído.")
