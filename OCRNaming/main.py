import os
import re
import sys
import pytesseract
from pdf2image import convert_from_path
from PyPDF2 import PdfReader, PdfWriter
import unicodedata

# Configure o caminho do Tesseract, se necessário (Windows)
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def remove_acentos(s):
    """Remove acentos de uma string."""
    nfkd_form = unicodedata.normalize("NFD", s)
    return "".join([c for c in nfkd_form if not unicodedata.combining(c)])

def clean_filename(s):
    """
    Remove a extensão .pdf (se houver), remove acentos e espaços extras.
    """
    s = re.sub(r'\.pdf$', '', s, flags=re.IGNORECASE)
    s = remove_acentos(s)
    return s.strip()

def extract_student_name(image):
    """
    Aplica OCR na imagem e procura pela frase "folha de respostas". 
    Em seguida, retorna a próxima linha não vazia (que assumimos conter o nome do aluno).
    Se não encontrar, retorna None.
    """
    text = pytesseract.image_to_string(image, lang='por')
    lines = text.splitlines()
    student_name = None
    for i, line in enumerate(lines):
        if "folha de respostas" in line.lower():
            # Procura pela próxima linha não vazia e que não contenha novamente "folha de respostas"
            for j in range(i+1, len(lines)):
                candidate = lines[j].strip()
                if candidate and "folha de respostas" not in candidate.lower():
                    student_name = candidate
                    break
            break
    return student_name

def main():
    input_folder = "input"
    output_folder = "output"
    os.makedirs(input_folder, exist_ok=True)
    os.makedirs(output_folder, exist_ok=True)
    
    # Lista todos os PDFs na pasta de input
    pdf_files = [f for f in os.listdir(input_folder) if f.lower().endswith(".pdf")]
    if not pdf_files:
        print("Nenhum arquivo PDF encontrado na pasta 'input'.")
        sys.exit()
    
    for pdf_file in pdf_files:
        input_pdf_path = os.path.join(input_folder, pdf_file)
        base_name = clean_filename(pdf_file)  # Ex: "FUND 1A"
        print(f"Processando: {input_pdf_path}")
        
        # Abre o PDF original para extrair páginas
        try:
            reader = PdfReader(input_pdf_path)
            num_pages = len(reader.pages)
        except Exception as e:
            print(f"Erro ao abrir o PDF {pdf_file}: {e}")
            continue
        
        # Converte todas as páginas para imagens (pdf2image)
        try:
            images = convert_from_path(input_pdf_path)
        except Exception as e:
            print(f"Erro na conversão do PDF para imagens: {e}")
            continue
        
        for i in range(num_pages):
            image = images[i]
            student_name = extract_student_name(image)
            if not student_name:
                student_name = f"pagina_{i+1}"
            # Cria o nome do arquivo de saída: "<base_name> - <student_name>.pdf"
            output_filename = f"{base_name} - {student_name}.pdf"
            output_pdf_path = os.path.join(output_folder, output_filename)
            
            # Extrai a página i do PDF original e salva em um novo PDF
            writer = PdfWriter()
            writer.add_page(reader.pages[i])
            try:
                with open(output_pdf_path, "wb") as f_out:
                    writer.write(f_out)
                print(f"Página {i+1} de {pdf_file} salva como: {output_filename}")
            except Exception as e:
                print(f"Erro ao salvar página {i+1} de {pdf_file}: {e}")

if __name__ == "__main__":
    main()
