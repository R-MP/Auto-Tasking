import os
import re
import tempfile
import sys
from pdf2image import convert_from_path
from docx import Document
from docx.shared import Inches
from docx2pdf import convert as docx2pdf_convert
from PIL import Image, ImageEnhance

# Define as pastas de entrada e saída
input_folder = "input"
output_folder = "output"
os.makedirs(input_folder, exist_ok=True)
os.makedirs(output_folder, exist_ok=True)

# Se necessário, defina o caminho para o poppler (no Windows, por exemplo)
poppler_path = None  # Exemplo: r"C:\poppler-24.08.0\Library\bin"

# Define a largura desejada para a imagem inserida no Word: 80% da largura de A4 (A4 ~8.27 in)
target_width_inch = 8.27 * 0.95  # aproximadamente 6.62 inches

print("Iniciando processamento recursivo...")

# Percorre recursivamente a pasta input
for root, dirs, files in os.walk(input_folder):
    # Calcula o caminho relativo para recriar a estrutura em output
    rel_path = os.path.relpath(root, input_folder)
    out_subfolder = os.path.join(output_folder, rel_path)
    os.makedirs(out_subfolder, exist_ok=True)
    
    for file in files:
        if file.lower().endswith(".pdf"):
            pdf_path = os.path.join(root, file)
            base_name = os.path.splitext(file)[0]
            print(f"\nProcessando PDF: {pdf_path}")
            
            # Converte o PDF em imagens (uma por página), com DPI 300 e use_cropbox=True
            try:
                if poppler_path:
                    pages = convert_from_path(pdf_path, dpi=300, use_cropbox=True, poppler_path=poppler_path)
                else:
                    pages = convert_from_path(pdf_path, dpi=300, use_cropbox=True)
            except Exception as e:
                print(f"Erro ao converter PDF para imagens: {e}")
                continue
            
            page_number = 1
            for page_image in pages:
                with tempfile.TemporaryDirectory() as tmpdirname:
                    # Salva a imagem da página em um arquivo temporário
                    temp_img_path = os.path.join(tmpdirname, "page.jpg")
                    page_image.save(temp_img_path, "JPEG")
                    
                    # Cria um documento Word com página A4 e margens zeradas
                    doc = Document()
                    section = doc.sections[0]
                    section.top_margin = Inches(0)
                    section.left_margin = Inches(0)
                    section.right_margin = Inches(0)
                    section.bottom_margin = Inches(0)
                    
                    # Insere a imagem no documento, dimensionada para target_width_inch
                    doc.add_picture(temp_img_path, width=Inches(target_width_inch))
                    
                    # Salva o documento Word temporário
                    temp_docx_path = os.path.join(tmpdirname, "temp.docx")
                    doc.save(temp_docx_path)
                    
                    # Converte o DOCX para PDF
                    temp_pdf_path = os.path.join(tmpdirname, "temp.pdf")
                    try:
                        docx2pdf_convert(temp_docx_path, temp_pdf_path)
                    except Exception as e:
                        print(f"Erro na conversão DOCX->PDF: {e}")
                        continue
                    
                    # Converte o PDF temporário para imagem (deve ter 1 página)
                    try:
                        if poppler_path:
                            pdf_images = convert_from_path(temp_pdf_path, dpi=300, poppler_path=poppler_path)
                        else:
                            pdf_images = convert_from_path(temp_pdf_path, dpi=300)
                    except Exception as e:
                        print(f"Erro na conversão PDF->imagem: {e}")
                        continue
                    
                    if pdf_images:
                        final_img = pdf_images[0]
                        
                        # Aumenta o contraste da imagem final usando ImageEnhance
                        enhancer = ImageEnhance.Contrast(final_img)
                        final_img = enhancer.enhance(2)  # Fator 1.5 para aumentar o contraste em 50%
                        
                        # Define o nome do arquivo final: "<base_name> - <page_number>.jpg"
                        output_img_name = f"{base_name}.jpg"
                        output_img_path = os.path.join(out_subfolder, output_img_name)
                        try:
                            final_img.save(output_img_path, "JPEG")
                            print(f"Salvo: {output_img_path}")
                        except Exception as e:
                            print(f"Erro ao salvar {output_img_path}: {e}")
                    else:
                        print("Nenhuma imagem gerada a partir do PDF temporário.")
                page_number += 1

print("\nProcessamento concluído!")
