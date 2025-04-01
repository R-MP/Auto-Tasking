import os
import csv
import re
import unicodedata
from docxtpl import DocxTemplate

def remove_acentos(s):
    """Remove acentos de uma string."""
    nfkd_form = unicodedata.normalize("NFD", s)
    return "".join([c for c in nfkd_form if not unicodedata.combining(c)])

def clean_filename(s):
    """
    Limpa a string para uso como nome de arquivo:
      - Remove ocorrências de ".pdf" (caso existam),
      - Remove acentos,
      - Remove caracteres inválidos,
      - Remove espaços extras.
    """
    s = re.sub(r'\.pdf$', '', s, flags=re.IGNORECASE)
    s = remove_acentos(s)
    s = s.strip()
    s = re.sub(r"[^\w\s-]", "", s)
    s = re.sub(r"\s+", " ", s)
    return s

def generate_comunicado_docx(template_path, student1_name, student1_password, student2_name, student2_password, output_docx_path):
    """
    Abre o template DOCX (deve estar com os placeholders no formato Jinja2:
    {{ NOME1 }}, {{ SENHA1 }}, {{ NOME2 }}, {{ SENHA2 }}),
    renderiza-o com os dados dos alunos e salva o resultado em output_docx_path.
    """
    doc = DocxTemplate(template_path)
    context = {
        "NOME1": student1_name,
        "SENHA1": student1_password,
        "NOME2": student2_name,
        "SENHA2": student2_password
    }
    doc.render(context)
    doc.save(output_docx_path)

def main():
    print("Iniciando processamento com docxtpl...")
    
    # Cria as pastas "input" e "output" se não existirem
    input_folder = "input"
    output_folder = "output"
    os.makedirs(input_folder, exist_ok=True)
    os.makedirs(output_folder, exist_ok=True)
    
    # Procura por arquivos CSV na pasta "input"
    csv_files = [f for f in os.listdir(input_folder) if f.lower().endswith(".csv")]
    if not csv_files:
        print("Nenhum arquivo CSV encontrado na pasta 'input'.")
        return
    
    # Agrega os dados de todos os CSVs (cada linha: (nome, senha))
    # A primeira linha de cada CSV é ignorada (cabeçalho)
    students = []
    for csv_file in csv_files:
        csv_path = os.path.join(input_folder, csv_file)
        with open(csv_path, newline="", encoding="utf-8") as csvfile:
            reader = csv.reader(csvfile, delimiter=";")
            first_line = True
            for row in reader:
                if first_line:
                    first_line = False
                    continue  # ignora cabeçalho
                if len(row) < 2:
                    continue
                name = row[0].strip()
                password = row[1].strip()
                # Remove ".pdf" do final, se existir
                name = re.sub(r'\.pdf$', '', name, flags=re.IGNORECASE)
                students.append((name, password))
    
    if not students:
        print("Nenhum dado de aluno encontrado nos CSVs.")
        return

    # Template DOCX – o arquivo deve estar no mesmo diretório que o script
    template_path = "template_comunicado.docx"
    if not os.path.exists(template_path):
        print("Template DOCX não encontrado!")
        return

    # Processa os alunos em grupos de 2 e gera um arquivo DOCX para cada grupo
    page_counter = 1
    i = 0
    while i < len(students):
        student1_name, student1_password = students[i]
        if i + 1 < len(students):
            student2_name, student2_password = students[i+1]
        else:
            student2_name, student2_password = "", ""
        
        # Define o nome do arquivo de saída: "pagina X.docx"
        output_docx_path = os.path.join(output_folder, f"pagina {page_counter}.docx")
        generate_comunicado_docx(template_path, student1_name, student1_password, student2_name, student2_password, output_docx_path)
        print(f"Comunicado gerado para grupo {page_counter}: {student1_name} | {student2_name} -> {output_docx_path}")
        
        page_counter += 1
        i += 2

    print("Processamento concluído.")

if __name__ == "__main__":
    print("Processamento iniciado... Aguarde.")
    main()
    print("Processamento concluído.")
