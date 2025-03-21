import os
import sys
import csv
import random
import string
import unicodedata
from PyPDF2 import PdfReader, PdfWriter

def extrair_nome_aluno(texto_pagina):
    """
    Procura a string 'DATA DE NASCIMENTO' e retorna tudo que estiver
    depois dessa expressão como o nome do aluno.
    """
    linhas = texto_pagina.splitlines()
    for linha in linhas:
        if "DATA DE NASCIMENTO" in linha.upper():
            indice = linha.upper().find("DATA DE NASCIMENTO")
            nome_aluno = linha[indice + len("DATA DE NASCIMENTO"):].strip()
            return nome_aluno
    return "Nome_Indefinido"

def gerar_senha(tamanho=5):
    """
    Gera uma senha aleatória composta apenas por dígitos.
    Exemplo: 12345, 90817, etc.
    """
    caracteres = string.digits
    return ''.join(random.choice(caracteres) for _ in range(tamanho))

def remove_acentos(input_str):
    """
    Remove os acentos de uma string.
    """
    nfkd_form = unicodedata.normalize("NFD", input_str)
    somente_ascii = nfkd_form.encode("ASCII", "ignore")
    return somente_ascii.decode("utf-8")

def main():
    pasta_origem = "input"
    pasta_save = "saves"  # Pasta principal para os outputs separados
    os.makedirs(pasta_origem, exist_ok=True)
    os.makedirs(pasta_save, exist_ok=True)

    # Lista todos os PDFs na pasta de origem
    arquivos_pdf = [arq for arq in os.listdir(pasta_origem) if arq.lower().endswith(".pdf")]
    if not arquivos_pdf:
        print("Nenhum arquivo PDF encontrado na pasta 'input'.")
        sys.exit()

    # Processa cada arquivo PDF separadamente
    for arquivo in arquivos_pdf:
        caminho_pdf = os.path.join(pasta_origem, arquivo)
        # Cria uma pasta com o mesmo nome (sem extensão) para esse arquivo
        nome_base = os.path.splitext(arquivo)[0]
        pasta_saida = os.path.join(pasta_save, nome_base)
        os.makedirs(pasta_saida, exist_ok=True)
        print(f"Processando PDF: {caminho_pdf} em pasta: {pasta_saida}")

        lista_links = []  # Armazenará (nome do arquivo, nome do aluno) para gerar o HTML

        # Caminho do CSV com as senhas referentes somente a este PDF de entrada
        csv_path = os.path.join(pasta_saida, "senhas.csv")
        with open(csv_path, "w", encoding="utf-8", newline="") as csvfile:
            writer = csv.writer(csvfile, delimiter=";")
            writer.writerow(["Arquivo", "Senha"])  # Cabeçalho do CSV

            with open(caminho_pdf, "rb") as f:
                leitor = PdfReader(f)
                total_paginas = len(leitor.pages)
                for i in range(total_paginas):
                    pagina = leitor.pages[i]
                    texto = pagina.extract_text() or ""
                    
                    # Extrai o nome do aluno a partir do conteúdo da página
                    nome_aluno = extrair_nome_aluno(texto)
                    if not nome_aluno:
                        nome_aluno = f"pagina_{i+1}"

                    # Remove caracteres problemáticos para nomes de arquivos
                    nome_aluno_seguro = "".join(c for c in nome_aluno if c.isalnum() or c in " -_").strip()
                    if not nome_aluno_seguro:
                        nome_aluno_seguro = f"pagina_{i+1}"

                    # Define o nome final do PDF para esta página
                    nome_arquivo_saida = f"{nome_aluno_seguro}.pdf"
                    caminho_saida_pdf = os.path.join(pasta_saida, nome_arquivo_saida)

                    # Cria o PDF com apenas esta página e o bloqueia com uma senha
                    escritor = PdfWriter()
                    escritor.add_page(pagina)
                    senha = gerar_senha()
                    escritor.encrypt(senha)
                    with open(caminho_saida_pdf, "wb") as out_f:
                        escritor.write(out_f)

                    writer.writerow([nome_arquivo_saida, senha])
                    lista_links.append((nome_arquivo_saida, nome_aluno_seguro))
                    print(f"  Criado e travado: {caminho_saida_pdf} | Senha: {senha}")

        # Gera o código-fonte HTML com links (em grupos de 3)
        html_snippet = '<table style="border-collapse: collapse; width: 100%;">\n<tbody>\n'
        for i in range(0, len(lista_links), 3):
            bloco = lista_links[i:i+3]
            html_snippet += "  <tr>\n"
            for (arquivo_pdf, nome_aluno) in bloco:
                arquivo_pdf_html = remove_acentos(arquivo_pdf).replace(" ", "-")
                html_snippet += (
                    '    <td style="width: 33.3333%; padding: 8px;">'
                    f'<a href="https://colegiouniverso.com.br/wp-content/uploads/2025/03/{arquivo_pdf_html}" '
                    'style="display: block; background-color: navy; color: white; text-decoration: none; '
                    'border: 1px solid white; border-radius: 25px; padding: 8px; text-align: center;">'
                    f'{nome_aluno}</a></td>\n'
                )
            if len(bloco) < 3:
                for _ in range(3 - len(bloco)):
                    html_snippet += '    <td style="width: 33.3333%;"></td>\n'
            html_snippet += "  </tr>\n"
        html_snippet += "</tbody>\n</table>"


        # Salva o HTML na pasta deste input (arquivo "z.txt")
        html_path = os.path.join(pasta_saida, "z.txt")
        with open(html_path, "w", encoding="utf-8") as f:
            f.write(html_snippet)
        print(f"HTML gerado em: {html_path}\n")

    print("Processo concluído!")

if __name__ == "__main__":
    main()
