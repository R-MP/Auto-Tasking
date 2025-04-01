import os
import re
import unicodedata

def remove_acentos(s):
    """Remove acentos de uma string."""
    nfkd_form = unicodedata.normalize("NFD", s)
    return "".join([c for c in nfkd_form if not unicodedata.combining(c)])

def clean_filename(s):
    """
    Remove a extensão ".pdf", acentos e espaços extras de um nome.
    """
    s = re.sub(r'\.pdf$', '', s, flags=re.IGNORECASE)
    s = remove_acentos(s)
    s = s.strip()
    return s

def main():
    input_folder = "input"
    
    # Garante que a pasta input exista
    if not os.path.exists(input_folder):
        os.makedirs(input_folder, exist_ok=True)
        print("Pasta 'input' criada. Coloque os arquivos PDF nela e execute novamente.")
        return

    # Lista todos os arquivos PDF na pasta input
    pdf_files = [f for f in os.listdir(input_folder) if f.lower().endswith(".pdf")]
    if not pdf_files:
        print("Nenhum arquivo PDF encontrado na pasta 'input'.")
        return

    # Gera o código HTML agrupando os links em linhas de 3
    html_snippet = '<table style="border-collapse: collapse; width: 100%;">\n<tbody>\n'
    for i in range(0, len(pdf_files), 3):
        grupo = pdf_files[i:i+3]
        html_snippet += "  <tr>\n"
        for pdf_file in grupo:
            # Limpa o nome removendo a extensão e acentos
            nome_aluno = clean_filename(pdf_file)
            # Para o link, substitui os espaços por hífens e anexa ".pdf"
            link_pdf = nome_aluno.replace(" ", "-") + ".pdf"
            # URL base (ajuste se necessário)
            url = f"https://colegiouniverso.com.br/wp-content/uploads/2025/03/{link_pdf}"
            html_snippet += (
                '    <td style="width: 33.3333%; padding: 8px;">'
                f'<a href="{url}" style="display: block; background-color: navy; color: white; text-decoration: none; '
                'border: 1px solid white; border-radius: 25px; padding: 8px; text-align: center;">'
                f'{nome_aluno}</a></td>\n'
            )
        # Preenche as colunas restantes se houver menos de 3 itens no grupo
        if len(grupo) < 3:
            for _ in range(3 - len(grupo)):
                html_snippet += '    <td style="width: 33.3333%;"></td>\n'
        html_snippet += "  </tr>\n"
    html_snippet += "</tbody>\n</table>"

    # Salva o HTML em um arquivo "z.txt" na pasta input (pode ser ajustado)
    output_html = os.path.join(input_folder, "z.txt")
    with open(output_html, "w", encoding="utf-8") as f:
        f.write(html_snippet)
    
    print(f"HTML gerado em: {output_html}")

if __name__ == "__main__":
    main()
