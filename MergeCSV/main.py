import os
import csv
import re
import unicodedata

def remove_acentos(s):
    """
    Remove os acentos de uma string utilizando a normalização Unicode (NFD).
    """
    nfkd_form = unicodedata.normalize("NFD", s)
    return "".join([c for c in nfkd_form if not unicodedata.combining(c)])

def main():
    input_folder = "input"      # Pasta onde estão os CSVs (em subpastas ou na raiz)
    output_file = "geral.csv"   # Arquivo final a ser gerado
    
    all_rows = []  # Lista que armazenará os dados (nome e senha)

    # Percorre recursivamente a pasta "input"
    for root, dirs, files in os.walk(input_folder):
        for file in files:
            if file.lower().endswith(".csv"):
                csv_path = os.path.join(root, file)
                print(f"Processando: {csv_path}")
                try:
                    with open(csv_path, newline="", encoding="utf-8") as csvfile:
                        reader = csv.reader(csvfile, delimiter=";")
                        # Lê a primeira linha para descartar o cabeçalho
                        header = next(reader, None)
                        for row in reader:
                            if len(row) < 2:
                                continue
                            nome = row[0].strip()
                            senha = row[1].strip()
                            # Remove a ocorrência de ".pdf" do final do nome, se houver
                            nome = re.sub(r'\.pdf$', '', nome, flags=re.IGNORECASE)
                            # Remove acentos do nome
                            nome = remove_acentos(nome)
                            # Garante que a senha tenha 5 dígitos (adiciona zeros à esquerda se necessário)
                            senha = senha.zfill(5)
                            # Adiciona apóstrofo para que o Excel interprete como texto
                            senha = "'" + senha
                            all_rows.append([nome, senha])
                except Exception as e:
                    print(f"Erro ao processar {csv_path}: {e}")

    # Escreve os dados agregados no arquivo final
    if all_rows:
        try:
            with open(output_file, "w", newline="", encoding="utf-8") as fout:
                writer = csv.writer(fout, delimiter=";")
                for row in all_rows:
                    writer.writerow(row)
            print(f"\nArquivo final '{output_file}' gerado com sucesso com {len(all_rows)} linhas.")
        except Exception as e:
            print(f"Erro ao escrever o arquivo final: {e}")
    else:
        print("Nenhum dado foi encontrado nos arquivos CSV.")

if __name__ == "__main__":
    main()
