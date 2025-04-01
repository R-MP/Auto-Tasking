import os
import csv

def process_csv_files(input_folder, output_txt):
    """
    Lê todos os arquivos CSV na pasta 'input_folder', ignorando o cabeçalho,
    e verifica quais alunos possuem senha com exatamente 4 dígitos.
    Para cada caso, gera uma linha no formato:
      (senha) - (nome do aluno) - (nome do arquivo CSV)
    Exibe as linhas no console e salva todas em 'output_txt'.
    """
    found_lines = []
    
    # Lista todos os arquivos CSV na pasta input_folder
    csv_files = [f for f in os.listdir(input_folder) if f.lower().endswith('.csv')]
    if not csv_files:
        print("Nenhum arquivo CSV encontrado na pasta '{}'.".format(input_folder))
        return
    
    for csv_file in csv_files:
        csv_path = os.path.join(input_folder, csv_file)
        with open(csv_path, newline='', encoding='utf-8') as f:
            reader = csv.reader(f, delimiter=";")
            first_line = True
            for row in reader:
                if first_line:
                    first_line = False
                    continue  # ignora a linha de cabeçalho
                if len(row) < 2:
                    continue
                nome = row[0].strip()
                senha = row[1].strip()
                # Se a senha tiver 4 dígitos, adiciona a linha
                if len(senha) == 4:
                    line = f"({senha}) - ({nome}) - ({csv_file})"
                    found_lines.append(line)
                    print(line)
    
    # Se houver alunos com senha de 4 dígitos, salva no arquivo de texto
    if found_lines:
        with open(output_txt, "w", encoding="utf-8") as f:
            for line in found_lines:
                f.write(line + "\n")
        print(f"\nArquivo de saída gerado: {output_txt}")
    else:
        print("Nenhum aluno com senha de 4 dígitos foi encontrado.")

def main():
    input_folder = "input"          # Pasta onde estão os arquivos CSV
    output_txt = "output/4digit_pass.txt"  # Nome do arquivo de saída
    os.makedirs(input_folder, exist_ok=True)  # Garante que a pasta input existe
    process_csv_files(input_folder, output_txt)

if __name__ == "__main__":
    main()
