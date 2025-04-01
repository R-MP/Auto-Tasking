import subprocess
import os
import glob
import sys
import shutil

def main():
    # Executa o primeiro script
    print("Gerando arquivo docx...")
    try:
        subprocess.run(["python", "generate.py"], check=True)
    except subprocess.CalledProcessError as e:
        print("Erro na execução do generate.py:", e)
        sys.exit(1)
    print("generate.py concluído.")

    # Executa o segundo script
    print("Executando a conversão para PDF...")
    try:
        subprocess.run(["python", "convert.py"], check=True)
    except subprocess.CalledProcessError as e:
        print("Erro na execução do convert.py:", e)
        sys.exit(1)
    print("convert.py concluído.")

    # Procura na pasta "input" pelo(s) arquivo(s) CSV
    input_folder = "input"
    csv_files = glob.glob(os.path.join(input_folder, "*.csv"))
    if not csv_files:
        print("Nenhum arquivo CSV encontrado na pasta 'input'.")
        sys.exit(1)

    # Ordena os CSVs e usa o último (ou ajuste conforme sua preferência)
    csv_files.sort()
    final_csv = os.path.basename(csv_files[-1])
    base_name = os.path.splitext(final_csv)[0]
    print(f"Usando o nome '{base_name}' do CSV para renomear o PDF final.")

    final_folder = "final"
    os.makedirs(final_folder, exist_ok=True)

    # Renomeia o arquivo final PDF
    # Supondo que o arquivo final gerado seja "final - .pdf" (com um espaço antes do ponto)
    old_pdf = "final - .pdf"
    new_pdf = os.path.join(final_folder, f"final - {base_name}.pdf")
    
    if os.path.exists(old_pdf):
        os.rename(old_pdf, new_pdf)
        print(f"Arquivo renomeado para '{new_pdf}'.")
    else:
        print(f"Arquivo '{old_pdf}' não encontrado.")
    
    for folder in ["output", "temp_pdf"]:
        if os.path.exists(folder):
            shutil.rmtree(folder)
            print(f"Pasta '{folder}' deletada.")
        else:
            print(f"Pasta '{folder}' não encontrada.")

    # Move todos os arquivos da pasta "input" para a pasta "used"
    used_folder = "used"
    os.makedirs(used_folder, exist_ok=True)
    for file in os.listdir("input"):
        src_path = os.path.join("input", file)
        dst_path = os.path.join(used_folder, file)
        shutil.move(src_path, dst_path)
        print(f"Arquivo '{file}' movido para '{used_folder}'.")


if __name__ == "__main__":
    main()
