import win32com.client as win32
import os
import pandas as pd
import datetime

# Criando a integração com o Outlook
outlook = win32.Dispatch('outlook.application')

# Carregue o arquivo CSV com ";" como separador
caminho_do_arquivo = r'C:\beneficiarios\benef.csv'
dataframe = pd.read_csv(caminho_do_arquivo, sep=';')

# Salve o DataFrame de volta como um arquivo CSV com "," como separador
novo_caminho_do_arquivo = r'C:\beneficiarios\novo_arquivo.csv'
dataframe.to_csv(novo_caminho_do_arquivo, sep=',', index=False)
# Isso salvará o DataFrame no novo arquivo CSV com "," como separador

# Defina o caminho do arquivo de log
log_file_path = r'C:\beneficiarios\email_log.txt'

# Abra o arquivo de log em modo de apêndice (append)
with open(log_file_path, 'a') as log_file:
    log_file.write(f"--- Início do Log ({datetime.datetime.now()}) ---\n")

# Ler os dados do CSV usando Pandas
df = pd.read_csv(r"C:\beneficiarios\novo_arquivo.csv", dtype={'VENDAS': str})

df.columns = df.columns.str.strip()

for _, row in df.iterrows():
    destinatario = {
        "email": row["EMAIL"],
        "vendas": str(row["VENDAS"]).zfill(10)  # Nome do arquivo da coluna "VENDAS" com 10 dígitos
    }

    # Criar e-mail
    email = outlook.CreateItem(0)

    # Configurar informações do e-mail
    email.To = destinatario["email"]
    email.Subject = "Boleto Mensal - UnimedRV"
    email.HTMLBody = """
        <p>Olá, espero que este e-mail o(a) encontre bem!</p>

        <p>Segue em anexo o boleto mensal do plano de saúde da Unimed Rio Verde.</p>

        <p>Por gentileza conferir os dados de seu boleto.</p>

        <p>Atenciosamente,</p>

        <p>Departamento Financeiro.</p>
    """

    # Construir o caminho completo para o arquivo PDF
    pdf_path = rf'C:\boletos\{destinatario["vendas"]}.pdf'

    # Verificar se o arquivo PDF existe
    if os.path.exists(pdf_path):
        # Anexar o arquivo PDF correspondente
        email.Attachments.Add(pdf_path)
    else:
        with open(log_file_path, 'a') as log_file:
            log_file.write(f"Arquivo PDF não encontrado para {destinatario['email']} - {datetime.datetime.now()}\n")
        print(f"Arquivo PDF não encontrado para {destinatario['email']}")

    # Enviar o e-mail
    email.Send()
    print(f"E-mail enviado para {destinatario['email']}")

    # Registrar o envio no arquivo de log
    with open(log_file_path, 'a') as log_file:
        log_file.write(f"E-mail enviado para {destinatario['email']} - {datetime.datetime.now()}\n")

with open(log_file_path, 'a') as log_file:
    log_file.write(f"--- Fim do Log ({datetime.datetime.now()}) ---\n")

print("Todos os e-mails foram enviados!")
