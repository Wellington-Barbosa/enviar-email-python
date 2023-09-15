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

# Ler os dados do CSV usando Pandas
df = pd.read_csv(r"C:\beneficiarios\novo_arquivo.csv", dtype={'VENDAS': str})

df.columns = df.columns.str.strip()

for _, row in df.iterrows():
    destinatario = {
        "email": row["EMAIL"],
        "vendas": str(row["VENDAS"]).zfill(10)  # Nome do arquivo da coluna "VENDAS" com 10 dígitos
        #"vendas": row["VENDAS"]  # Nome do arquivo da coluna "VENDAS"
    }

    # Criar e-mail
    email = outlook.CreateItem(0)

    # Configurar informações do e-mail
    email.To = destinatario["email"]
    email.Subject = "Boleto Mensal - UnimedRV"
    email.HTMLBody = """
        <p>Olá,</p> 
        
        <p>Espero encontrá-lo(a)</p>

        <p>Segue em anexo o boleto simples do contrato de plano de saúde da Unimed Rio Verde.</p>

        <p><b>Esse e-mail não recebe retorno, por gentileza, não responder.</b></p>

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
        print(f"Arquivo PDF não encontrado para {destinatario['email']}")

    # Enviar o e-mail
    email.Send()
    data_hora = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    mensagem = f"{data_hora} - E-mail enviado para {destinatario['email']}"
    print(mensagem)
    ## print(f"E-mail enviado para {destinatario['email']}")

    # Salvar a mensagem em um arquivo de log
    log_path = r'C:\beneficiarios\log.txt'
    with open(log_path, 'a') as log_file:
        log_file.write(mensagem + '\n')

print("Todos os e-mails foram enviados!")
