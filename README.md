# Documentação do Script de Envio de E-mails e Boletos

Este script Python automatiza o processo de envio de e-mails contendo boletos mensais da Unimed Rio Verde para os beneficiários, com base em um arquivo CSV de entrada. Ele utiliza a biblioteca `win32com` para integração com o Outlook e a biblioteca `pandas` para manipulação de dados. Abaixo estão detalhadas as principais funcionalidades e o funcionamento do script.

## Pré-Requisitos

Antes de usar este script, é necessário ter os seguintes pré-requisitos configurados:

1. **Microsoft Outlook**: O Microsoft Outlook deve estar instalado e configurado com uma conta de e-mail válida.
2. **Arquivo CSV de Beneficiários**: Você precisa ter um arquivo CSV contendo informações dos beneficiários, incluindo os endereços de e-mail e os números de vendas. Certifique-se de que o arquivo esteja no formato correto, com ";" como separador.
3. **Diretórios de Armazenamento**: Você deve criar os diretórios onde os boletos PDF serão armazenados (`C:\boletos`) e onde o log de envio será salvo (`C:\beneficiarios`).

## Funcionamento do Script

O script realiza as seguintes etapas:

1. **Integração com o Outlook**: Ele cria uma integração com o Outlook usando a biblioteca `win32com`.

2. **Leitura do Arquivo CSV de Beneficiários**: O script carrega o arquivo CSV contendo informações dos beneficiários em um DataFrame do Pandas. Certifique-se de ajustar o caminho do arquivo de entrada (`caminho_do_arquivo`) de acordo com a localização do seu arquivo CSV.

3. **Conversão e Salvamento do DataFrame**: O script converte o DataFrame para um novo arquivo CSV com "," como separador. Isso é necessário para garantir que os números de vendas estejam no formato correto. O novo arquivo é salvo no caminho especificado em `novo_caminho_do_arquivo`.

4. **Loop pelos Beneficiários**: O script percorre cada linha do DataFrame para cada beneficiário.

5. **Criação do E-mail**: Para cada beneficiário, ele cria um e-mail no Outlook com as informações necessárias, como destinatário, assunto e corpo do e-mail. O corpo do e-mail contém uma mensagem padrão.

6. **Anexação do Boleto PDF**: O script verifica se o arquivo PDF correspondente ao número de vendas do beneficiário existe no diretório `C:\boletos`. Se existir, ele anexa o arquivo PDF ao e-mail.

7. **Envio do E-mail**: O e-mail é enviado para o beneficiário.

8. **Registro no Log**: O script registra a data e hora do envio do e-mail no arquivo de log especificado em `log_path`.

9. **Conclusão do Processo**: Após o envio de todos os e-mails, o script exibe a mensagem "Todos os e-mails foram enviados!".

## Personalização

Você pode personalizar o corpo do e-mail, o formato do nome do arquivo PDF e outras configurações de acordo com suas necessidades, modificando o código diretamente.

Certifique-se de que todas as bibliotecas necessárias estejam instaladas no ambiente Python em que você pretende executar o script. Você pode instalar as bibliotecas ausentes usando o `pip`, se necessário.

Lembre-se de manter os arquivos de boletos no diretório `C:\boletos` e ajustar os caminhos dos diretórios conforme necessário.

**Observação**: Este script é uma simplificação e pode precisar de ajustes adicionais, dependendo dos detalhes específicos do seu ambiente e das necessidades de envio de e-mails. Certifique-se de testá-lo em um ambiente de desenvolvimento antes de usar em produção.