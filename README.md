Automacao Indicadores

🚀 O que esse projeto faz

Lê bases de vendas, lojas e e-mails.

Gera indicadores de desempenho (faturamento, ticket médio, diversidade de produtos).

Compara com metas diárias e anuais.

Gera planilhas individuais por loja com os dados.

Manda e-mails automáticos para cada gerente com os indicadores + planilha em anexo.

Envia um resumo para a diretoria com ranking de desempenho.

----------------------------------------------------------------------------------------------------------------

📁 Estrutura esperada dos arquivos

Bases de Dados/

Emails.xlsx          # Contém o nome da loja, e-mail e gerente

Lojas.csv            # Contém ID e nome das lojas

Vendas.xlsx          # Base de vendas com colunas como Data, Valor Final, Produto, Código Venda, etc.

----------------------------------------------------------------------------------------------------------------

🛠️ Requisitos

Python 3.x

pandas

pywin32 (pra controlar o Outlook via COM)

Planilhas bem organizadas (não faz milagre, só automatiza)

pip install pandas pywin32

----------------------------------------------------------------------------------------------------------------

🧠 Principais Lógicas

Junta os dados de vendas com as lojas.

Separa as vendas por loja.

Calcula os seguintes indicadores:

Faturamento diário / anual

Ticket médio diário / anual

Diversidade de produtos diário / anual

Compara os indicadores com metas definidas e atribui cores (verde ou vermelho).

Cria e-mails formatados com HTML para os gerentes (com planilha anexa).

Gera e envia um ranking geral para a diretoria (com anexos dos rankings).

----------------------------------------------------------------------------------------------------------------

📬 Exemplo de envio de e-mail

Cada gerente recebe um e-mail assim:

Assunto: OnePage Dia 30/06 - Loja Floripa

Bom dia, João
O resultado de ontem (30/06) da Loja Floripa foi:
[Faturamento, Ticket Médio, Produtos Diversificados...]
[✓] Planilha detalhada em anexo

E a diretoria recebe:

Assunto: 

Ranking Dia 30/06
Melhor loja do Dia: Loja A - R$5.000,00  
Pior loja do Dia: Loja B - R$300,00  

...

[Anexos: Ranking Dia, Ranking Anual]

----------------------------------------------------------------------------------------------------------------

⚠️ Observações

Os caminhos para os arquivos estão fixos no script. Ajuste conforme sua máquina.

Requer Outlook instalado e configurado (não funciona com Gmail, Thunderbird, etc.).

As metas estão definidas no código, mas podem ser externalizadas depois.

Rodou o script e ninguém recebeu nada? Provavelmente o Outlook tá de birra.

