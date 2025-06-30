import pandas as pd
import pathlib
import win32com.client as win32
import os

# Caminho para os dados
emails_path = r"C:\Users\Rafael\Desktop\PYTHONARQUIVOS\Projetos\Projeto AutomacaoIndicadores\Bases de Dados\Emails.xlsx"
lojas_path = r"C:\Users\Rafael\Desktop\PYTHONARQUIVOS\Projetos\Projeto AutomacaoIndicadores\Bases de Dados\Lojas.csv"
vendas_path = r"C:\Users\Rafael\Desktop\PYTHONARQUIVOS\Projetos\Projeto AutomacaoIndicadores\Bases de Dados\Vendas.xlsx"

# Verificando se o arquivo existe antes de carregá-lo
if os.path.exists(emails_path):
    emails = pd.read_excel(emails_path)
else:
    print(f"O arquivo de emails não foi encontrado em: {emails_path}")

if os.path.exists(lojas_path):
    lojas = pd.read_csv(lojas_path, encoding='latin1', sep=';')
else:
    print(f"O arquivo de lojas não foi encontrado em: {lojas_path}")

if os.path.exists(vendas_path):
    vendas = pd.read_excel(vendas_path)
else:
    print(f"O arquivo de vendas não foi encontrado em: {vendas_path}")

# Continue com o restante do seu código aqui
# Importa bases de dados
emails = pd.read_excel(emails_path)
lojas = pd.read_csv(lojas_path, encoding='latin1', sep=';')
vendas = pd.read_excel(vendas_path)

# Incluir nome da loja em vendas
vendas = vendas.merge(lojas, on='ID Loja')

# Dicionário de lojas
discionario_lojas = {}
for loja in lojas['Loja']:
    discionario_lojas[loja] = vendas.loc[vendas['Loja'] == loja, :]

# Dia indicador
dia_indicador = vendas['Data'].max()

# Caminho para o backup
caminho_backup = pathlib.Path(
    r"C:\Users\Rafael\Desktop\PYTHONARQUIVOS\Projetos\Projeto AutomacaoIndicadores\Backup Arquivos Lojas")

# Verifica se o diretório existe, caso contrário, cria
if not caminho_backup.exists():
    print(f"O diretório {caminho_backup} não existe. Criando o diretório.")
    caminho_backup.mkdir(parents=True)  # Cria o diretório

# Listando arquivos no diretório
try:
    arquivos_pasta_backup = caminho_backup.iterdir()
    lista_nomes_backup = [arquivo.name for arquivo in arquivos_pasta_backup]
except Exception as e:
    print(f"Ocorreu um erro ao listar arquivos: {e}")

# Salvar os arquivos de Excel
for loja in discionario_lojas:
    # Cria nova pasta se não existir
    nova_pasta = caminho_backup / loja
    if not nova_pasta.exists():
        print(f"Criando a pasta para a loja: {loja}")
        nova_pasta.mkdir(parents=True)

    nome_arquivo = f'{dia_indicador.month}_{dia_indicador.day}_{loja}.xlsx'
    local_arquivo = caminho_backup / loja / nome_arquivo
    discionario_lojas[loja].to_excel(local_arquivo)

# Meta de faturamento
meta_faturamento_dia = 1000
meta_faturamento_ano = 1650000
meta_qtdeprodutos_dia = 4
meta_qtdeprodutos_ano = 120
meta_ticketmedio_dia = 500
meta_ticketmedio_ano = 500

for loja in discionario_lojas:
    vendas_loja = discionario_lojas[loja]
    vendas_loja_dia = vendas_loja.loc[vendas_loja['Data'] == dia_indicador, :]

    # Faturamento
    faturamento_ano = vendas_loja['Valor Final'].sum()
    faturamento_dia = vendas_loja_dia['Valor Final'].sum()

    # Diversidade de produtos
    qtde_produtos_ano = len(vendas_loja['Produto'].unique())
    qtde_produtos_dia = len(vendas_loja_dia['Produto'].unique())

    # Ticket médio
    valor_venda = vendas_loja.groupby('Código Venda').sum(numeric_only=True)
    ticket_medio_ano = valor_venda['Valor Final'].mean()

    valor_venda_dia = vendas_loja_dia.groupby(
        'Código Venda').sum(numeric_only=True)
    ticket_medio_dia = valor_venda_dia['Valor Final'].mean()

    outlook = win32.Dispatch('outlook.application')

    nome = emails.loc[emails['Loja'] == loja, 'Gerente'].values[0]
    mail = outlook.CreateItem(0)
    mail.To = emails.loc[emails['Loja'] == loja, 'E-mail'].values[0]
    mail.Subject = f'OnePage Dia {dia_indicador.day}/{dia_indicador.month} - Loja {loja}'

    # Checagem das metas e atribuição das cores
    if faturamento_dia >= meta_faturamento_dia:
        cor_fat_dia = 'green'
    else:
        cor_fat_dia = 'red'
    if faturamento_ano >= meta_faturamento_ano:
        cor_fat_ano = 'green'
    else:
        cor_fat_ano = 'red'
    if qtde_produtos_dia >= meta_qtdeprodutos_dia:
        cor_qtde_dia = 'green'
    else:
        cor_qtde_dia = 'red'
    if qtde_produtos_ano >= meta_qtdeprodutos_ano:
        cor_qtde_ano = 'green'
    else:
        cor_qtde_ano = 'red'
    if ticket_medio_dia >= meta_ticketmedio_dia:
        cor_ticket_dia = 'green'
    else:
        cor_ticket_dia = 'red'
    if ticket_medio_ano >= meta_ticketmedio_ano:
        cor_ticket_ano = 'green'
    else:
        cor_ticket_ano = 'red'

    # Corpo do e-mail
    mail.HTMLBody = f'''
    <p>Bom dia, {nome}</p>
    <p>O resultado de ontem <strong>({dia_indicador.day}/{dia_indicador.month})</strong> da <strong>Loja {loja}</strong> foi:</p>
    <table>
    <tr>
        <th>Indicador</th>
        <th>Valor Dia</th>
        <th>Meta Dia</th>
        <th>Cenário Dia</th>
    </tr>
    <tr>
        <td>Faturamento</td>
        <td style="text-align: center">R${faturamento_dia:.2f}</td>
        <td style="text-align: center">R${meta_faturamento_dia:.2f}</td>
        <td style="text-align: center"><font color="{cor_fat_dia}">◙</td>
    </tr>
    <tr>
        <td>Diversidade de Produtos</td>
        <td style="text-align: center">{qtde_produtos_dia}</td>
        <td style="text-align: center">{meta_qtdeprodutos_dia}</td>
        <td style="text-align: center"><font color="{cor_qtde_dia}">◙</td>
    </tr>
    <tr>
        <td>Ticket Médio</td>
        <td style="text-align: center">R${ticket_medio_dia:.2f}</td>
        <td style="text-align: center">R${meta_ticketmedio_dia:.2f}</td>
        <td style="text-align: center"><font color="{cor_ticket_dia}">◙</td>
    </tr>
    </table>
    <br>
    <table>
    <tr>
        <th>Indicador</th>
        <th>Valor Ano</th>
        <th>Meta Ano</th>
        <th>Cenário Ano</th>
    </tr>
    <tr>
        <td>Faturamento</td>
        <td style="text-align: center">R${faturamento_ano:.2f}</td>
        <td style="text-align: center">R${meta_faturamento_ano:.2f}</td>
        <td style="text-align: center"><font color="{cor_fat_ano}">◙</td>
    </tr>
    <tr>
        <td>Diversidade de Produtos</td>
        <td style="text-align: center">{qtde_produtos_ano}</td>
        <td style="text-align: center">{meta_qtdeprodutos_ano}</td>
        <td style="text-align: center"><font color="{cor_qtde_ano}">◙</td>
    </tr>
    <tr>
        <td>Ticket Médio</td>
        <td style="text-align: center">R${ticket_medio_ano:.2f}</td>
        <td style="text-align: center">R${meta_ticketmedio_ano:.2f}</td>
        <td style="text-align: center"><font color="{cor_ticket_ano}">◙</td>
    </tr>
    </table>
    <p>Segue em anexo a planilha com todos os dados para mais detalhes.</p>
    <p>Qualquer dúvida estou à disposição.</p>
    <p>Att., Rafael</p>
    '''

    # Anexos (pode colocar quantos quiser)
    attachment = pathlib.Path.cwd() / caminho_backup / loja / \
        f'{dia_indicador.month}_{dia_indicador.day}_{loja}.xlsx'
    mail.Attachments.Add(str(attachment))

    try:
        mail.Send()
        print(f'E-mail da Loja {loja} enviado')
    except Exception as e:
        print(f'Erro ao enviar e-mail para Loja {loja}: {e}')

# Salvar faturamento anual
faturamento_lojas = vendas.groupby(
    'Loja')[['Loja', 'Valor Final']].sum(numeric_only=True)
faturamento_lojas_ano = faturamento_lojas.sort_values(
    by='Valor Final', ascending=False)

nome_arquivo = f'{dia_indicador.month}_{dia_indicador.day}_Ranking Anual.xlsx'
faturamento_lojas_ano.to_excel(caminho_backup / nome_arquivo)

# Salvar faturamento do dia
vendas_dia = vendas.loc[vendas['Data'] == dia_indicador, :]
faturamento_lojas_dia = vendas_dia.groupby(
    'Loja')[['Loja', 'Valor Final']].sum(numeric_only=True)
faturamento_lojas_dia = faturamento_lojas_dia.sort_values(
    by='Valor Final', ascending=False)

nome_arquivo = f'{dia_indicador.month}_{dia_indicador.day}_Ranking Dia.xlsx'
faturamento_lojas_dia.to_excel(caminho_backup / nome_arquivo)

# Envio do e-mail para a diretoria
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = emails.loc[emails['Loja'] == 'Diretoria', 'E-mail'].values[0]
mail.Subject = f'Ranking Dia {dia_indicador.day}/{dia_indicador.month}'
mail.HTMLBody = f'''
Prezados, bom dia

Melhor loja do Dia em Faturamento: Loja {faturamento_lojas_dia.index[0]} com Faturamento R${faturamento_lojas_dia.iloc[0, 0]:.2f}.
Pior loja do Dia em Faturamento: Loja {faturamento_lojas_dia.index[-1]} com Faturamento R${faturamento_lojas_dia.iloc[-1, 0]:.2f}.

Melhor loja do Ano em Faturamento: Loja {faturamento_lojas_ano.index[0]} com Faturamento R${faturamento_lojas_ano.iloc[0, 0]:.2f}.
Pior loja do Ano em Faturamento: Loja {faturamento_lojas_ano.index[-1]} com Faturamento R${faturamento_lojas_ano.iloc[-1, 0]:.2f}.

Segue em anexo os rankings do ano e do dia para todas as lojas.

Qualquer dúvida estou à disposição.

Att.,
Rafael
'''

# Anexos (pode colocar quantos quiser)
attachment = pathlib.Path.cwd() / caminho_backup / \
    f'{dia_indicador.month}_{dia_indicador.day}_Ranking Anual.xlsx'
mail.Attachments.Add(str(attachment))
attachment = pathlib.Path.cwd() / caminho_backup / \
    f'{dia_indicador.month}_{dia_indicador.day}_Ranking Dia.xlsx'
mail.Attachments.Add(str(attachment))

mail.Send()
print('E-mail da Diretoria enviado.')
