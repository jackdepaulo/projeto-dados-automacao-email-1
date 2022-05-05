import pandas as pd
import win32com.client as win32

# -- ANÁLISE DE DADOS
#  importar a base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')

#  visualizar a base de dados
pd.set_option('display.max_columns', None)
print(tabela_vendas)

#  FATURAMENTO POR LOJA
print('-' * 25, 'FATURAMENTO', '-' * 25)
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)

#  quantidade de produtos vendidos por loja
print('-' * 25, 'QUANTIDADE', '-' * 25)
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)


#  ticket médio por produto em cada loja
print('-' * 25, 'TICKET MEDIO', '-' * 25)
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
print(ticket_medio)


#  --AUTOMAÇÃO-- enviar um email com o relatório

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'JJJ@GMAIL.COM'  # exemplo
mail.Subject = 'Relatório de Vendas'
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue o Relatório de Vendas por cada loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade Vendida:</p>
{quantidade.to_html()}

<p>Ticket médio dos Produtos em cada loja:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida estou à disposição.</p>
<p>Att., Jaqueline</p>

'''

mail.Send()
print('Email enviado')

