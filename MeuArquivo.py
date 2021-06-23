import pandas as pd
import win32com.client as win32

# importar a base de dados
tabela_vendas = pd.read_excel('vendas_por_loja.xlsx')

# visualizar a base de dados
pd.set_option('display.max_columns', None)
print(tabela_vendas)

# faturamento por loja
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)

# quantidade de produtos vendidos por loja
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)

# ticket médio de produto em cada  loja (média)
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
print(ticket_medio)

# envio de email com o relatorio
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'jarbaswssilva@gmail.com'
mail.Subject = 'Relatório de vendas por loja'
mail.HTMLBody = f'''
Prezados,

<p>Segue o Relatório de vendas por cada loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final':'R${:,.2f}'.format})}

<p>Quantidade vendida:</p>
{quantidade.to_html()}

<p>Ticket Médio dos produtos de cada loja:</p>
{ticket_medio.to_html(formatters={'Ticket Médio':'R${:,.2f}'.format})}

<p>Qualquer dúvida, estou a disposição.</p>

<p>Att,</p>
<p>JARBAS.</p>
 '''

mail.Send()

print('email enviado')
