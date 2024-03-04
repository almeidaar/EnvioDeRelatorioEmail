import pandas as pd
import win32com.client as win32

# importar a base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx') # este comando quer dizer que o pandas "pd." irá ler "read" o arquivo "Vendas ..." e armazenar na variável "tabela_vendas"

# visualizar a base de dados
pd.set_option('display.max_columns', None)
print (tabela_vendas) 

# calcular faturamento por loja
faturamento = tabela_vendas[['Produto', 'Preco Unitario']].groupby('Produto').sum()

#print(faturamento)
print(faturamento)

# calculr quantidade de produtos vendidos por loja
qntd_loja = tabela_vendas[['Produto', 'Quantidade Vendida']].groupby('Produto').sum()
print(qntd_loja)

print('-' * 50)

# calcular ticket médio por produto em cada loja
ticket_medio = (faturamento['Preco Unitario'] / qntd_loja['Quantidade Vendida']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
print(ticket_medio)

# enviar um email com relatório / o uso do "f" na linha 33, diz que dentro do texto as chaves "{}" conterão uma variável 
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'nsilva@meucashcard.com.br'
mail.Subject = 'Relatorio de Vendas por Loja'
mail.Body = 'Message Body'
mail.HTMLBody = f''' 
<p>Prezados,</p> 

<p>Segue o Relatório de Vendas da loja de São Paulo.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Preco Unitario': 'R${:,.2f}'.format})}

<p>Quantidade Vendida:</p>
{qntd_loja.to_html()}

<p>Ticket Médio dos Produtos em cada Loja:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida estou à disposição.</p>

<p>Atenciosamete, </p>
<p>    Nicholas Almeida</p>

''' 
mail.Send()

print('Email Enviado')