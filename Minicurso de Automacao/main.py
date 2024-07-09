import pandas as pd
import win32com.client as win32

#importar base de dados
tabela_vendas = pd.read_excel ('Vendas.xlsx')

#visualizar a base de dados
pd.set_option ('display.max_columns', None)
print(tabela_vendas)

#faturamento por loja, filtro e agrupamento dos dados, com soma
faturamento = tabela_vendas[['ID Loja','Valor Final']].groupby('ID Loja').sum()
print(faturamento)

#quantidade de produtos vendidos por loja, filtro agrupamento e soma
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)
print('-' * 50)

#ticket medio por produto em cada loja, to frame transforma em tabela
ticket_medio = (faturamento['Valor Final']/ quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Medio'})
print(ticket_medio)

#enviar um email com relatório
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'quesiavirginia2806@gmail.com'
mail.Subject = 'Relatório de Vendas por Loja'
mail.HTMLBody = f''' 
<p>Prezados,</p>

<p>Segue o Relatório de Vendas por cada loja.</p>

<p>Faturamento:</p> 
{faturamento.to_htm(formaters= {'Valor Final': 'R${: ,.2f}'.format})}

<p>Quantidade Vendida:</p> 
{quantidade.to_html()}

<p>Ticket Medio dos Produtos em cada Loja:</p> 
{ticket_medio.to_html()}

<p>Qualquer duvida estou a disposição.</p>

Att, Quesia Silva </p>

'''

mail.Send ()

#abrir e fechar paragrafo html, usar tag <p> e </p>