import pandas as pd
import win32com.client as win32

tabela_vendas = pd.read_excel('Vendas.xlsx')
pd.set_option('display.max_columns', None)

faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'seuemail@outlook.com'
mail.Subject = 'Relatório de Vendas por Loja'
mail.HTMLBody = f'''Prezado(a), boa tarde!
Segue o relatório de vendas por cada loja

Faturamento: 
{faturamento}

Quantidade vendida:
{quantidade}

Ticket médio dos produtos em casa loja:
{ticket_medio}

Atenciosamente, Júlio Cézar
'''
mail.Send()
