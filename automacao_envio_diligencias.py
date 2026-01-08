import pandas as pd
import win32com.client as win32

#importar base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')

#visualizar a tabela
pd.set_option('display.max_columns', None)

#faturamento por loja
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()

#quantidade de produtos vendidos por loja
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()

# ticket médio por produto em cada loja
ticke_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()

print(ticke_medio)

#enviar email
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'seuemail@outlook.com'
mail.Subject = 'Relatório de Vendas por Loja'
mail.HTMLBody = '''Prezado(a), boa tarde!
Segue o relatório de vendas por cada loja

Faturamento: 
{}

Quantidade vendida:
{}

Ticket médio dos produtos em casa loja:
{}

Atenciosamente, Júlio Cézar
'''

mail.send()
