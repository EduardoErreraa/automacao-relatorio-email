
import pandas as pd
import win32com.client as win32

salesTable = pd.read_excel('Vendas.xlsx')

# remove column view limit
pd.set_option('display.max_columns', None )

# filter the table and group the data
income = salesTable[['ID Loja','Valor Final']].groupby('ID Loja').sum()
soldProduct = salesTable[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()

# convert to table
averageTicket = (income['Valor Final'] / soldProduct['Quantidade']).to_frame()
averageTicket = averageTicket.rename(columns={0: 'Ticket Médio'})

print(averageTicket)

# send the email
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'testePython@gmail.com'
mail.Subject = 'Sales Report by Store'
mail.HTMLBody = f'''

Table_01
{income.to_html()}

Table_02
{soldProduct.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

Table_03
{averageTicket.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

'''

mail.Send()