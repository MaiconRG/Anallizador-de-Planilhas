    # Importando uma biblioteca Pandas para gerenciamento de planilhas e dando um apelido de "pd"
    # Iportando o pndas pelo terminal utilizando o comando: pip install pandas
import pandas as pd 

    # Importando e apilidando como "win32"
    # Importando o sistema de envio de E-mail pelo comando: pip install pywin32 
import win32com.client as win32 

    # Importar a base de dados
    # Lê e adiciona atabela na variavel
tabela_vendas =  pd.read_excel('Vendas.xlsx') 

    # Visualizar a base de dados
    # Define para mostrar todas as colunas sem limite atraves do None
pd.set_option('display.max_columns', None) 
print(tabela_vendas)
print('-' * 50)

    # Faturamento por loja
    # Filtra a tabela por Id e valor e agrupa os ids somando os seus valores 
faturamento = tabela_vendas[['ID Loja','Valor Final']].groupby('ID Loja').sum() 
print(faturamento)
print('-' * 50)

    # Quantidade de produtos vendidos por loja
    # Filtrando tabela para mostrar o Id loja e quantidade
quantidade = tabela_vendas[['ID Loja','Quantidade']].groupby('ID Loja').sum() 
print(quantidade)
print('-' * 50)

    # Ticket médio por produto em cada loja
    # Realiza o calculo das tabelas e tranforma em uma tabela
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame() 
    # Renomeia o ticket Medio de "0" para "Ticket Médio"
ticket_medio = ticket_medio.rename(columns={0:"Ticket Médio"}) 
print('Ticket Medio')
print(ticket_medio)

    # Enviar um email com relatório
    # Abre o APP outlook
outlook = win32.Dispatch('outlook.application') 
    # Cria um Email
mail =  outlook.CreateItem(0) 
    # Local de Destino do E-mail
mail.To = 'mr.wingretschmann@outlook.com' 
    # Assunto do E-mail
mail.Subject = 'Base de Dados' 

    # As ''' ''' Servem para poder escrever livremente em varias linhas
    # O Corpo é formado por uma estrutura HTML
    # o f antes de uma string serve para poder chamar uma variavel atraves de {}
    # Atraves do pandas podemos comverter arquivos em tabela e html atraves do .to_frame() e .to_html()
    # Formatação dos numers atraves do  NomeColuna:'R${:,.2f}'.format
mail.HTMLBody = f''' 
<h3> Prezados Colaboradores Segue os valores de Vendas de Cada Loja </h3>

<p>
<h4> Faturamento </h4>
{faturamento.to_html(formatters={'Valor Final':'R${:,.2f}'.format})} 
</p>

<p>
<h4> Quantidade Vendida </h4>
{quantidade.to_html()}
</p>

<p>
<h4> Ticket Médio Dos Produtos de Cada Loja </h4>
{ticket_medio.to_html(formatters={'Ticket Médio':'R${:,.2f}'.format})}
</p>

<p> Obrigado Pela Atenção, qualquer dúvida estou a disposição, Atenciosamente Maicon </p>
''' 
    # Realiza o Envio
mail.Send() 
    # Confirma que o E-mail foi Enviado co sucesso
print('Enviado') 