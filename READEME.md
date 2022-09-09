            ----------------------- | ANALIZADOR DE DADOS | -----------------------

------------ | [OBJETIVO] | ------------ 
Pegar os dados de uma planilha filtrando e porcessando, gerando um E-mail com o valor das vedas, quantidade e ticket médio.

------------ | [BIBLIOTECAS] | ------------ 
# pandas 
# pywin32 

------------ | [REQUISITOS] | ------------ 
# Conexão com a internet
# Outlook Instalado
# Planilha para ser analisada

------------ | [UTILZAÇÃO] | ------------ 
1° Defina o nome correto da planilha na linha 8
# "tabela_vendas =  pd.read_excel('NOME_DA_TABELA.FORMATO')"

2° Defina o nome das colunas da tabela nas linhas 16 e 21
# faturamento = tabela_vendas[[ 'NOME_COLUNA1','NOME_COLUNA2']].groupby('ID Loja').sum()

3° Defina o destinatário do E-mail na linha 34
# mail.To = 'EMAIL_DESTINO@EMAIL.COM'

4° Defina o título do E-mail na linha 35
# mail.Subject = 'TITULO'

5° Defina a estrutura do E-mail na linha 42 => 61
# O Corpo do e-mail é estruturado em HTML, realize as edições necessarias para o envio