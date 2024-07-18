#Importa a biblioteca pandas, usada para manipulação e análise de dados, especialmente em estruturas de dados como DataFrames.
import pandas as pd

#Importa a biblioteca numpy, essencial para operações numéricas e manipulação de arrays multidimensionais.
import numpy as np

#Importa a biblioteca plotly.express, uma API de alto nível da Plotly para criar gráficos interativos de forma simples.
import plotly.express as px


# Importa a biblioteca smtplib com o alias 'smt', que permite enviar e-mails usando o protocolo SMTP.
# Importa a biblioteca email.message com o alias 'em', que fornece a classe Message para criar objetos de mensagem de e-mail.
import smtplib as smt
import email.message as em

# Importa a função format_currency da biblioteca babel.numbers com o alias 'fc', usada para formatar valores monetários de acordo com as convenções locais.
from babel.numbers import format_currency as fc


#Carregando dados de vendas a partir de um arquivo Excel usando pandas, e fazendo a leitura das primeiras cinco linhas do dataset como amostra
df_vendas = pd.read_excel(r'C:\Users\Ana Carolina Meliti\OneDrive\Área de Trabalho\case-vendas\vendas.xlsx')
df_vendas.head(5)

### CONTROLE DE ESTOQUE

#Agrupando os dados de vendas por produto e somando as quantidades vendidas
product_distinct = df_vendas.groupby('Produto').sum()
#Selecionando apenas a coluna 'Quantidade' do DataFrame resultante e selecionando valores em ordem decrescente
product_distinct = product_distinct[['Quantidade']].sort_values(by='Quantidade', ascending=False)
# Formata os valores da coluna 'Quantidade' para incluir pontos como separadores de milhar.
product_distinct['Quantidade'] = product_distinct['Quantidade'].apply(lambda x: "{:,.0f}".format(x).replace(',', '.'))

# Converte o DataFrame product_distinct para uma string HTML, removendo os índices e adicionando estilos CSS para centralizar o texto
# e definir a cor e o fundo das células no corpo da tabela.
prodist_html = product_distinct.to_html(index=False, justify='center', border=0).replace('<tbody>', '<tbody style= "text-align:center; color: #494949;background: #EBEBEB">')

# Exibe o DataFrame product_distinct no formato padrão de DataFrame do pandas.
display(product_distinct)
# Exibe a string HTML formatada da tabela product_distinct.
display(prodist_html)

### PRODUTOS MAIS VENDIDOS (R$)

# Cria uma nova coluna 'Faturamento' no DataFrame df_vendas, calculando o faturamento de cada venda como a quantidade vendida multiplicada pelo valor unitário do produto.
df_vendas['Faturamento'] = df_vendas['Quantidade'] * df_vendas['Valor Unitário']

# Agrupa o DataFrame df_vendas pela coluna 'Produto' e soma os valores das colunas numéricas para cada grupo de produtos.
product_faturamento = df_vendas.groupby('Produto').sum()
# Seleciona a coluna 'Faturamento' do DataFrame resultante, ordena os produtos pelo faturamento em ordem decrescente,
# e redefine os índices para que a coluna de produto volte a ser uma coluna regular ao invés de um índice.
product_faturamento = product_faturamento[['Faturamento']].sort_values(by='Faturamento', ascending=False).reset_index()
# Formata os valores da coluna 'Faturamento' para o formato de moeda brasileira (BRL) usando a função fc do babel.
product_faturamento['Faturamento'] = product_faturamento['Faturamento'].apply(lambda x: fc(x, 'BRL', locale='pt_BR'))

# Converte o DataFrame product_faturamento para uma string HTML, removendo os índices e adicionando estilos CSS para centralizar o texto
# e definir a cor e o fundo das células no corpo da tabela.
profat_html = product_faturamento.to_html(index=False, justify='center', border=0).replace('<tbody>', '<tbody style= "text-align:center; color: #494949;background: #EBEBEB">')


# Exibe o DataFrame product_faturamento no formato padrão de DataFrame do pandas.
display(product_faturamento)
# Exibe a string HTML formatada da tabela product_faturamento.
display(profat_html)

#### LOJAS/ESTADOS QUE MAIS VENDERAM (R$)

# Agrupa o DataFrame df_vendas pela coluna 'Loja' e soma os valores das colunas numéricas para cada grupo de lojas.
# O argumento 'numeric_only=True' garante que apenas as colunas numéricas sejam somadas.
faturamento_loja = df_vendas.groupby('Loja').sum(numeric_only=True)
# Seleciona a coluna 'Faturamento' do DataFrame resultante, ordena as lojas pelo faturamento em ordem decrescente,
# e redefine os índices para que a coluna de loja volte a ser uma coluna regular ao invés de um índice.
faturamento_loja = faturamento_loja[['Faturamento']].sort_values(by='Faturamento', ascending=False).reset_index()
# Formata os valores da coluna 'Faturamento' para o formato de moeda brasileira (BRL) usando a função fc do babel.
faturamento_loja['Faturamento'] = faturamento_loja['Faturamento'].apply(lambda x: fc(x, 'BRL', locale='pt_BR'))


# Converte o DataFrame faturamento_loja para uma string HTML, removendo os índices e adicionando estilos CSS para centralizar o texto
# e definir a cor e o fundo das células no corpo da tabela.
fatloja_html = product_faturamento.to_html(index=False, justify='center', border=0).replace('<tbody>', '<tbody style= "text-align:center; color: #494949;background: #EBEBEB">')

# Exibe o DataFrame faturamento_loja no formato padrão de DataFrame do pandas.
display(faturamento_loja)
# Exibe a string HTML formatada da tabela faturamento_loja.
display(fatloja_html)


### TICKET MÉDIO

# Cria uma nova coluna 'Ticket Médio' no DataFrame df_vendas, que é simplesmente uma cópia da coluna 'Valor Unitário'.
# Isso implica que o valor unitário do produto é considerado o ticket médio para cada transação.
df_vendas['Ticket Médio'] = df_vendas['Valor Unitário']

# Agrupa o DataFrame df_vendas pela coluna 'Loja' e calcula a média das colunas numéricas para cada grupo de lojas.
# O argumento 'numeric_only=True' garante que apenas as colunas numéricas sejam consideradas na média.
ticket_medio = df_vendas.groupby('Loja').mean(numeric_only=True)
# Seleciona a coluna 'Ticket Médio' do DataFrame resultante, ordena as lojas pelo ticket médio em ordem decrescente,
# e redefine os índices para que a coluna de loja volte a ser uma coluna regular ao invés de um índice.
ticket_medio = ticket_medio[['Ticket Médio']].sort_values('Ticket Médio', ascending=False).reset_index()
# Formata os valores da coluna 'Ticket Médio' para o formato de moeda brasileira (BRL) usando a função fc do babel.
ticket_medio['Ticket Médio'] = ticket_medio['Ticket Médio'].apply(lambda x: fc(x, 'BRL', locale='pt_BR'))

# Converte o DataFrame ticket_medio para uma string HTML, removendo os índices e adicionando estilos CSS para centralizar o texto
# e definir a cor e o fundo das células no corpo da tabela.
tickmed_html = product_faturamento.to_html(index=False, justify='center', border=0).replace('<tbody>', '<tbody style= "text-align:center; color: #494949;background: #EBEBEB">')

# Exibe o DataFrame ticket_medio no formato padrão de DataFrame do pandas.
display(ticket_medio)
# Exibe a string HTML formatada da tabela ticket_medio.
display(tickmed_html)

### AUTOMATIZAÇÃO DO ENVIO DO E-MAIL PARA ÁREA RESPONSÁVEL

# Cria o corpo do e-mail usando uma string formatada (f-string), incluindo o HTML formatado das tabelas.
# O e-mail contém seções para faturamento por loja, quantidade vendida por produto, ticket médio por loja e estoque de produtos disponíveis.
#Mensagem de e-mail 

corpo_email = f"""
<h1><p color="#272727">Olá, Líder!</p></h1>

<p color="#272727">Esperamos que esteja bem, trouxe novidades aqui... </p>
<div color="#272727">Saiba que você é nosso <b>grande progatonista</b> nessa nossa história.</div>

<p color="#272727">Verifique as atualizações da meta mensal da sua equipe aqui:</p>

<p color="#272727"><b>Faturamento por loja</b></p>
<div>{fatloja_html}</div>

<p color="#272727"><b>Quantidade vendida por produto</b></p>
<div>{profat_html}</div>

<p color="#272727"><b>Ticket médio por loja:</b></p>
<div>{tickmed_html}</div>

<p color="#272727"><b>Estoque de produtos disponíveis:</b></p>
<div>{prodist_html}</div>

<p>
<div color="#272727">Conte sempre conosco!</div>
<div color="#272727"><b>Atenciosamente,</b></div>
<div color="#272727">Equipe planejamento.</div>
"""

#Configurações de envio de e-mail 
# Define o assunto do e-mail.
msg = em.Message()

# Define o remetente e o destinatário do e-mail.
msg['Subject'] = "Relatório atualizado de metas" 
msg['From'] = 'customerautoenterprise@gmail.com'
msg['To'] = 'todos@gmail.com' 
password = '' # Define a senha do remetente (não é seguro deixar senhas hardcoded em código).
msg.add_header('Content-Type', 'text/html') # Define o corpo do e-mail como o HTML formatado criado anteriormente.
msg.set_payload(corpo_email )

# Inicializa a conexão com o servidor SMTP do Gmail e começa a comunicação TLS.
s = smt.SMTP('smtp.gmail.com: 587')
s.starttls()
# Login Credentials for sending the mail
s.login(msg['From'], password)
s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
print('Email enviado') # Imprime uma mensagem de confirmação após o envio do e-mail.