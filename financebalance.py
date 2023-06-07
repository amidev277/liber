# Importar a tabela de finanças pessoais 
import pandas as pd
import numpy as np
import win32com.client as win32
#tabela_vendas=pd.read_excel('projetos.xlsx')

# Visualizar a base de dados 
#pd.set_option('display.max_columns', None)
tabela_vendas=pd.read_excel("projetos.xlsx", sheet_name="Financa" ,skiprows=32,nrows=3,usecols=[20,21])
#tabela_vendas=tabela_vendas
tabela_vendas=tabela_vendas.rename(columns={'Unnamed: 20':'Descrição'})
tabela_vendas=tabela_vendas.rename(columns={'Unnamed: 21':'Valor'})
tabela_vendas = tabela_vendas.replace('<td>Balanço</td>', '<td style="color: red;">Balanço</td>')
tabela_vendas= tabela_vendas.replace('<table>', '<table style="color: blue;">')
tabela_vendas_cleaned = tabela_vendas.replace(np.nan, '', regex=True)
html_table = tabela_vendas_cleaned.to_html()
html_table = html_table.replace('<table>', '<table style="color: blue;">')
html_table = html_table.replace('<th>', '<th style="color: red;">')
html_table = html_table.replace('<td>', '<td style="color: green;">')
html_table = tabela_vendas_cleaned.to_html()
print(tabela_vendas_cleaned)
print('-' *50)
# Tabela resumo 
resumo=pd.read_excel("projetos.xlsx", sheet_name="Financa" ,skiprows=17,nrows=37,usecols="B:N")
resumo=resumo.rename(columns={'Unnamed: 2':'Janeiro'})
resumo=resumo.rename(columns={'Unnamed: 3':'Fevereiro'})
resumo=resumo.rename(columns={'Unnamed: 4':'Março'})
resumo=resumo.rename(columns={'Unnamed: 5':'Abril'})
resumo=resumo.rename(columns={'Unnamed: 6':'Maio'})
resumo=resumo.rename(columns={'Unnamed: 7':'Junho'})
resumo=resumo.rename(columns={'Unnamed: 8':'Julho'})
resumo=resumo.rename(columns={'Unnamed: 9':'Agosto'})
resumo=resumo.rename(columns={'Unnamed: 10':'Setembro'})
resumo=resumo.rename(columns={'Unnamed: 11':'Outubro'})
resumo=resumo.rename(columns={'Unnamed: 12':'Novembro'})
resumo=resumo.rename(columns={'Unnamed: 13':'Dezembro'})
resumo_cleaned = resumo.replace(np.nan, '', regex=True)
html_table = resumo_cleaned.to_html()
html_table = html_table.replace('<table>', '<table style="color: blue;">')
html_table = html_table.replace('<th>', '<th style="color: red;">')
html_table = html_table.replace('<td>', '<td style="color: green;">')
html_table = resumo_cleaned.to_html()
print(resumo_cleaned)

import smtplib
import email.message

def enviar_email(tabela_vendas_cleaned):  
    corpo_email = f"""
    <p></p>
    <p></p>
    <p> Estado financeiro: {tabela_vendas_cleaned.to_html(formatters={'Valor': '€{:,.2f}'.format})}</p>
    
    <p></p>
    <p> Relatório financeiro: {resumo_cleaned.to_html(formatters={'Janeiro, Fevereiro, Março, Abril, Maio, Junho, Julho, Agosto, Setembro, Outubro, Novembro, Dezembro': '€{:,.2f}'.format})}</p>
     
    <p>Tudo aquilo que cuidamos cresce.</p>
    <p></p>
    """

    msg = email.message.Message()
    msg['Subject'] = "Oganização financeira"
    msg['From'] = 'alisboagamer7@gmail.com'
    msg['To'] = 'alisboagamer7@gmail.com'
    password = 'jfqtkdryqgvdghih' 
    msg.add_header('Content-Type', 'text/html')
    msg.set_payload(corpo_email )

    s = smtplib.SMTP('smtp.gmail.com: 587')
    s.starttls()
    # Login Credentials for sending the mail
    s.login(msg['From'], password)
    s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
    print('Email enviado')

#if cotacao < 5.10:
enviar_email(tabela_vendas_cleaned)

# Deploy com heroku  