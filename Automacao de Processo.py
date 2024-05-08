#!/usr/bin/env python
# coding: utf-8

# ### Passo 1 - Importar Arquivos e Bibliotecas

# In[ ]:


import pandas as pd
import pathlib 
import smtplib
import email.message
import mimetypes
from email.message import EmailMessage
from email.mime.text import MIMEText


# In[ ]:


emails = pd.read_excel(r'Bases de Dados\Emails.xlsx')
lojas = pd.read_csv(r'Bases de Dados\Lojas.csv', encoding='latin1', sep=';') 
vendas = pd.read_excel(r'Bases de Dados\Vendas.xlsx')
display(emails)
display(lojas)
display(vendas)


# ### Passo 2 - Definir Criar uma Tabela para cada Loja e Definir o dia do Indicador

# In[ ]:


vendas = vendas.merge(lojas, on='ID Loja')
display(vendas)


# In[ ]:


dicionario_lojas = {}
for loja in lojas['Loja']:
    dicionario_lojas[loja] = vendas.loc[vendas['Loja'] == loja, :]

display(dicionario_lojas['Rio Mar Recife'])
display(dicionario_lojas['Shopping Vila Velha'])


# In[ ]:


dia_indicador = vendas['Data'].max()
print(dia_indicador)
print('{}/{}/{}'.format(dia_indicador.day, dia_indicador.month, dia_indicador.year))


# ### Passo 3 - Salvar a planilha na pasta de backup

# In[ ]:


#identificar se a pasta já exite
caminho_backup = pathlib.Path(r'Backup Arquivos Lojas')
arquivos_pasta_backup = caminho_backup.iterdir()

lista_nomes_backup = [arquivo.name for arquivo in arquivos_pasta_backup]

for loja in dicionario_lojas:
    if not loja in lista_nomes_backup:
        nova_pasta = caminho_backup / loja
        nova_pasta.mkdir()

#salvar dentro da pasta
    nome_arquivo = '{}_{}_{}_{}.xlsx'.format(dia_indicador.year, dia_indicador.month, dia_indicador.day, loja)
    local_arquivo = caminho_backup / loja / nome_arquivo    
    dicionario_lojas[loja].to_excel(local_arquivo)


# ### Passo 4 - Calcular o indicador para 1 loja

# In[ ]:


#definição de metas
meta_faturamento_dia = 1000
meta_faturamento_ano = 1650000
meta_qtdeprodutos_dia = 4
meta_qtdeprodutos_ano = 120
meta_ticketmedio_dia = 500
meta_ticketmedio_ano = 500


# In[ ]:


for loja in dicionario_lojas:
    vendas_loja_ano = dicionario_lojas[loja]
    vendas_loja_dia = vendas_loja_ano.loc[vendas_loja_ano['Data']==dia_indicador, :]

    #faturamento
    faturamento_ano = vendas_loja_ano['Valor Final'].sum()
    faturamento_dia = vendas_loja_dia['Valor Final'].sum()
    #print(faturamento_ano)
    #print(faturamento_dia)

    #diversidade de produtos
    qtde_produtos_ano = len(vendas_loja_ano['Produto'].unique())
    #print(qtde_produtos_ano)

    #qtde_produtos_dia = 
    qtde_produtos_dia = len(vendas_loja_dia['Produto'].unique())
    #print(qtde_produtos_dia)

    #ticket medio
    valor_venda_ano = vendas_loja_ano.groupby('Código Venda').sum(numeric_only=True)
    ticket_medio_ano = valor_venda_ano['Valor Final'].mean()
    #print(ticket_medio_ano)

    valor_venda_dia = vendas_loja_dia.groupby('Código Venda').sum(numeric_only=True)
    ticket_medio_dia = valor_venda_dia['Valor Final'].mean()
    #print(ticket_medio_dia)

    if faturamento_dia >= meta_faturamento_dia:
        cor_fat_dia = 'green'
    else:
        cor_fat_dia = 'red'
    if faturamento_ano >= meta_faturamento_ano:
        cor_fat_ano = 'green'
    else:
        cor_fat_ano = 'red'
    if qtde_produtos_dia >= meta_qtdeprodutos_dia:
        cor_qtde_dia = 'green'
    else:
        cor_qtde_dia = 'red'
    if qtde_produtos_ano >= meta_qtdeprodutos_ano:
        cor_qtde_ano = 'green'
    else:
        cor_qtde_ano = 'red'
    if ticket_medio_dia >= meta_ticketmedio_dia:
        cor_ticket_dia = 'green'
    else:
        cor_ticket_dia = 'red'
    if ticket_medio_ano >= meta_ticketmedio_ano:
        cor_ticket_ano = 'green'
    else:
        cor_ticket_ano = 'red'

    nome = emails.loc[emails['Loja']==loja, 'Gerente'].values[0]
    
    def enviar_emails():
        corpo_email = f'''
        <p>Bom dia, {nome} </p>

        <p>O resultado de ontem <strong>({dia_indicador.day}/{dia_indicador.month})</strong> da <strong>Loja {loja}</strong> foi:</p>

        <table>
          <tr>
            <th>Indicador</th>
            <th>Valor Dia</th>
            <th>Meta Dia</th>
            <th>Cenário Dia</th>
          </tr>
          <tr>
            <td>Faturamento</td>
            <td style="text-align: center">R${faturamento_dia:.2f}</td>
            <td style="text-align: center">R${meta_faturamento_dia:.2f}</td>
            <td style="text-align: center"><font color="{cor_fat_dia}">◙</font></td>
          </tr>
          <tr>
            <td>Diversidade de Produtos</td>
            <td style="text-align: center">{qtde_produtos_dia}</td>
            <td style="text-align: center">{meta_qtdeprodutos_dia}</td>
            <td style="text-align: center"><font color="{cor_qtde_dia}">◙</font></td>
          </tr>
          <tr>
            <td>Ticket Médio</td>
            <td style="text-align: center">R${ticket_medio_dia:.2f}</td>
            <td style="text-align: center">R${meta_ticketmedio_dia:.2f}</td>
            <td style="text-align: center"><font color="{cor_ticket_dia}">◙</font></td>
          </tr>
        </table>
        <br>
        <table>
          <tr>
            <th>Indicador</th>
            <th>Valor Ano</th>
            <th>Meta Ano</th>
            <th>Cenário Ano</th>
          </tr>
          <tr>
            <td>Faturamento</td>
            <td style="text-align: center">R${faturamento_ano:.2f}</td>
            <td style="text-align: center">R${meta_faturamento_ano:.2f}</td>
            <td style="text-align: center"><font color="{cor_fat_ano}">◙</font></td>
          </tr>
          <tr>
            <td>Diversidade de Produtos</td>
            <td style="text-align: center">{qtde_produtos_ano}</td>
            <td style="text-align: center">{meta_qtdeprodutos_ano}</td>
            <td style="text-align: center"><font color="{cor_qtde_ano}">◙</font></td>
          </tr>
          <tr>
            <td>Ticket Médio</td>
            <td style="text-align: center">R${ticket_medio_ano:.2f}</td>
            <td style="text-align: center">R${meta_ticketmedio_ano:.2f}</td>
            <td style="text-align: center"><font color="{cor_ticket_ano}">◙</font></td>
          </tr>
        </table>

        <p>Segue em anexo a planilha com todos os dados para mais detalhes.</p>

        <p>Qualquer dúvida estou à disposição.</p>
        <p>Att., Ricardo Ming</p>
        '''

        msg = EmailMessage()
        
        msg['Subject'] = "OnePage Dia {} - Loja {}".format(dia_indicador.strftime("%d/%m"), loja)
        msg['From'] = 'ricardoxming@gmail.com'
        msg['To'] = emails.loc[emails['Loja'] == loja, 'E-mail'].values[0]
        password = '****************'
        msg.add_header('Content-Type', 'text/html')
        msg.set_content(corpo_email, subtype='html')


        attachment = pathlib.Path.cwd() / caminho_backup / loja / f'{dia_indicador.year}_{dia_indicador.month}_{dia_indicador.day}_{loja}.xlsx'

        mime_type, encoding = mimetypes.guess_type(attachment.name)
        with open(attachment, 'rb') as fp:
            dados = fp.read()
        msg.add_attachment(dados, filename=f'{dia_indicador.month}_{dia_indicador.day}_{loja}.xlsx', maintype='application', subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet')


        s = smtplib.SMTP('smtp.gmail.com: 587')
        s.starttls()

        # Login Credentials for sending the mail
        s.login(msg['From'], password)
        s.send_message(msg)
    

    enviar_emails()
    print('E-mail da Loja {} enviado'.format(loja))


# ### Passo 7 - Criar ranking para diretoria

# In[ ]:


faturamento_lojas = vendas.groupby('Loja')[['Loja', 'Valor Final']].sum(numeric_only=True)
faturamento_lojas_ano = faturamento_lojas.sort_values(by='Valor Final', ascending=False)
display(faturamento_lojas)

nome_arquivo = '{}_{}_{}_Ranking Anual.xlsx'.format(dia_indicador.year, dia_indicador.month, dia_indicador.day)   
faturamento_lojas_ano.to_excel(r'Backup Arquivos Lojas\{}'.format(nome_arquivo))

vendas_dia = vendas.loc[vendas['Data']==dia_indicador,:]
faturamento_lojas_dia = vendas_dia.groupby('Loja')[['Loja', 'Valor Final']].sum(numeric_only=True)
faturamento_lojas_dia = faturamento_lojas_dia.sort_values(by='Valor Final', ascending=False)
display(faturamento_lojas_dia)

nome_arquivo = '{}_{}_{}_Ranking Dia.xlsx'.format(dia_indicador.year, dia_indicador.month, dia_indicador.day)   
faturamento_lojas_dia.to_excel(r'Backup Arquivos Lojas\{}'.format(nome_arquivo))


# ### Passo 8 - Enviar e-mail para diretoria

# In[ ]:


corpo_email = f'''
Prezados, bom dia

Melhor loja do Dia em Faturamento: Loja {faturamento_lojas_dia.index[0]} com Faturamento R${faturamento_lojas_dia.iloc[0, 0]:.2f}
Pior loja do Dia em Faturamento: Loja {faturamento_lojas_dia.index[-1]} com Faturamento R${faturamento_lojas_dia.iloc[-1, 0]:.2f}

Melhor loja do Ano em Faturamento: Loja {faturamento_lojas_ano.index[0]} com Faturamento R${faturamento_lojas_ano.iloc[0, 0]:.2f}
Pior loja do Ano em Faturamento: Loja {faturamento_lojas_ano.index[-1]} com Faturamento R${faturamento_lojas_ano.iloc[-1, 0]:.2f}


Segue em anexo os rankings do ano e do dia de todas as lojas.

Qualquer dúvida, estou à disposição.

Att.,
Ricardo Ming 
'''

msg = EmailMessage()
msg['Subject'] = "Ranking Dia {}".format(dia_indicador.strftime("%d/%m"))
msg['From'] = 'ricardoxming@gmail.com'
msg['To'] = emails.loc[emails['Loja'] == 'Diretoria', 'E-mail'].values[0]
password = 'kewswjfrbzafooto'
msg.add_header('Content-Type', 'text/html')
msg.set_content(corpo_email)


attachment = pathlib.Path.cwd() / caminho_backup / f'{dia_indicador.year}_{dia_indicador.month}_{dia_indicador.day}_Ranking Anual.xlsx'
mime_type, encoding = mimetypes.guess_type(attachment.name)
with open(attachment, 'rb') as fp:
    dados = fp.read()
msg.add_attachment(dados, filename=f'{dia_indicador.year}_{dia_indicador.month}_{dia_indicador.day}_Ranking Anual.xlsx', maintype='application', subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet')
s = smtplib.SMTP('smtp.gmail.com: 587')
s.starttls()

attachment = pathlib.Path.cwd() / caminho_backup / f'{dia_indicador.year}_{dia_indicador.month}_{dia_indicador.day}_Ranking Dia.xlsx'
mime_type, encoding = mimetypes.guess_type(attachment.name)
with open(attachment, 'rb') as fp:
    dados = fp.read()
msg.add_attachment(dados, filename=f'{dia_indicador.year}_{dia_indicador.month}_{dia_indicador.day}_Ranking Dia.xlsx', maintype='application', subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet')
s = smtplib.SMTP('smtp.gmail.com: 587')
s.starttls()




# Login Credentials for sending the mail
s.login(msg['From'], password)
s.send_message(msg)

#enviar_emails()
print('E-mail da diretoria enviado')

