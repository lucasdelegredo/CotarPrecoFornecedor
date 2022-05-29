import pandas as pd
import win32com.client as win32

cotacao = pd.read_excel('cotacoes.xlsx')
fornecedor = pd.read_excel('fornecedores.xlsx')


def tratar_cotacao(cotacoes,fornecedores):
    #tratar as informações e tirar o index
    cotacoes=cotacoes.set_index('Item')

    #print(cotacoes)
    #print(fornecedores)
    #print(cotacoes['Descrição RC'])


    ####Processar as informações####

    ####FILTRAR AS INFORMAÇÕES BASEADAS NO INPUT####
    grupo_fornecedor = input("Grupo de fornecedor: ")

    cotacoes_pendentes = (cotacoes.loc[
            (cotacoes['Descrição RC']==str(grupo_fornecedor)) 
                        & 
            (cotacoes['Status']=='PENDENTE')])

    fornecedor_email = (fornecedores.loc[fornecedores['Descrição RC']==str(grupo_fornecedor)])

    ####LISTA DE EMAIL DO RESPECTIVO FORNECEDOR RESPONSAVEL####
    fornecedor_listemail = fornecedor_email['Grupo de Emails'].values[0]  #PEGAR APENAS O EMAIL (sem índice)

    #print(fornecedor_listemail)

    #PEGAR OS 20 PRIMEIROS PARA TRABALHAR E MANDAR A COTAÇÃO
    cotacoes_pendentes = cotacoes_pendentes.head(20)

    #print(cotacoes_pendentes)

    ####TRATAR A INFORMAÇÃO QUE SERÁ ENVIADA PARA O FORNECEDOR PARA COTAR - não precisamos da coluna de 'Descrição da RC' e nem de 'Status' pois é apenas para o nosso controle####
    cotacao_enviar = cotacoes_pendentes.drop('Descrição RC',1)
    cotacao_enviar = cotacao_enviar.drop('Status',1)

    enviar_email(fornecedor_listemail,cotacao_enviar)

#CRIAR UMA FUNÇÃO PARA MANDAR O EMAIL E PASSAR DOIS PARAMETROS (itens a serem cotados, lista de emails dos fornecedores)
def enviar_email(email,itens_cotados):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = email
    mail.Subject = 'Teste Cotação de Itens'
    mail.HTMLBody = f'''

        <h3>Segue lista referente aos itens a serem cotados</h3>
        <br>
        <h4>Itens cotados</h4>
        <br>
        <h2>Cotações:</h2>
        {itens_cotados.to_html()}

    '''
    mail.Send()

tratar_cotacao(cotacao,fornecedor)
