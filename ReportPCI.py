import pandas as pd
import win32com.client
from time import sleep
import sys
import xlwings as xw
import datetime 
import os
import matplotlib.pyplot as plt
from DPCP import calendario

''' Este script tem a finalidade de buscar todos os PCIs que estão em fluxo de cadastro, buscar as obras em que estes PCIs estão atrelados e enviar e-mail para cada departamento que tem seu PCI pendente.
Funcionamento:
-Ler os códigos de materiais na planilha "Código e campos.xlsx"
-Colocar os códigos e campos na transação ZM277 para ver os departamentos pendentes
-Por meio da transação ZM277, acessar o SLA destes materiais para obter a data de entrada deste material em cada departamento
-Colocar os materiais na ZP059 para ver quais obras os materiais estão atrelados
-Montar uma tabela consolidada a partir das informações das duas transações
-Enviar e-mail aos departamentos responsáveis
-Calcular quanto tempo o material está em cada departamento pendente, utilizando as informações obtidas no SLA
-Enviar e-mail à gestão do DPCP informando os dempos de cada departamento 

@autor: Gustavo Nunes Ferraz
@departamento: DPCP
@Última modificação: 13/01/2024

Histórico de modificações:
-29/10/2024: Script finalizado
-13/01/2025: Acesso ao SLA dos PCIs + Envio de e-mail ao supervisor do DPCP informando os tempos por departamento + Adição do Readme e requirments.
-16/01/2025: Adição de um input perguntando se o usuário quer enviar um relatório de tempos por departamentos ao supervisor do DPCP
'''

#Confidentially informations - variables which contains directory's path

# PATH_EMAILS = 
# PATH_PCI = 
# FILE_CODIGO_CAMPOS = 
# PATH_TEMP = 
# FILE_CODIGO =
# FILE_CAMPO = 
# FILE_ZM255 = 
# FILE_ZP059 = 
# FILE_SLA = 
# FILE_MATERIAIS_ZM255 = 

def verifica_planilhas_abertas():
    excel = win32com.client.Dispatch("Excel.Application")
    for arquivo in excel.Workbooks:
        if any(f in arquivo.Name for f in [FILE_CODIGO, FILE_CAMPO, FILE_ZM255, FILE_ZP059, FILE_MATERIAIS_ZM255, FILE_CODIGO_CAMPOS, FILE_SLA]):
            print("Atenção! Algum dos arquivos abaixo está aberto:")
            print([FILE_CODIGO, FILE_CAMPO, FILE_ZM255, FILE_ZP059, FILE_MATERIAIS_ZM255, FILE_CODIGO_CAMPOS, FILE_SLA])
            print("Salve e feche os arquivos antes de executar o script")
            print("O programa será encerrado.")
            sleep(10)
            sys.exit()

def conecta_sap():
    try: #Garante que o SAP está aberto, caso contrário encerra o programa
    #Pega a aplicação COM do SAP e a primeira sessão para uso no script
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
        connection = application.Children(0)
        sessions = connection.Children
        session = sessions[0]
        return session
    except:
        print("SAP não está aberto, o programa será finalizado")
        sleep(2)
        sys.exit(0)
    #Testa se há um usuário logado no SAP, caso não haja, encerra o programa
    usuario = session.Info.User
    if usuario == '':
        print("SAP não está logado, o programa será finalizado")
        time.sleep(2)
        sys.exit(0)  

def conecta_outlook():
    try:
        outlook = win32com.client.Dispatch('outlook.application')
        return outlook
    except ConnectionError:
        print('Erro ao estabelecer conexão com outlook.')
        sleep(3)
        sys.exit(0)

def ler_campos():
    df_codigo_campos = pd.read_excel(PATH_PCI + "\\" + FILE_CODIGO_CAMPOS, header=None)
    df_codigo = df_codigo_campos[0]
    df_codigo = df_codigo.dropna()
    df_campo = df_codigo_campos[1]
    df_campo = df_campo.dropna()
    df_codigo.to_csv(PATH_TEMP + "\\" + FILE_CODIGO, index=False, header=None)
    df_campo.to_csv(PATH_TEMP + "\\" + FILE_CAMPO, index=False, header=None)

def remove_arquivos_antigos():
    if os.path.exists(PATH_TEMP + "\\" + FILE_ZM255):
        try:
            os.remove(PATH_TEMP + "\\" + FILE_ZM255)
        except:
            print(r'Não foi possível remover o arquivo zm277.xlsx na pasta C:\TEMP. Remova manualmente o arquivo e execute novamente o script.')
            sys.exit()

    if os.path.exists(PATH_TEMP + "\\" + FILE_ZP059):
        try:
            os.remove(PATH_TEMP + "\\" + FILE_ZP059)
        except:
            print(r'Não foi possível remover o arquivo zp059.xlsx na pasta C:\TEMP. Remova manualmente o arquivo e execute novamente o script.')
            sys.exit()
    
    if os.path.exists(PATH_TEMP + "\\" + FILE_SLA):
        try:
            os.remove(PATH_TEMP + "\\" + FILE_SLA)
        except:
            print(r'Não foi possível remover o arquivo sla.xlsx na pasta C:\TEMP. Remova manualmente o arquivo e execute novamente o script.')
            sys.exit()

def zm277(session):
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nzm277"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/btn%_S_MAT_%_APP_%-VALU_PUSH").press()
    session.findById("wnd[1]/tbar[0]/btn[23]").press()
    session.findById("wnd[2]/usr/ctxtDY_PATH").text = PATH_TEMP
    session.findById("wnd[2]/usr/ctxtDY_FILENAME").text = FILE_CODIGO
    session.findById("wnd[2]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]/usr/btn%_S_BCRE_%_APP_%-VALU_PUSH").press()
    session.findById("wnd[1]/tbar[0]/btn[23]").press()
    session.findById("wnd[2]/usr/ctxtDY_PATH").text = PATH_TEMP
    session.findById("wnd[2]/usr/ctxtDY_FILENAME").text = FILE_CAMPO
    session.findById("wnd[2]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/usr/shell").pressToolbarContextButton("&MB_EXPORT")
    session.findById("wnd[0]/usr/shell").selectContextMenuItem("&XXL")
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = PATH_TEMP
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = FILE_ZM255
    session.findById("wnd[1]/usr/ctxtDY_PATH").setFocus()
    session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = len(PATH_TEMP)
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

    sleep(5)

    try:
        zm255_book = xw.Book(FILE_ZM255)
        zm255_book.close()
    except:
        pass

    session.findById("wnd[0]/tbar[1]/btn[13]").press()
    session.findById("wnd[0]/usr/shell").pressToolbarContextButton("&MB_EXPORT")
    session.findById("wnd[0]/usr/shell").selectContextMenuItem("&XXL")
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = PATH_TEMP
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = FILE_SLA
    session.findById("wnd[1]/usr/ctxtDY_PATH").setFocus()
    session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = len(PATH_TEMP)
    session.findById("wnd[1]/tbar[0]/btn[0]").press()  

    sleep(5)

    try:
        zm255_book = xw.Book(FILE_SLA)
        zm255_book.close()
    except:
        pass 
    
def zp059():
    print("Lendo dados da ZM277")
    df_zm255 = pd.read_excel(PATH_TEMP + "\\" + FILE_ZM255, usecols=['Material', 'Tipo de suprimento'])
    df_zm255 = df_zm255[df_zm255['Tipo de suprimento'] == 'F']
    df_zm255 = df_zm255.drop(columns=['Tipo de suprimento'])
    df_zm255.to_csv(PATH_TEMP + "\\" + FILE_MATERIAIS_ZM255, index=False, header=None)
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nzp059"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/ctxtS_IDNRK-LOW").text = "a"
    session.findById("wnd[0]/usr/ctxtP_WERKS").text = "5001"
    session.findById("wnd[0]/usr/ctxtP_STLAN").text = "2"
    session.findById("wnd[0]/usr/btn%_S_IDNRK_%_APP_%-VALU_PUSH").press()
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = ""
    session.findById("wnd[1]/tbar[0]/btn[23]").press()
    session.findById("wnd[2]/usr/ctxtDY_PATH").text = PATH_TEMP
    session.findById("wnd[2]/usr/ctxtDY_FILENAME").text = FILE_MATERIAIS_ZM255
    session.findById("wnd[2]/usr/ctxtDY_FILENAME").caretPosition = 4
    session.findById("wnd[2]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/usr/shell/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
    session.findById("wnd[0]/usr/shell/shellcont/shell").selectContextMenuItem("&XXL")
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = PATH_TEMP
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = FILE_ZP059
    session.findById("wnd[1]/usr/ctxtDY_PATH").setFocus()
    session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = len(PATH_TEMP)
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

    sleep(5)

    try:
        zm255_book = xw.Book(FILE_ZP059)
        zm255_book.close()
    except:
        pass

    return df_zm255

def monta_tabela_email():
    df_zm255 = pd.read_excel(PATH_TEMP + "\\" + FILE_ZM255, usecols=['Material', 'Bloqueio Criação'])
    df_zp059 = pd.read_excel(PATH_TEMP + "\\" + FILE_ZP059)
    df_email = df_zp059.merge(df_zm255, left_on='Compon. (Filho)', right_on='Material', how='left')
    df_email = df_email.drop(columns=['Material (Pai Imed.)', 'Nvl. Comp. (Filho)', 'Quantidade', 'Util.LisTéc.', 'Data Demanda', 'Kit (Últ. Nível)'])
    #df_email = df_email[['Material', 'Bloqueio Criação', 'Obra', 'Elemento PEP', 'Data Demanda', 'Kit (Últ. Nível)']]
    df_email = df_email[['Obra', 'Elemento PEP', 'Material', 'Bloqueio Criação']]
    return df_email

def envia_email(df_email, outlook):
    lista_destinatarios = pd.read_excel(PATH_EMAILS)
    lista_destinatarios = lista_destinatarios.ffill()
    lista_destinatarios['Departamento'] = lista_destinatarios['Departamento'].replace('Engenharia', 'ENGE')
    lista_destinatarios['Departamento'] = lista_destinatarios['Departamento'].replace('COMEX', 'DTRI')
    lista_destinatarios = lista_destinatarios[lista_destinatarios['Departamento'] != 'Gestão']

    departamentos_pendentes = df_email['Bloqueio Criação'].unique().tolist()

    for departamento in departamentos_pendentes:
        print('Enviando e-mail para: ', departamento)
        try:
            df_departamento = df_email[df_email['Bloqueio Criação'] == departamento]
            email_to = lista_destinatarios[lista_destinatarios['Departamento'] == departamento]['E-mail']
            email_to = email_to.to_list()
            email = outlook.CreateItem(0)
            email.Subject = f"Itens PCI Pendentes - {departamento}"
            # lista_materiais_pendentes_departamento = df_departamento['Material'].unique().tolist()
            lista_materiais_pendentes_departamento = df_departamento['Material'].drop_duplicates()
            lista_materiais_pendentes_departamento = lista_materiais_pendentes_departamento.to_frame().to_html(index=False)
            df_departamento = df_departamento.style.applymap(
                lambda x: 'background-color: yellow' if df_departamento['Material'].isin([x]).any() else '',
                subset=['Material']
            ).set_table_attributes('border="1" cellspacing="0" cellpadding="5"').to_html(index=False, escape=False)
            #df_departamento = df_departamento.to_html(index=False)
            
            #Enviar o e-mail para o analista responsável
            email.To = ';'.join(email_to)
            #email.To = 'gustavo.ferraz@tkelevator.com'
            email.CC = 'gustavo.ferraz@tkelevator.com'
            email.Importance = 2

            #Cria o corpo do email
            lista_email = []
            lista_email.append('<p>Prezados(as),</p>')
            lista_email.append(f'<p>Existem PCIs pendentes de cadastro de material. Segue a lista abaixo:</p>')
            lista_email.append('<br>')
            lista_email.append(lista_materiais_pendentes_departamento)
            lista_email.append('<p>')
            lista_email.append('Segue a tabela abaixo com as relações dos materiais:')
            lista_email.append(df_departamento)
            lista_email.append('<p>Atenciosamente,</p>')
            lista_email.append('Robô do DPCP')
            lista_string = '\n'.join(lista_email)

            email.HTMLBody = lista_string

            #Envia o e-mail
            email.Send()
        except:
            print('Erro no envio do e-mail para ', departamento)
            continue

def calcula_tempos_departamentos(df_tempos_departamentos, calendario_tke):
    df_sla = pd.read_excel(PATH_TEMP + "\\" + FILE_SLA)
    for index, row in df_tempos_departamentos.iterrows():
        indice_coluna_departamento = df_sla.columns.str.contains(row['Bloqueio Criação']).tolist().index(True) #Obtém o número do índice da coluna do departamento
        try:
            tempo_departamento = calendario_tke.diferenca_dias_uteis(datetime.datetime.today(), df_sla.loc[df_sla['Material'] ==  row['Material'], df_sla.columns[indice_coluna_departamento]])
            df_tempos_departamentos.loc[df_tempos_departamentos.index == index, 'Tempo no Departamento'] = tempo_departamento
        except:
            #tempo_departamento = 'ERRO SLA'
            pass

    df_tempos_departamentos['Tempo no Departamento'] = df_tempos_departamentos['Tempo no Departamento'].apply(lambda x: int(abs(x)))
    return df_tempos_departamentos

def pergunta_ao_usuario_se_envia_email_a_gestao_dpcp():
    loop = True
    while loop == True:
        resposta = input('Deseja enviar o tempo dos departamentos à gestão DPCP? (Y/N)')
        if not (resposta.upper() != 'Y' or resposta.upper() !='N'):
            print("Responda 'Sim' teclando 'Y' e 'Não' teclando 'N'. Tente novamente.")
        elif resposta.upper() =='N':
        
            return False
        elif resposta.upper() =='Y':
            return True

def envia_email_dpcp(df_tempos_departamentos):
    lista_departamentos_pendentes = df_tempos_departamentos['Bloqueio Criação'].unique().tolist()

    #Dataframe que serve para auxiliar o envio do report
    df_report_auxiliar = pd.DataFrame({'Departamento': lista_departamentos_pendentes, 'Número de materiais pendentes': None})

    #Analisa número de materiais para cada departamento
    for departamento in lista_departamentos_pendentes:
        num_materiais = len(df_tempos_departamentos.loc[df_tempos_departamentos['Bloqueio Criação'] == departamento, 'Material'].unique())
        if num_materiais == 1:
            df_report_auxiliar.loc[df_report_auxiliar['Departamento'] == departamento, 'Número de materiais pendentes'] = num_materiais
        else:
            df_report_auxiliar.loc[df_report_auxiliar['Departamento'] == departamento, 'Número de materiais pendentes'] = num_materiais

    #Monta gráfico
    for index, row in df_report_auxiliar.iterrows():
        #if row['Número de materiais pendentes'] != 1: #Caso haja mais de um material, é construído um gráfico
        df_construcao_grafico = df_tempos_departamentos.loc[df_tempos_departamentos['Bloqueio Criação'] == row['Departamento']]
        df_construcao_grafico = df_construcao_grafico.sort_values(by='Tempo no Departamento', ascending = False)
        plt.figure(figsize=(10, 6))
        num_bars = len(df_construcao_grafico)
    
    # Ajuste da largura da barra com base no número de materiais
        bars = plt.bar(df_construcao_grafico['Material'], df_construcao_grafico['Tempo no Departamento'], color='skyblue')
        plt.xlabel('Material')
        plt.ylabel('Tempo no Departamento (Dias Úteis)')
        plt.title('Histograma de Tempo no Departamento por Material')
        plt.xticks(rotation=45, ha='right')

        # Adding labels on top of each bar
        for bar in bars:
            yval = bar.get_height()
            plt.text(bar.get_x() + bar.get_width()/2, yval, f'{yval}', ha='center', va='bottom')

        plt.tight_layout()
        plt.savefig(PATH_TEMP + "\\" + row['Departamento'] + '.png')
    
    #Envia e-mail ao DPCP
    df_email = df_tempos_departamentos.drop_duplicates(subset=['Material', 'Bloqueio Criação'])
    df_email['Tempo no Departamento'] = df_email['Tempo no Departamento']
    df_email = df_email.rename(columns = {'Bloqueio Criação': 'Departamento', 'Tempo no Departamento': 'Tempo no Departamento (Dias Úteis)'})
    df_email = df_email.sort_values(by=['Tempo no Departamento (Dias Úteis)'])

    email = outlook.CreateItem(0)
    email.Subject = f"Report Tempo de Permanência de PCIs dos Departamentos"
    # lista_materiais_pendentes_departamento = df_departamento['Material'].unique().tolist()
    email.To = 'frederico.johann@tkelevator.com'
    email.CC = 'gustavo.ferraz@tkelevator.com'
    

    #Cria o corpo do email
    lista_email = []
    lista_email.append('<p>--Mensagem Automática---</p>')
    lista_email.append('<p>Olá.</p>')
    lista_email.append(f'<p>Segue abaixo o tempo de permanência de PCIs nos departamentos:</p>')
    #lista_email.append('<br>')

    df_tempos_departamentos_group = df_tempos_departamentos.groupby('Bloqueio Criação')

    for group, valores in df_tempos_departamentos_group:
        lista_email.append(f'<h2>{valores['Bloqueio Criação'].values[0]}</h2>')
        if df_report_auxiliar.loc[df_report_auxiliar['Departamento'].isin(valores['Bloqueio Criação']), 'Número de materiais pendentes'].values[0] != 1:
            lista_email.append(f'<p>O {valores['Bloqueio Criação'].values[0]} possuí {len(valores['Material'].unique())} materiais no fluxo. Segue as informações abaixo:')
            df = df_email.loc[df_email['Departamento'].isin(valores['Bloqueio Criação'])]
            df = df.sort_values(by=['Tempo no Departamento (Dias Úteis)'], ascending=False)
            lista_email.append(df.to_html(index=False))
            #Cria CID para envio do e-mail
            cid = email.Attachments.Add(PATH_TEMP + "\\" + valores['Bloqueio Criação'].values[0] + '.png')
            cid.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", f"cid:{valores['Bloqueio Criação'].values[0]}.png")
            lista_email.append(f'<img src="cid:{valores['Bloqueio Criação'].values[0]}.png"><br>')
            #lista_email.append(df.to_html())
        else:
            df = df_email.loc[df_email['Departamento'].isin(valores['Bloqueio Criação'])]
            lista_email.append(f'<p>O {valores['Bloqueio Criação'].values[0]} possuí somente um material no fluxo. Segue as informações abaixo:')
            lista_email.append(df.loc[df['Departamento'].isin(valores['Bloqueio Criação'])].to_html(index=False))


    lista_email.append('<p>Atenciosamente,</p>')
    lista_email.append('Robô do DPCP')
    lista_string = '\n'.join(lista_email)

    email.HTMLBody = lista_string

    #Envia o e-mail
    email.Send()

if __name__ == '__main__':
    print('\nIniciando o Script... \n')
    verifica_planilhas_abertas()
    print('Iniciando conexões com SAP, calendário TKE e E-mail.')
    session = conecta_sap()
    outlook = conecta_outlook()
    calendario_tke = calendario.Calendario()
    print('Lendo o arquivo código e campos.')
    ler_campos()
    print(r'Removendo extrações antigas na C:\Temp')
    remove_arquivos_antigos()
    print('Acessando a transação ZM277.')
    zm277(session)
    print('Acessando a transação ZP059.')
    df_zm255 = zp059()
    print('Montando a tabela e-mail.')
    df_email = monta_tabela_email()
    print('Enviando e-mail para os departamentos.')
    envia_email(df_email, outlook)
    resposta_usuario = pergunta_ao_usuario_se_envia_email_a_gestao_dpcp()
    #Se o usuário responder sim (Y), a função que envia report ao supervisor do DPCP é enviado. 
    if resposta_usuario:
        print('Enviando e-mail a gestão do DPCP.')
        df_tempos_departamentos = calcula_tempos_departamentos(df_email[['Material', 'Bloqueio Criação']], calendario_tke)
        envia_email_dpcp(df_tempos_departamentos)

    print('Script finalizado com sucesso.')
    sleep(3)