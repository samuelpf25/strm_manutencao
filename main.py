# última edição 04/09/2024
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import streamlit as st
import pandas as pd
from datetime import date, datetime
from io import BytesIO
import pytz
import matplotlib.pyplot as plt

fuso_horario_sp = pytz.timezone('America/Sao_Paulo')
# ocultar menu
hide_streamlit_style = """
<meta http-equiv="Content-Language" content="pt-br">
<style>
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
</style>
"""
st.markdown(hide_streamlit_style, unsafe_allow_html=True)
# fim ocultar menu

# 1) DECLARAÇÃO DE VARIÁVEIS GLOBAIS ####################################################################################
scope = ['https://spreadsheets.google.com/feeds']
k = st.secrets["senha"]
json = {
    "type": st.secrets["type"],
    "project_id": st.secrets["project_id"],
    "private_key_id": st.secrets["project_id"],
    "private_key": st.secrets["private_key"],
    "client_email": st.secrets["client_email"],
    "client_id": st.secrets["client_id"],
    "auth_uri": st.secrets["auth_uri"],
    "token_uri": st.secrets["token_uri"],
    "auth_provider_x509_cert_url": st.secrets["auth_provider_x509_cert_url"],
    "client_x509_cert_url": st.secrets["client_x509_cert_url"],
    "universe_domain": st.secrets["universe_domain"],
}

creds = ServiceAccountCredentials.from_json_keyfile_dict(json, scope)

cliente = gspread.authorize(creds)

sheet = cliente.open_by_url(
    'https://docs.google.com/spreadsheets/d/1zqIL_TnTewKwPkTTWtLlrsGBQnl9r6ZN6GSrjromXq4/edit?gid=0#gid=0').get_worksheet(
    0)  # https://docs.google.com/spreadsheets/d/1PhJXFOKdEAjcILQCDyJ-couaDM6EWBUXM1GVh-3gZWM/edit#gid=96577098

dados = sheet.get_all_records()  # Get a list of all records

df = pd.DataFrame(dados)
df = df.astype(str)


def conexao(aba="Outros",
            chave='1zqIL_TnTewKwPkTTWtLlrsGBQnl9r6ZN6GSrjromXq4',
            linha_inicial=2):  # Linha inicial para carregar os dados
    """
    Carrega os dados da planilha do Google Sheets a partir de uma determinada linha
    """
    sheet = cliente.open_by_key(chave).worksheet(aba)  # Abre a aba da planilha
    dados = sheet.get_all_values()[linha_inicial - 1:]  # Pega os dados a partir da linha especificada
    df = pd.DataFrame(dados[1:], columns=dados[0])  # Cria o DataFrame a partir dos dados
    return sheet, dados, df


areas = ['Ar-Condicionado ou Refrigeração','Elétrica ou Iluminação', 'Rede de Água ou Esgoto', 'Outros']
tipos = {
    areas[0]: ['Não está ligando','Está pingando','Retirada ou instalação de aparelho','Fazendo barulho alto','Controle não funciona','Outro'],
    areas[1]: ['Tomada (instalação/desinstalação/manutenção)','Iluminação (instalação/desinstalação/manutenção)','Manutenção em quadro elétrico','Falta de energia','Outros'],
    areas[2]: ['Falta de água','Vazamento de água ou esgoto','Torneira (instalação/desinstalação/manutenção)','Vaso Sanitário ou Mictório (instalação/desinstalação/manutenção)','Pia (instalação/desinstalação/manutenção)','Outro'],
    areas[3]: ['Porta ou Fechadura (instalação/desinstalação/manutenção)','Janela ou Vidro (instalação/desinstalação/manutenção)','Infiltração','Quadros / Placa de Formatura (instalação/desinstalação/manutenção)','Outro','Equipamentos']
}

#
datando = []
data_hora = []
nome_solicitante = []
area_manutencao = []
tipo_solicitacao = []
descricao_sucinta = []
predio = []
sala = []
data_solicitacao = []
ordem_servico = []
obsinterna = []
telefone = []
urg_uft = []
status_uft = []
data_status = []
alerta_coluna = []
pontos = []
obs_usuario = []
obs_interna = []
sala = []
email=[]
id_uft=[]


horas = ['08:00', '08:30', '09:00', '09:30', '10:00', '10:30', '11:00', '11:30', '12:00', '12:30', '13:00', '13:30',
         '14:00', '14:30', '15:00', '15:30', '16:00', '16:30', '17:00', '17:30', '18:00', '18:30', '19:00', '19:30']

# 2) padroes #####################################################################################################
padrao = '<p style="font-family:Courier; color:Blue; font-size: 16px;">'
infor = '<p style="font-family:Courier; color:Green; font-size: 16px;">'
alerta = '<p style="font-family:Courier; color:Red; font-size: 17px;">'
titulo = '<p style="font-family:Courier; color:Blue; font-size: 20px;">'
cabecalho = '<div id="logo" class="span8 small"><h1>CONTROLE DE ORDENS DE SERVIÇO - UFT</h1></div>'


# @st.cache
# def carrega_todos(status,indice,os,obsemail,obsinterna):
#     status = st.selectbox('Selecione o status:', status, index=indice)
#     os = st.text_input('Número da OS:', value=os[n])
#     obs_email = st.text_area('Observação para o Usuário:', value=obsemail[n])
#     obs_interna = st.text_area('Observação Interna:', value=obsinterna[n])
#     s = st.text_input("Senha:", value="", type="password")  # , type="password"
#     return status,os,obs_email,obs_interna,s

# 3) FUNÇÕES GLOBAIS #############################################################################################


def registra_historico(codigo,status,obsusuario,obsinterna):
    chave = '1zqIL_TnTewKwPkTTWtLlrsGBQnl9r6ZN6GSrjromXq4'
    aba = 'historico'
    sheet2, dados2, df2 = conexao(aba=aba, chave=chave)
    celula = sheet2.find("**vazio**")

    data_hoje = datetime.now(fuso_horario_sp)
    data_reg = data_hoje.strftime('%d/%m/%Y')
    hora_reg = data_hoje.strftime('%H:%M:%S')
    sheet2.update_acell('A' + str(celula.row), data_reg)
    sheet2.update_acell('B' + str(celula.row), hora_reg)
    sheet2.update_acell('C' + str(celula.row), codigo)
    sheet2.update_acell('D' + str(celula.row), status)
    sheet2.update_acell('E' + str(celula.row), obsusuario)
    sheet2.update_acell('F' + str(celula.row), obsinterna)

def next_available_row(worksheet):
    str_list = list(filter(None, worksheet.col_values(1)))
    return str(len(str_list) + 1)


def to_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='OS')
    workbook = writer.book
    worksheet = writer.sheets['OS']
    format1 = workbook.add_format({'num_format': '0.00'})
    worksheet.set_column('A:A', None, format1)
    writer.close()
    processed_data = output.getvalue()
    return processed_data


# 4) PÁGINAS ####################################################################################


st.sidebar.title('Gestão Manutenção Predial')
a = k
# pg=st.sidebar.selectbox('Selecione a Página',['Solicitações em Aberto','Solicitações a Finalizar','Consulta'])
pg = st.sidebar.radio('', ['Edição individual', 'Edição em Lote', 'Alertas', 'Consulta', 'Prioridades do dia'])
status = ['', 'Todas Ativas', 'OS Aberta', 'Pendente de Material', 'Pendente Solicitante', 'Pendente Outros',
          'Atendida', 'Material Solicitado', 'Material Disponível', 'Indeferido','Cancelada']
status_todos = ['', 'OS Aberta', 'Pendente de Material', 'Pendente Solicitante', 'Pendente Outros', 'Atendida',
                'Material Solicitado', 'Material Disponível', 'Indeferido','Cancelada']
if (pg == 'Edição individual'):
    # PÁGINA EDIÇÃO INDIVIDUAL ******************************************************************************************
    st.markdown(cabecalho, unsafe_allow_html=True)
    st.subheader(pg)
    # cabeçalho

    col1, col2 = st.columns(2)
    filtrando = col1.multiselect('Selecione o Status para Filtrar', status)
    # print(filtrando)
    filtra_os = col2.text_input('Filtrar OS:', value='')
    # data_hora	nome_solicitante	area_manutencao	tipo_solicitacao	descricao_sucinta	predio	sala	data_solicitacao		telefone	urg_uft	status_uft	data_status	alerta_coluna	pontos

    for dic in df.index:
        if (filtrando == ['Todas Ativas']):
            filtrando = status_todos
        if filtra_os != '':
            if df['status_uft'][dic] in filtrando and df['area_manutencao'][dic] != '' and str(
                    df['ordem_servico'][dic]) == str(filtra_os):
                # print(df['Código da UFT'][dic])
                data_hora.append(df['data_hora'][dic])
                nome_solicitante.append(df['nome_solicitante'][dic])
                area_manutencao.append(df['area_manutencao'][dic])
                tipo_solicitacao.append(df['tipo_solicitacao'][dic])
                descricao_sucinta.append(df['descricao_sucinta'][dic])
                predio.append(df['predio'][dic])
                sala.append(df['sala'][dic])
                data_solicitacao.append(df['data_solicitacao'][dic])
                telefone.append(df['telefone'][dic])
                urg_uft.append(df['urg_uft'][dic])
                status_uft.append(df['status_uft'][dic])
                data_status.append(df['data_status'][dic])
                alerta_coluna.append(df['alerta_coluna'][dic])
                pontos.append(df['pontos'][dic])
                ordem_servico.append(df['ordem_servico'][dic])
                obs_usuario.append(df['obs_usuario'][dic])
                obs_interna.append(df['obs_interna'][dic])
                id_uft.append(df['id_uft'][dic])
                email.append(df['email'][dic])

        else:

            if df['status_uft'][dic] in filtrando and df['area_manutencao'][dic] != '':
                data_hora.append(df['data_hora'][dic])
                nome_solicitante.append(df['nome_solicitante'][dic])
                area_manutencao.append(df['area_manutencao'][dic])
                tipo_solicitacao.append(df['tipo_solicitacao'][dic])
                descricao_sucinta.append(df['descricao_sucinta'][dic])
                predio.append(df['predio'][dic])
                sala.append(df['sala'][dic])
                data_solicitacao.append(df['data_solicitacao'][dic])
                telefone.append(df['telefone'][dic])
                urg_uft.append(df['urg_uft'][dic])
                status_uft.append(df['status_uft'][dic])
                data_status.append(df['data_status'][dic])
                alerta_coluna.append(df['alerta_coluna'][dic])
                pontos.append(df['pontos'][dic])
                ordem_servico.append(df['ordem_servico'][dic])
                obs_usuario.append(df['obs_usuario'][dic])
                obs_interna.append(df['obs_interna'][dic])
                id_uft.append(df['id_uft'][dic])
                email.append(df['email'][dic])

    if len(data_hora) > 1 and (filtra_os != ''):
        st.markdown(
            alerta + f'<Strong><i>Foram encontradas {len(data_hora)} Ordens de Serviço com este mesmo número, selecione abaixo a solicitação correspondente:</i></Strong></p>',
            unsafe_allow_html=True)
    selecionado = st.selectbox('Nº da UFT:', id_uft)
    st.markdown(
        alerta + f'<Strong><i>Foram encontradas {len(id_uft)} solicitações com este filtro.</i></Strong></p>',
        unsafe_allow_html=True)
    if (len(ordem_servico) > 0):
        n = id_uft.index(selecionado)

        # apresentar dados da solicitação
        st.markdown(titulo + '<b>Dados da Solicitação</b></p>', unsafe_allow_html=True)
        # st.text('<p style="font-family:Courier; color:Blue; font-size: 20px;">Nome: '+ nome[n]+'</p>',unsafe_allow_html=True)

        st.markdown(padrao + '<b>Nome</b>: ' + str(nome_solicitante[n]) + '</p>', unsafe_allow_html=True)
        st.markdown(padrao + '<b>Prédio</b>: ' + str(predio[n]) + '</p>', unsafe_allow_html=True)
        st.markdown(padrao + '<b>Sala</b>: ' + sala[n] + '</p>', unsafe_allow_html=True)
        st.markdown(padrao + '<b>Data da Solicitação</b>: ' + str(data_hora[n]) + '</p>', unsafe_allow_html=True)
        st.markdown(padrao + '<b>Descrição</b>: ' + str(descricao_sucinta[n]) + '</p>', unsafe_allow_html=True)
        st.markdown(padrao + '<b>Telefone</b>: ' + telefone[n] + '</p>', unsafe_allow_html=True)
        st.markdown(padrao + '<b>E-mail:</b>' + email[n] + '</p>', unsafe_allow_html=True)

        with st.expander("Histórico da OS"):
            bot = st.button("Carregar Histórico")
            if bot == True:
                with st.spinner('Carregando dados...'):
                    chave = '1zqIL_TnTewKwPkTTWtLlrsGBQnl9r6ZN6GSrjromXq4'
                    aba = 'historico'
                    sheet2, dados2, df2 = conexao(aba=aba, chave=chave, linha_inicial=1)
                    dados_hist = df2[['data_reg','hora_reg','codigo','status','obs_usuario','obs_interna']]
                    #print(dados_hist)
                    # sheet.update_acell('AC1', selecionado)  # Numero UFT
                    df_hist = df2.astype(str)
                    filtro = df_hist['codigo'].isin([str(id_uft[n])])
                    # df_hist = pd.DataFrame(dados_hist[filtro])
                    dad_h = dados_hist[filtro]

                    # df_hist = df_hist.astype(str)
                    st.dataframe(dad_h.astype(str))
        celula = sheet.find(str(id_uft[n]))
        # procurando status equivalente na lista
        indice = 0
        cont = 0
        numero = ""
        for j in status_todos:
            if j == status_uft[n]:
                indice = cont
                numero = j
            cont = cont + 1

        cont = 0

        i_urg = 0
        for urg in ['Baixa', 'Média', 'Alta']:
            if urg == urg_uft[n]:
                i_urg = cont
                break
            cont = cont + 1

        cont = 0
        i_area = 0
        for urg in areas:
            if urg == area_manutencao[n]:
                i_area = cont
                break
            cont = cont + 1

        area_reg = st.selectbox('Selecione a área:', areas, index=i_area)

        cont = 0
        i_area = 0
        for urg in areas:
            if urg == area_reg:
                i_area = cont
                break
            cont = cont + 1

        print('############################ TIPO LIST: ##################################')
        print(tipos[areas[i_area]])

        cont = 0
        i_tipo = 0
        for urg in tipos[areas[i_area]]:
            if urg == tipo_solicitacao[n]:
                i_tipo = cont
                break
            cont = cont + 1

        print(i_tipo)

        with st.form(key='my_form'):
            tipo_reg = st.selectbox('Selecione o tipo de solicitação:', tipos[areas[i_area]], index=i_tipo)
            n_os = st.text_input('Nº de OS:', value=ordem_servico[n])
            status_reg = st.selectbox('Selecione o status:', status_todos, index=indice)
            obs_usr = st.text_area('Observação para o Usuário:', value=obs_usuario[n])
            obs_int = st.text_area('Observação Interna:', value=obs_interna[n])
            urg_m = st.selectbox('Urgência UFT:', ['Baixa', 'Média', 'Alta'], index=i_urg)

            s = st.text_input("Senha:", value="", type="password")  # , type="password"

            botao = st.form_submit_button('Registrar')

        if (botao == True and s == a):
            if (sheet.cell(celula.row, 21).value == id_uft[n] and sheet.cell(celula.row, 1).value != ''):
                with st.spinner('Registrando dados...Aguarde!'):
                    st.markdown(infor + '<b>Registro efetuado!</b></p>', unsafe_allow_html=True)

                    sheet.update_acell('K' + str(celula.row), n_os)
                    sheet.update_acell('L' + str(celula.row), urg_m)
                    sheet.update_acell('M' + str(celula.row), status_reg)
                    sheet.update_acell('D' + str(celula.row), area_reg)
                    sheet.update_acell('E' + str(celula.row), tipo_reg)
                    sheet.update_acell('N' + str(celula.row), obs_usr)  # obs_email
                    sheet.update_acell('O' + str(celula.row), obs_int)  # obs_interna

                    data_hoje = datetime.now(fuso_horario_sp)
                    data_reg = data_hoje.strftime('%d/%m/%Y')
                    sheet.update_acell('P' + str(celula.row), data_reg)

                    sheet.update_acell('X' + str(celula.row), 'sim' if status_reg == 'Cancelada' else '')
                st.success('Registro efetuado!')
                with st.spinner('Registrando histórico..Aguarde!'):
                    registra_historico(selecionado, status_reg, obs_usr, obs_int)
            else:
                st.error('Código de OS inválido!')
        elif (botao == True and s != a):
            st.markdown(alerta + '<b>Senha incorreta!</b></p>', unsafe_allow_html=True)
    else:
        st.markdown(infor + '<b>Não há itens na condição ' + pg + '</b></p>', unsafe_allow_html=True)

elif pg == 'Edição em Lote':

    # PÁGINA EDIÇÃO EM LOTE  ******************************************************************************************
    st.markdown(cabecalho, unsafe_allow_html=True)

    st.subheader(pg)
    col1, col2 = st.columns(2)
    filtrando = col1.multiselect('Selecione o Status para Filtrar', status)
    if (filtrando == ['Todas Ativas']):
        filtrando = status_todos
    os_gerais = []
    for dic in df.index:
        if (filtrando != ''):
            if df['ordem_servico'][dic] != '' and df['status_uft'][dic] in filtrando:
                os_gerais.append(df['ordem_servico'][dic])
        else:
            if df['ordem_servico'] != '':
                os_gerais.append(df['ordem_servico'][dic])

    filtra_os = col2.multiselect('Filtrar OS:', os_gerais)
    for dic in df.index:
        print(str(df['ordem_servico'][dic]))
        # print(str(filtra_os))
        if filtra_os != '':
            if df['status_uft'][dic] in filtrando and df['area_manutencao'][dic] != '' and (
                    str(df['ordem_servico'][dic]) in filtra_os) and (str(df['ordem_servico'][dic]) != '') and (
                    str(df['ordem_servico'][dic]) != 0):
                data_hora.append(df['data_hora'][dic])
                nome_solicitante.append(df['nome_solicitante'][dic])
                area_manutencao.append(df['area_manutencao'][dic])
                tipo_solicitacao.append(df['tipo_solicitacao'][dic])
                descricao_sucinta.append(df['descricao_sucinta'][dic])
                predio.append(df['predio'][dic])
                sala.append(df['sala'][dic])
                data_solicitacao.append(df['data_solicitacao'][dic])
                telefone.append(df['telefone'][dic])
                urg_uft.append(df['urg_uft'][dic])
                status_uft.append(df['status_uft'][dic])
                data_status.append(df['data_status'][dic])
                alerta_coluna.append(df['alerta_coluna'][dic])
                pontos.append(df['pontos'][dic])
                ordem_servico.append(df['ordem_servico'][dic])
                obs_usuario.append(df['obs_usuario'][dic])
                obs_interna.append(df['obs_interna'][dic])
                id_uft.append(df['id_uft'][dic])
                email.append(df['email'][dic])
        else:
            if df['status_uft'][dic] in filtrando and df['area_manutencao'][dic] != '':
                # print(df['Código da UFT'][dic])
                data_hora.append(df['data_hora'][dic])
                nome_solicitante.append(df['nome_solicitante'][dic])
                area_manutencao.append(df['area_manutencao'][dic])
                tipo_solicitacao.append(df['tipo_solicitacao'][dic])
                descricao_sucinta.append(df['descricao_sucinta'][dic])
                predio.append(df['predio'][dic])
                sala.append(df['sala'][dic])
                data_solicitacao.append(df['data_solicitacao'][dic])
                telefone.append(df['telefone'][dic])
                urg_uft.append(df['urg_uft'][dic])
                status_uft.append(df['status_uft'][dic])
                data_status.append(df['data_status'][dic])
                alerta_coluna.append(df['alerta_coluna'][dic])
                pontos.append(df['pontos'][dic])
                ordem_servico.append(df['ordem_servico'][dic])
                obs_usuario.append(df['obs_usuario'][dic])
                obs_interna.append(df['obs_interna'][dic])
                id_uft.append(df['id_uft'][dic])
                email.append(df['email'][dic])
    # if len(n_solicitacao)>1:
    #    st.markdown(alerta + f'<Strong><i>Foram encontradas {len(n_solicitacao)} Ordens de Serviço com este mesmo número, exclua da lista abaixo a solicitação que não for correspondente a que queira editar:</i></Strong></p>',unsafe_allow_html=True)

    selecionado = st.multiselect('Nº da solicitação:', ordem_servico, ordem_servico)
    filtro = selecionado
    dados1 = df[['area_manutencao', 'predio', 'data_solicitacao', 'ordem_servico', 'status_uft']]
    filtrar = dados1['ordem_servico'].isin(filtro)
    # print(dados1[filtrar]['Ordem de Serviço'].value_counts())
    # print(dados1[filtrar]['Ordem de Serviço'].value_counts().values)
    lista_repetidos = list(dados1[filtrar]['ordem_servico'].value_counts().values)
    repeticao = 0
    for repetido in lista_repetidos:
        #    valor = dados1['Ordem de Serviço'].value_counts()
        if int(repetido) > 1:
            st.markdown(
                alerta + f'<Strong><i>Foram encontradas Ordens de Serviço com números repetidos, exclua da lista a solicitação que não for correspondente a que queira editar</i></Strong></p>',
                unsafe_allow_html=True)
            repeticao = 1
            break
    st.dataframe(dados1[filtrar].head())
    # selecionado=n_solicitacao
    # print(nome[n_solicitacao.index(selecionado)])
    if (1 > 0):  # len(n_solicitacao)
        # procurando status equivalente na lista
        with st.form(key='my_form'):

            status_reg = st.selectbox('Selecione o status:', status_todos)
            obs_usr = st.text_area('Observação para o Usuário:', value='')
            obs_int = st.text_area('Observação Interna:', value='')
            urg_m = st.selectbox('Urgência UFT:', ['Baixa', 'Média', 'Alta'])

            s = st.text_input("Senha:", value="", type="password")  # , type="password"
            botao = st.form_submit_button('Registrar')
            efetuado = 0
        if (botao == True and s == a):
            with st.spinner('Registrando dados...Aguarde!'):
                for selecionado_i in selecionado:
                    celula = sheet.find(str(selecionado_i))
                    # sheet.update_acell('P' + str(celula.row), status)  # Status
                    # print(sheet.cell(celula.row, 20).value)
                    # print(repeticao)
                    # print(selecionado_i)
                    if (sheet.cell(celula.row, 11).value == selecionado_i and sheet.cell(celula.row, 1).value != '' and repeticao == 0):
                        efetuado = 1
                        if urg_m != '':
                            sheet.update_acell('L' + str(celula.row), urg_m)

                        if (status_reg != ''):
                            sheet.update_acell('M' + str(celula.row), status_reg)
                            data_hoje = datetime.now(fuso_horario_sp)
                            data_reg = data_hoje.strftime('%d/%m/%Y')
                            sheet.update_acell('P' + str(celula.row), data_reg)
                            with st.spinner('Registrando histórico..Aguarde!'):
                                registra_historico(id_uft[ordem_servico.index(selecionado_i)], status_reg, obs_usr, obs_int)

                        # sheet.update_acell('R' + str(celula.row), '')  # apagar Sim para enviar e-mail
                        if (obs_usr != ''):
                            # sheet.update_acell('S' + str(celula.row), obsemail)  # obs_email
                            sheet.update_acell('N' + str(celula.row), obs_usr)
                        if (obs_int != ''):
                            # sheet.update_acell('AA' + str(celula.row), obsinterna)  # obs_interna
                            sheet.update_acell('O' + str(celula.row), obs_int)  # obs_interna

                        sheet.update_acell('X' + str(celula.row), 'sim' if status_reg == 'Cancelada' else '')

                            # st.markdown(infor+'<b>Registro efetuado!</b></p>',unsafe_allow_html=True)
            if (efetuado == 1):
                st.success('Registro efetuado!')
            elif (efetuado == 0 and repeticao == 1):
                st.error('Remova as OS com números repetidos!')
            elif (efetuado == 1 and repeticao == 1):
                st.error('Dados parcialmente cadastrados! As OS com números repetidos não foram registradas!')
        elif (botao == True and s != a):
            st.markdown(alerta + '<b>Senha incorreta!</b></p>', unsafe_allow_html=True)
    else:
        st.markdown(infor + '<b>Não há itens na condição ' + pg + '</b></p>', unsafe_allow_html=True)


elif pg == 'Alertas':

    # Alertas  ******************************************************************************************
    st.markdown(cabecalho, unsafe_allow_html=True)

    st.subheader(pg)

    chave = '1zqIL_TnTewKwPkTTWtLlrsGBQnl9r6ZN6GSrjromXq4'

    aba = 'Alertas'
    sheet2, dados2, df2 = conexao(aba=aba, chave=chave, linha_inicial=1)
    dados_hist = df2[['data_hora', 'nome_solicitante', 'area_manutencao', 'predio', 'data_solicitacao','ordem_servico', 'urg_uft', 'status_uft', 'data_status', 'obs_usuario','obs_interna']]
    dad_h = dados_hist

    df_filtered = df2[df2['area_manutencao'].notna() & (df2['area_manutencao'] != '')]
    st.markdown(
        alerta + f'<Strong><i>Número de OS em Alertas: {len(df_filtered)}.</i></Strong></p>',
        unsafe_allow_html=True)
    st.dataframe(dad_h.astype(str))

    # Remover valores vazios
    df_filtered = df2[df2['area_manutencao'].str.strip() != '']
    df_filtered = df_filtered[df_filtered['status_uft'].str.strip() != '']

    # Agregar dados
    chart_data1 = df_filtered.groupby(['area_manutencao', 'status_uft']).size().unstack(fill_value=0)

    # Remover valores vazios
    df_filtered = df2[df2['area_manutencao'].str.strip() != '']
    df_filtered = df_filtered[df_filtered['tipo_solicitacao'].str.strip() != '']

    # Agregar dados
    chart_data2 = df_filtered.groupby(['area_manutencao', 'tipo_solicitacao']).size().unstack(fill_value=0)

    # Remover valores vazios
    df_filtered = df2[df2['area_manutencao'].str.strip() != '']
    df_filtered = df_filtered[df_filtered['predio'].str.strip() != '']

    # Agregar dados
    chart_data3 = df_filtered.groupby(['area_manutencao', 'predio']).size().unstack(fill_value=0)

    # Exibir gráfico de barras
    st.write("Área de Manutenção x Status")
    st.bar_chart(chart_data1)

    st.write("Área de Manutenção x Tipo de Solicitação")
    st.bar_chart(chart_data2)

    st.write("Área de Manutenção x Prédio")
    st.bar_chart(chart_data3)

############################## GRAF #######################################
    # Agrupar por area_manutencao e contar o total de OS por área
    # Filtrar os dados para ignorar valores vazios em 'area_manutencao'
    df_filtered = df2[df2['area_manutencao'].notna() & (df2['area_manutencao'] != '')]

    # Agrupar por area_manutencao e contar o total de OS por área
    chart_data = df_filtered.groupby('area_manutencao').size()

    # Rótulos e tamanhos para o gráfico de pizza
    labels = chart_data.index
    sizes = chart_data.values

    # Explode apenas a fatia maior (opcional, aqui explodimos a maior área)
    explode = [0.1 if size == max(sizes) else 0 for size in sizes]

    # Criar o gráfico de pizza
    fig1, ax1 = plt.subplots()
    ax1.pie(sizes, explode=explode, labels=labels, autopct='%1.1f%%',
            shadow=True, startangle=90)
    ax1.axis('equal')  # Para garantir que o gráfico fique em formato de círculo.
    ax1.set_title('OS por Área de Manutenção')
    # Exibir o gráfico no Streamlit
    st.pyplot(fig1)

    ############################## GRAF #######################################
    df_filtered = df2[df2['predio'].notna() & (df2['predio'] != '')]

    # Agrupar por area_manutencao e contar o total de OS por área
    chart_data = df_filtered.groupby('predio').size()

    # Rótulos e tamanhos para o gráfico de pizza
    labels = chart_data.index
    sizes = chart_data.values

    # Explode apenas a fatia maior (opcional, aqui explodimos a maior área)
    explode = [0.1 if size == max(sizes) else 0 for size in sizes]

    # Criar o gráfico de pizza
    fig1, ax1 = plt.subplots()
    ax1.pie(sizes, explode=explode, labels=labels, autopct='%1.1f%%',
            shadow=True, startangle=90)
    ax1.axis('equal')  # Para garantir que o gráfico fique em formato de círculo.
    ax1.set_title('OS por Prédio')
    # Exibir o gráfico no Streamlit
    st.pyplot(fig1)

    ############################## GRAF #######################################
    df_filtered = df2[df2['status_uft'].notna() & (df2['status_uft'] != '')]

    # Agrupar por area_manutencao e contar o total de OS por área
    chart_data = df_filtered.groupby('status_uft').size()

    # Rótulos e tamanhos para o gráfico de pizza
    labels = chart_data.index
    sizes = chart_data.values

    # Explode apenas a fatia maior (opcional, aqui explodimos a maior área)
    explode = [0.1 if size == max(sizes) else 0 for size in sizes]

    # Criar o gráfico de pizza
    fig1, ax1 = plt.subplots()
    ax1.pie(sizes, explode=explode, labels=labels, autopct='%1.1f%%',
            shadow=True, startangle=90)
    ax1.axis('equal')  # Para garantir que o gráfico fique em formato de círculo.
    ax1.set_title('OS por Status')
    # Exibir o gráfico no Streamlit
    st.pyplot(fig1)
elif pg == 'Consulta':

    # PÁGINA DE CONSULTA ************************************************************************************************
    # st.markdown(cabecalho, unsafe_allow_html=True)
    # st.subheader(pg)
    # titulos = ['data_hora','nome_solicitante','area_manutencao','tipo_solicitacao','descricao_sucinta','sala','data_solicitacao','telefone','urg_uft','status_uft','data_status','alerta_coluna','pontos','ordem_servico','obs_usuario','obs_interna']
    #
    # sample_data = df[titulos]
    #
    # AwesomeTable(pd.json_normalize(sample_data), columns=titulos, show_order=True, show_search=True, show_search_order_in_sidebar=True)
    #

    # dados = sheet.get_all_records()  # Get a list of all records
    # df = pd.DataFrame(dados)
    data_hora.append('')
    nome_solicitante.append('')
    area_manutencao.append('')
    tipo_solicitacao.append('')
    descricao_sucinta.append('')
    predio.append('')
    sala.append('')
    data_solicitacao.append('')
    obsinterna.append('')
    telefone.append('')
    urg_uft.append('')
    status_uft.append('')
    data_status.append('')
    alerta_coluna.append('')
    pontos.append('')
    ordem_servico.append('')
    obs_usuario.append('')
    obs_interna.append('')
    email.append('')

    for dic in df.index:
        if df['predio'][dic] != '':
            # print(df['Código da UFT'][dic])
            data_hora.append(df['data_hora'][dic])
            nome_solicitante.append(df['nome_solicitante'][dic])
            area_manutencao.append(df['area_manutencao'][dic])
            tipo_solicitacao.append(df['tipo_solicitacao'][dic])
            descricao_sucinta.append(df['descricao_sucinta'][dic])
            predio.append(df['predio'][dic])
            sala.append(df['sala'][dic])
            data_solicitacao.append(df['data_solicitacao'][dic])
            telefone.append(df['telefone'][dic])
            urg_uft.append(df['urg_uft'][dic])
            status_uft.append(df['status_uft'][dic])
            data_status.append(df['data_status'][dic])
            alerta_coluna.append(df['alerta_coluna'][dic])
            pontos.append(df['pontos'][dic])
            ordem_servico.append(df['ordem_servico'][dic])
            obs_usuario.append(df['obs_usuario'][dic])
            obs_interna.append(df['obs_interna'][dic])
            id_uft.append(df['id_uft'][dic])
            email.append(df['email'][dic])

    st.markdown(cabecalho, unsafe_allow_html=True)
    st.subheader(pg)


    titulos = ['data_hora', 'nome_solicitante', 'area_manutencao', 'tipo_solicitacao', 'descricao_sucinta',
               'sala', 'data_solicitacao', 'telefone', 'urg_uft', 'status_uft', 'data_status',
               'alerta_coluna', 'pontos', 'ordem_servico', 'obs_usuario', 'obs_interna', 'predio','sala','email']
    dados = df[titulos].astype(str).fillna('')
    dad = dados

    st.markdown(
        alerta + f'<Strong><i>Número de OS com o filtro correspondente: {len(dad[dad['area_manutencao'].str.strip() != ''])}.</i></Strong></p>',
        unsafe_allow_html=True)
    # Formulário de filtro
    with st.form(key='form1'):
        coluna_busca = st.selectbox('Coluna para busca por argumento', titulos)
        texto = st.text_input('Busca por argumento na coluna selecionada: ')
        btn1 = st.form_submit_button('Filtrar')

        # Filtragem
        if btn1:
            if texto and coluna_busca:
                dad = dados[dados[coluna_busca].str.contains(texto, na=False)]
            else:
                dad = dados
            st.dataframe(dad)
            df_xlsx = to_excel(dad)
            st.download_button(label='📥 Baixar Resultado do Filtro em Excel', data=df_xlsx,
                               file_name='filtro_planilha.xlsx')

        # dados_graf=pd.DataFrame(dados[filtrar],columns=[coluna1,coluna2])
        # fig = px.bar(dados_graf, x=coluna1, y=coluna2, barmode='group', height=400)
        # st.plotly_chart(fig)
        # plost.line_chart(dados_graf, coluna1, coluna2)

        # else:
        #    st.dataframe(df[titulos])


    try:
        df2 = dad
        # Remover valores vazios
        df_filtered = df2[df2['area_manutencao'].str.strip() != '']
        df_filtered = df_filtered[df_filtered['status_uft'].str.strip() != '']

        # Agregar dados
        chart_data1 = df_filtered.groupby(['area_manutencao', 'status_uft']).size().unstack(fill_value=0)

        # Remover valores vazios
        df_filtered = df2[df2['area_manutencao'].str.strip() != '']
        df_filtered = df_filtered[df_filtered['tipo_solicitacao'].str.strip() != '']

        # Agregar dados
        chart_data2 = df_filtered.groupby(['area_manutencao', 'tipo_solicitacao']).size().unstack(fill_value=0)

        # Remover valores vazios
        df_filtered = df2[df2['area_manutencao'].str.strip() != '']
        df_filtered = df_filtered[df_filtered['predio'].str.strip() != '']

        # Agregar dados
        chart_data3 = df_filtered.groupby(['area_manutencao', 'predio']).size().unstack(fill_value=0)

        # Exibir gráfico de barras
        st.write("Área de Manutenção x Status")
        st.bar_chart(chart_data1)

        st.write("Área de Manutenção x Tipo de Solicitação")
        st.bar_chart(chart_data2)

        st.write("Área de Manutenção x Prédio")
        st.bar_chart(chart_data3)

        ############################## GRAF #######################################
        # Agrupar por area_manutencao e contar o total de OS por área
        # Filtrar os dados para ignorar valores vazios em 'area_manutencao'
        df_filtered = df2[df2['area_manutencao'].notna() & (df2['area_manutencao'] != '')]

        # Agrupar por area_manutencao e contar o total de OS por área
        chart_data = df_filtered.groupby('area_manutencao').size()

        # Rótulos e tamanhos para o gráfico de pizza
        labels = chart_data.index
        sizes = chart_data.values

        # Explode apenas a fatia maior (opcional, aqui explodimos a maior área)
        explode = [0.1 if size == max(sizes) else 0 for size in sizes]

        # Criar o gráfico de pizza
        fig1, ax1 = plt.subplots()
        ax1.pie(sizes, explode=explode, labels=labels, autopct='%1.1f%%',
                shadow=True, startangle=90)
        ax1.axis('equal')  # Para garantir que o gráfico fique em formato de círculo.
        ax1.set_title('OS por Área de Manutenção')
        # Exibir o gráfico no Streamlit
        st.pyplot(fig1)

        ############################## GRAF #######################################
        df_filtered = df2[df2['predio'].notna() & (df2['predio'] != '')]

        # Agrupar por area_manutencao e contar o total de OS por área
        chart_data = df_filtered.groupby('predio').size()

        # Rótulos e tamanhos para o gráfico de pizza
        labels = chart_data.index
        sizes = chart_data.values

        # Explode apenas a fatia maior (opcional, aqui explodimos a maior área)
        explode = [0.1 if size == max(sizes) else 0 for size in sizes]

        # Criar o gráfico de pizza
        fig1, ax1 = plt.subplots()
        ax1.pie(sizes, explode=explode, labels=labels, autopct='%1.1f%%',
                shadow=True, startangle=90)
        ax1.axis('equal')  # Para garantir que o gráfico fique em formato de círculo.
        ax1.set_title('OS por Prédio')
        # Exibir o gráfico no Streamlit
        st.pyplot(fig1)

        ############################## GRAF #######################################
        df_filtered = df2[df2['status_uft'].notna() & (df2['status_uft'] != '')]

        # Agrupar por area_manutencao e contar o total de OS por área
        chart_data = df_filtered.groupby('status_uft').size()

        # Rótulos e tamanhos para o gráfico de pizza
        labels = chart_data.index
        sizes = chart_data.values

        # Explode apenas a fatia maior (opcional, aqui explodimos a maior área)
        explode = [0.1 if size == max(sizes) else 0 for size in sizes]

        # Criar o gráfico de pizza
        fig1, ax1 = plt.subplots()
        ax1.pie(sizes, explode=explode, labels=labels, autopct='%1.1f%%',
                shadow=True, startangle=90)
        ax1.axis('equal')  # Para garantir que o gráfico fique em formato de círculo.
        ax1.set_title('OS por Status')
        # Exibir o gráfico no Streamlit
        st.pyplot(fig1)
    except:
        pass



elif pg == 'Prioridades do dia':
    st.markdown(cabecalho, unsafe_allow_html=True)
    st.subheader(pg)

    chave = '1zqIL_TnTewKwPkTTWtLlrsGBQnl9r6ZN6GSrjromXq4'
    aba = st.selectbox('Selecione a área', areas)
    sheet2, dados2, df2 = conexao(aba=aba, chave=chave)
    st.dataframe(df2)
