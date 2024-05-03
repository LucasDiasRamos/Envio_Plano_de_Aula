import pandas as pd 
import os
import configparser
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
#leitura das planilhas
tabela_PL = pd.read_excel('data.xlsx')
tabela_nome = pd.read_excel('nome_profs.xlsx')

config = configparser.ConfigParser()
config.read('config.ini')

senha = config['DEFAULT']['MINHA_SENHA']


relatorio_vazios = []
relatorio_preenchidos = []
datas_nao_preenchidas = []



for nome in tabela_nome['NOME'].unique():
    resultado = tabela_PL[tabela_PL['PROFESSOR'] == nome]
    agendamentos = len(resultado)
    # Verifica se há conteúdo previsto ou realizado não preenchido
    conteudo_previsto_vazio = resultado['CONTEÚDO PREVISTO'].isnull()
    conteudo_realizado_vazio = resultado['CONTEÚDO REALIZADO'].isnull()
    datas_nao_preenchidas_prof = resultado[conteudo_previsto_vazio |
                                           conteudo_realizado_vazio]
   

    
    #Busca as pendencias do professor pelos conteudos para gerar a planilha
    if not resultado.empty: 
        conteudo_previsto = resultado['CONTEÚDO PREVISTO'].isnull().sum()
        conteudo_realizado = resultado['CONTEÚDO REALIZADO'].isnull().sum()
    else: 
        conteudo_previsto = 0 
        conteudo_realizado = 0
        
    #Cria a aba de professores com pendencia     
    if conteudo_previsto > 0 or conteudo_realizado > 0 : 
        relatorio_vazios.append({'NOME':nome,
                                 'TOTAL AGENDAMENTOS':agendamentos,
                                 'Conteúdo Previsto': conteudo_previsto, 
                                 'Conteúdo Realizado': conteudo_realizado})   
        
   #Cria a aba de professores sem pendencia       
    else:
        relatorio_preenchidos.append({'NOME':nome,'TOTAL AGENDAMENTOS':agendamentos})

    # IF para pegar as colunas e criar a mensagem 
    if not datas_nao_preenchidas_prof.empty:
        for index, row in datas_nao_preenchidas_prof.iterrows():
            datas_nao_preenchidas.append({
                'PROFESSOR': nome,
                'TURMA': row['CODTURMA'],
                'DISCIPLINA': row['DISCIPLINA'],
                'DATA-DIA': row['DATA - Dia'],
                'DATA-MES': row['DATA - Mês'],
                'HORA INICIAL': row['HORAINICIAL'],
                'HORA FINAL': row['HORAFINAL'],
                'CONTEÚDO PREVISTO': row['CONTEÚDO PREVISTO'],
                'CONTEÚDO REALIZADO': row['CONTEÚDO REALIZADO']
            })

        
# Criar um DataFrame a partir dos resultados das buscas
df_relatorio_vazios = pd.DataFrame(relatorio_vazios)
df_relatorio_preenchidos = pd.DataFrame(relatorio_preenchidos)

# Criar a planilha com as abas "Vazios" e "Preenchidos" 
with pd.ExcelWriter('relatorio.xlsx') as writer:
    df_relatorio_vazios.to_excel(writer , sheet_name='Vazios', index= False)
    df_relatorio_preenchidos.to_excel(writer , sheet_name='Preenchidos', index= False)

print("Planilha de resultados gerada com sucesso")

# Configurações do servidor SMTP e informações do remetente
smtp_server = 'smtp-mail.outlook.com'
port = 587
sender_email = 'lucas.ramos@ms.senai.br'
password = senha
recipient_email = 'lukasmateusskt@gmail.com'

# Cria a mensagem com as informações das datas não preenchidas formatadas como uma tabela em HTML
mensagem = """
    <!DOCTYPE html>
<html>
    <head>
    <style>
        table {
            border-collapse: collapse;
            width: 100%;
        }
        th, td {
            border: 0.5px solid black;
            padding: 8px;
            text-align: left;
        }
        th {
            background-color: #f2f2f2;
        }
        td {
            padding-right: 10px; 
        }
    </style>
</head>


<body>
    <p>Boa noite professor,</p>
    <p>Abaixo, anexo o relatório de pendências de lançamento no SGE referentes ao preenchimento de conteúdo previsto e realizado de suas turmas.</p>
    <p>Considerando a obrigatoriedade de cumprimento de normas institucionais onde, dentre outras, consta a obrigatoriedade de “Registro no SGE de frequência, conteúdo previsto e conteúdo realizado diariamente” solicitamos que a regularização seja feita até amanhã 24/04/2024 haja vista o prazo para fechamento de produção.</p>
    <p>Reforço, ainda, que o não cumprimento das atividades docentes obrigatórias pode implicar em medidas disciplinares, conforme previsto no Regimento Escolar.</p>
    <p>Em tempo, caso o relatório apresente alguma inconsistência por favor relatar e evidenciar para que possamos pedir as correções cabíveis.</p>
    <p>Certos de podermos contar com sua colaboração agradecemos antecipadamente.</p>

      <table border='0.5'>
        <tr>
            <th>Código da Turma</th>
            <th>Disciplina</th>
            <th>Professor</th>
            <th>Data Dia</th>
            <th>Data Mês</th>
            <th>Hora Inicial</th>
            <th>Hora Final</th>
            <th>Conteúdo Previsto</th>
            <th>Conteúdo Realizado</th>
        </tr>
    

"""
for data_nao_preenchida in datas_nao_preenchidas:
    mensagem += f"""
    <tr>
        <td>{data_nao_preenchida['TURMA']}</td>
        <td>{data_nao_preenchida['DISCIPLINA']}</td>
        <td>{data_nao_preenchida['PROFESSOR']}</td>
        <td>{data_nao_preenchida['DATA-DIA']}</td>
        <td>{data_nao_preenchida['DATA-MES']}</td>
        <td>{data_nao_preenchida['HORA INICIAL']}</td>
        <td>{data_nao_preenchida['HORA FINAL']}</td>
        <td>{data_nao_preenchida['CONTEÚDO PREVISTO']}</td>
        <td>{data_nao_preenchida['CONTEÚDO REALIZADO']}</td>

    </tr>


"""
mensagem += """
        </table>
</body>
</html>

"""

# Cria o objeto MIMEMultipart
msg = MIMEMultipart()
msg['From'] = sender_email
msg['To'] = recipient_email
msg['Subject'] = 'Datas não preenchidas'
msg.attach(MIMEText(mensagem, 'html'))

# Envia o e-mail
try:
    server = smtplib.SMTP(smtp_server, port)
    server.starttls()
    server.login(sender_email, password)
    server.sendmail(sender_email, recipient_email, msg.as_string())
    server.quit()
    print("E-mail para o professor enviado com sucesso!")
except Exception as e:
    print(f"Erro ao enviar e-mail: {e}")