import configparser
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd

# leitura das planilhas
tabela_PL = pd.read_excel('data.xlsx')
tabela_nome = pd.read_excel('nome_profs.xlsx')

config = configparser.ConfigParser()
config.read('config.ini')

senha = config['DEFAULT']['MINHA_SENHA']


relatorio_vazios = []
relatorio_preenchidos = []


# Tratamento de valores NaN
tabela_PL.fillna('', inplace=True)

# Configurações do servidor SMTP e informações do remetente
smtp_server = 'smtp-mail.outlook.com'
port = 587
sender_email = 'lucas.ramos@ms.senai.br'
password = senha

for _, professor in tabela_nome.iterrows():

    nome = professor['NOME']
    email = professor['EMAIL']

    datas_nao_preenchidas = []

    resultado = tabela_PL[tabela_PL['PROFESSOR'] == nome]
    agendamentos = len(resultado)
    # Verifica se há conteúdo previsto ou realizado não preenchido
    conteudo_previsto_vazio = resultado['CONTEÚDO PREVISTO'] == ''
    conteudo_realizado_vazio = resultado['CONTEÚDO REALIZADO'] == ''
    datas_nao_preenchidas_prof = resultado[conteudo_previsto_vazio |
                                           conteudo_realizado_vazio]

    # Busca as pendencias do professor pelos conteudos para gerar a planilha
    if not datas_nao_preenchidas_prof.empty:
        conteudo_previsto = conteudo_previsto_vazio.sum()
        conteudo_realizado = conteudo_realizado_vazio.sum()
        relatorio_vazios.append({'NOME': nome,
                                 'TOTAL AGENDAMENTOS': agendamentos,
                                 'Conteúdo Previsto': conteudo_previsto,
                                 'Conteúdo Realizado': conteudo_realizado})
        
   # Cria a aba de professores sem pendencia
    else:
        relatorio_preenchidos.append(
            {'NOME': nome, 'TOTAL AGENDAMENTOS': agendamentos})

   
 
    # IF para pegar as colunas e criar a mensagem
    if not datas_nao_preenchidas_prof.empty:
        for index, row in datas_nao_preenchidas_prof.iterrows():
            datas_nao_preenchidas.append({
                'PROFESSOR': nome,
                'TURMA': row['CODTURMA'],
                'DISCIPLINA': row['DISCIPLINA'],
                'DATA-DIA': int(row['DATA - Dia']),
                'DATA-MES': row['DATA - Mês'],
                'HORA INICIAL': row['HORAINICIAL'],
                'HORA FINAL': row['HORAFINAL'],
                'CONTEÚDO PREVISTO': row['CONTEÚDO PREVISTO'],
                'CONTEÚDO REALIZADO': row['CONTEÚDO REALIZADO']
            })

        # Cria o objeto MIMEMultipart
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] =  email
        msg['Subject'] = 'Pendência de Lançamento SGE'
        
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
        # Cria a mensagem com as informações das datas não preenchidas formatadas como uma tabela em HTML
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
        msg.attach(MIMEText(mensagem, 'html'))

        # Envia o e-mail
        try:
            server = smtplib.SMTP(smtp_server, port)
            server.starttls()
            server.login(sender_email, password)
            server.sendmail(sender_email, email, msg.as_string())
            server.quit()
            print(f"E-mail enviado para {nome} ({email})")

        except Exception as e:
            print(f"Erro ao enviar e-mail para {nome} ({email}): {e}")


         

# Criar um DataFrame a partir dos resultados das buscas
df_relatorio_vazios = pd.DataFrame(relatorio_vazios)
df_relatorio_preenchidos = pd.DataFrame(relatorio_preenchidos)

# Criar a planilha com as abas "Vazios" e "Preenchidos"
with pd.ExcelWriter('relatorio.xlsx') as writer:
    df_relatorio_vazios.to_excel(writer, sheet_name='Vazios', index=False)
    df_relatorio_preenchidos.to_excel(writer, sheet_name='Preenchidos', index=False)

print("Planilha de resultados gerada com sucesso")
