import pandas as pd 
import os

#leitura das planilhas
tabela_PL = pd.read_excel('data.xlsx')
tabela_nome = pd.read_excel('nome_profs.xlsx')



relatorio_vazios = []
relatorio_preenchidos = []



for nome in tabela_nome['NOME'].unique():
    resultado = tabela_PL[tabela_PL['PROFESSOR'] == nome]
    agendamentos = len(resultado)
    
    #Busca as pendencias do professor pelos conteudos 
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
        
# Criar um DataFrame a partir dos resultados das buscas
df_relatorio_vazios = pd.DataFrame(relatorio_vazios)
df_relatorio_preenchidos = pd.DataFrame(relatorio_preenchidos)

# Criar a planilha com as abas "Vazios" e "Preenchidos" 
with pd.ExcelWriter('relatorio.xlsx') as writer:
    df_relatorio_vazios.to_excel(writer , sheet_name='Vazios', index= False)
    df_relatorio_preenchidos.to_excel(writer , sheet_name='Preenchidos', index= False)

print("Planilha de resultados gerada com sucesso")
