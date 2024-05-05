import polars as pl
import pandas as pd
from statistics import mode

#RAIZ ARQUIVO PRINCIPAL ALUNOS
r_principal_alunos = 'C:/Users/clini/Documents/CPA/avaliacaoDisciplinas/2024.1/CPA_GRD241_02052024.xlsx'

def drop_nao_avaliadas(df):
    data = list()
    nao_av_list = pl.read_excel('C:/Users/clini/Documents/CPA/avaliacaoDisciplinas/2024.1/disciplinasNaoAvaliadas.xlsx')['DISCIPLINA'].to_list()

    for values in df.iter_rows():
        if values[2] in nao_av_list:
            continue

        linha = dict()
        linha['CURSO'] = values[0]
        linha['IDTURMADISC'] = values[1]
        linha['DISCIPLINA'] = values[2]
        linha['RA'] = values[7]
        linha['ALUNO'] = values[8]
        linha["EMAIL"] = values[10]
        data.append(linha)


    return pd.DataFrame(data)


linkBase = ''

df = pl.read_excel(r_principal_alunos)
df = df.drop('TELEFONE')
df = df.drop('PERIODO_LETIVO')
df = df.drop('CODDISC')
df = df.drop('EMAIL')
df = drop_nao_avaliadas(df)
# print(df)

# df_nao_avaliadas = pl.read_excel('C:/Users/clini/Documents/CPA/avaliacaoDisciplinas/2024.1/disciplinasNaoAvaliadas.xlsx')

ra_list = df['RA'].to_list()
ra_most_appearence = mode(ra_list)
num_ra_most_appearences = ra_list.count(ra_most_appearence)

print(num_ra_most_appearences)
print(ra_most_appearence)

# LISTAS E VARIAVEIS AUXILIARES
listaRAaux = []
dataSaida = []
contagemAlunos = 0

# PRINT PARA AJUDAR NA VIZUALIZAÇÃO DO CABEÇALHO DA TABELA
# print(df.columns.tolist())
'''['CURSO', 'IDTURMADISC', 'DISCIPLINA', 'RA', 'ALUNO', 'EMAIL']'''

def gerarLinkAluno(dt1, linkBase):
    link=''

    for i in range(1,int((len(dt1.keys())-4)/2)):
        link += f'cd{i}={dt1[f"CD{i}"]}&'

    for i in range(1,int((len(dt1.keys())-4)/2)):
        link += f'nd{i}={dt1[f"ND{i}"]}&'

    link += f'name={dt1["NOME"]}&ra=00{dt1["RA"]}'
    linkBase += link
    return linkBase

for key, value in df.iterrows():
    # print(value)
    if value['RA'] not in listaRAaux:
        listaRAaux.append(value['RA'])
        Aluno = dict.fromkeys(['RA', 'NOME', 'EMAIL', 'LINK'])  # Iniciam como None
        Aluno['RA'] = value['RA']
        Aluno['NOME'] = value['ALUNO'].title()
        Aluno['EMAIL'] = value['EMAIL']
        contagemAlunos += 1
        i = 0
        for chave, valor in df[df['RA'] == value['RA']].iterrows():
                Aluno[f'CD{i}'] = valor['IDTURMADISC']
                Aluno[f'ND{i}'] = valor['DISCIPLINA']
                i += 1
        Aluno["LINK"] = gerarLinkAluno(Aluno, linkBase)
        dataSaida.append(Aluno) # INCLUSÃO DO DICIONÁRIO DE CADA DISCENTE NA LISTA DE SAÍDA

        # Print de acompanhamento e verificação de quantidade de docentes
        print(contagemAlunos)

# GERAÇÃO DE DATA FRAME DA LISTAGEM DE DISCENTES
data = pd.DataFrame(dataSaida)

# GERAÇÃO DO XLSX DE SAÍDA ENVIADO PARA O SETOR DE COMUNICAÇÃO FAZER A DISTRIBUIÇÃO PARA PESQUISA
data.to_excel('C:/Users/clini/Documents/CPA/avaliacaoDisciplinas/2024.1/links/links_Discentes_2024_1.xlsx')
