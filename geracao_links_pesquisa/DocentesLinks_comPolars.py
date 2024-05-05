import polars as pl
import pandas as pd

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
        linha['PROFESSOR'] = values[3]
        linha["EMAIL"] = values[4]
        data.append(linha)


    return pd.DataFrame(data)


# RAIZ E NOME DO ARQUIVO
r_pricipal_docentes = "C:/Users/clini/Documents/CPA/avaliacaoDisciplinas/2024.1/CPA_PROFESSORES_GRD241.xlsx"

# INSERINDO O LINK BASE GERADO PARA PESQUISA
link_base = ''

# LENDO O XLSX DOS DOCENTES RECEBIDO DO DBA
df = pl.read_excel(r_pricipal_docentes)

# LISTAS E VARIAVEIS AUXILIARES
listaProfessores = []
dadosSaida = []
contagemProfessores = 0

# PRINT PARA AJUDAR NA VIZUALIZAÇÃO DO CABEÇALHO DA TABELA
'''print(df.columns)
['PERIODO_LETIVO', 'CURSO', 'IDTURMADISC', 'CODDISC', 'DISCIPLINA', 'TURNO', 'TURMA', 'PROFESSOR', 'EMAIL']'''

df = df.drop('TURMA')
df = df.drop('PERIODO_LETIVO')
df = df.drop('CODDISC')
df = df.drop('TURNO')
df = drop_nao_avaliadas(df)

print(df)


def gerarLinkProfessores(dt1):
    linky='wwwwwww//'

    for i in range(1,int((len(dt1.keys())-2))):
        linky += f'nd{i}={dt1[f"ND{i}"]}&'

    linky += f'nome={dt1["NOME"]}'

    return linky


for key,value in df.iterrows():
    if value['PROFESSOR'] not in listaProfessores:
        listaProfessores.append(value['PROFESSOR'])
        Professor = dict()
        Professor['NOME'] = value['PROFESSOR'].title()
        Professor['EMAIL'] = value['EMAIL']
        Professor['LINK'] = link_base
        listaDisciplinasprof = []
        # Contador de codigos de disciplinas
        i=1
        for chave, valor in df[df['PROFESSOR'] == value['PROFESSOR']].iterrows():
            if valor['DISCIPLINA'] not in listaDisciplinasprof:
                listaDisciplinasprof.append(valor['DISCIPLINA'])
                Professor[f'ND{i}'] = valor['DISCIPLINA']
                i += 1
        # SE NÃO HOUVER SOMA O DOCENTE NÃO TEVE MATÉRIA A SER CONSIDERADA NA AVALIAÇÃO E NÃO ENTRA NA LISTAGEM FINAL
        if i ==1:
            continue

        Professor['LINK'] += gerarLinkProfessores(Professor)
        dadosSaida.append(Professor) # INCLUSÃO DO DICIONÁRIO DE CADA DOCENTE NA LISTA DE SAÍDA
        contagemProfessores += 1
        # Print de acompanhamento e verificação de quantidade de docentes
        print(contagemProfessores)

# GERAÇÃO DE DATA FRAME DA LISTAGEM DE DISCENTES
data = pd.DataFrame(dadosSaida)

# GERAÇÃO DO XLSX DE SAÍDA ENVIADO PARA O SETOR DE COMUNICAÇÃO FAZER A DISTRIBUIÇÃO PARA PESQUISA
data.to_excel('C:/Users/clini/Documents/CPA/avaliacaoDisciplinas/2024.1/links/links_Docentes_2024_1.xlsx', index=False)