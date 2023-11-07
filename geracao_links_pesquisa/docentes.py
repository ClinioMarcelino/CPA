import pandas as pd
import difflib

# RAIZ E NOME DO ARQUIVO
arquivo = '.xlsx'

# INSERINDO O LINK BASE GERADO PARA PESQUISA
link = ''

# LENDO O XLSX DOS DOCENTES RECEBIDO DO DBA
df = pd.DataFrame(pd.read_excel(arquivo))

# LISTAS E VARIAVEIS AUXILIARES
listaProfessores = []
dadosSaida = []
contagemProfessores = 0

# PRINT PARA AJUDAR NA VIZUALIZAÇÃO DO CABEÇALHO DA TABELA
'''print(df.columns.tolist())
['PERIODO_LETIVO', 'CURSO', 'IDTURMADISC', 'CODDISC', 'DISCIPLINA', 'TURNO', 'TURMA', 'PROFESSOR', 'EMAIL']'''

# GERAÇÃO DA LISTAGEM DE MATÉRIAS NÃO PARTICIPANTES NA PESQUISA. XLSX RECEBIDO DAS COORDENAÇÕES
materiasNaoIclusasNaPesquisa = pd.DataFrame(pd.read_excel('.xlsx'))['DISCIPLINA'].tolist()

# FUNÇÃO PARA VERIFICAÇÃO SE A DISCIPLINA ENTRA OU NÃO NA PESQUISA
def verificaDisciplinasValidas(nomeDisciplina):
    if difflib.get_close_matches(nomeDisciplina, materiasNaoIclusasNaPesquisa,1,0.8) == []:
        return False
    return True


# FUNÇÃO PARA GERAR LINK INDIVIDUAL DA PESQUISA
def gerarLinkProfessores(dt1):
    linky=''

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
        Professor['LINK'] = link
        listaDisciplinasprof = []
        # Contador de codigos de disciplinas
        i=1
        for chave,valor in df[df['PROFESSOR'] == value['PROFESSOR']].iterrows():
            if valor['DISCIPLINA'] not in listaDisciplinasprof:
                listaDisciplinasprof.append(valor['DISCIPLINA'])
                if not verificaDisciplinasValidas(valor['DISCIPLINA']):
                    Professor[f'ND{i}'] = valor['DISCIPLINA']
                    i+=1
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
data.to_excel('.xlsx', index=False)