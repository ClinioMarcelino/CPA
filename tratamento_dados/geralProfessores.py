from docx import Document
import pandas as pd
import numpy as np
import difflib
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Cm
from docx2pdf import convert


materiasNaoIclusasNaPesquisa = pd.DataFrame(pd.read_excel('C:/Users/Clinio.freitas/PycharmProjects/pythonCPA2023-2/excel/MateriasNaoParticipantes.xlsx'))['DISCIPLINA'].tolist()
def verificaDisciplinasValidas(nomeDisciplina):
    if difflib.get_close_matches(nomeDisciplina, materiasNaoIclusasNaPesquisa) == []:
        return False
    return True

def alunosPorTurma():
    df = pd.DataFrame(pd.read_excel('C:/Users/Clinio.freitas/Documents/CPA/AvaliacaoDisciplina2023/ALUNOS.xlsx'))

    codTurmas = df['IDTURMADISC'].tolist()
    cod = df['IDTURMADISC'].unique().tolist()

    lista=[]

    for c in cod:
        codigos = dict()
        codigos['IDTURMADISC'] = c
        codigos['nAlunos'] = codTurmas.count(c)
        lista.append(codigos)

    return pd.DataFrame(lista)

alunoPTurma = alunosPorTurma()
def respostasDisc():
    df = pd.read_excel('C:/Users/Clinio.freitas/PycharmProjects/pythonCPA2023-2/excel/qAaDrespostas.xlsx')

    ra = []
    linhas = list()

    for key, values in df.iterrows():
        if values['ra'] not in ra:
            ra.append(values['ra'])
            for i in range(1, 10):
                if not pd.isnull(values[f'cd{i}']):
                    lin = dict()
                    lin[f'codDisc'] = values[f'cd{i}']
                    lin[f'nameDisc'] = values[f'nd{i}']
                    for j in range(0, 19):
                        lin[f'r{j + 1}'] = values[int(f'{(((i - 1) * 19) + j)}')]
                    linhas.append(lin)

    data = pd.DataFrame(linhas)
    data = data.dropna()
    data.replace('Não se Aplica', np.nan, inplace=True)

    disciplinas = dict()
    listaDisc = []

    for key, value in data.iterrows():
        if value['codDisc'] not in listaDisc:
            listaDisc.append(value['codDisc'])
            df = data[data['codDisc'] == value['codDisc']]
            respostas = dict()
            for chave, valor in df.iterrows():
                for k in range(1, 20):
                    # for x in df[f'r{k}'].tolist():
                    a = [x for x in df[f'r{k}'].tolist() if np.isnan(x) == False]
                    respostas[f'r{k}'] = a
            disciplinas[value['codDisc']] = respostas

    a = pd.DataFrame(disciplinas)
    return a

data = pd.read_excel('C:/Users/Clinio.freitas/PycharmProjects/pythonCPA2023-2/excel/respostasProfs.xlsx')

ra = []
nomes = []
profs = []
linhas = []

for key, values in data.iterrows():
    if values['nome'] not in nomes:
        nomes.append(values['nome'])
        for i in range(1, 8):
            if not pd.isnull(values[f'nd{i}']):
                lin = dict()
                lin['nome'] = values['nome'].title()
                lin['disc'] = values[f'nd{i}']
                for j in range(0, 16):
                    lin[f'rp{j+4}'] = values[int(f'{((i-1)*16)+j}')]
                # print(lin)
                linhas.append(lin)
dataRespProfs = pd.DataFrame(linhas)
dataRespProfs.replace('Não se Aplica', 'ÑA', inplace=True)
# print(dataRespProfs)


nomeProfs = []
Professores = []

dfProf = pd.DataFrame(pd.read_excel('C:/Users/Clinio.freitas/PycharmProjects/pythonCPA2023-2/excel/PROFESSORES.xlsx'))
disciplinas = respostasDisc()
listaDisc = []

for key,value in dfProf.iterrows():
    if value['PROFESSOR'] not in nomeProfs:
        nomeProfs.append(value['PROFESSOR'])
        for chave, valor in dfProf[dfProf['PROFESSOR'] == value['PROFESSOR']].iterrows():
            if verificaDisciplinasValidas(valor['DISCIPLINA']):
                continue
            Materia = dict()
            Materia['nome'] = value['PROFESSOR'].title()
            Materia['disc'] = valor['DISCIPLINA'].title()
            Materia['IDTURMADIS'] = valor['IDTURMADISC']
            Materia['CURSO'] = valor['CURSO']
            Materia['CODDISC'] = valor['CODDISC']
            Materia['TURNO'] = valor['TURNO'].title()
            Materia['TURMA'] = valor['TURMA']
            for k,v in dataRespProfs[dataRespProfs['nome']==Materia['nome']].iterrows():
                if v['disc'].title() == valor['DISCIPLINA'].title():
                    for i in range(4,20):
                        Materia[f'rp{i}'] = v[f'rp{i}']
            Professores.append(Materia)

Prof = pd.DataFrame(Professores)
# Prof = Prof.dropna()
Prof.replace(np.nan,'S/R')
print(Prof)

professores = Prof['nome'].unique().tolist()
buffados = []

for x in professores:
    aux = dict()
    teste = Prof[Prof['nome'] == x]
    a = teste['IDTURMADIS'].tolist()
    ax = []
    contador = 0
    cunta=0
    for i in a:
        # print(max(alunoPTurma[alunoPTurma['IDTURMADISC'] == i]["nAlunos"].values))
        try:
            contador += max(alunoPTurma[alunoPTurma['IDTURMADISC'] == i]["nAlunos"].values)
            for j in range(4,20):
                if j == 4:
                    cunta += len(disciplinas[i][f"r{j}"])
                for k in disciplinas[i][f"r{j}"]:
                    ax.append(k)
                # print(f'{j}  =  {disciplinas[i][f"r{j}"]}')
        except:
            # print(i)
            continue
    # print(ax)
    aux['Professor'] = x
    aux['Média'] = np.mean(ax)
    aux['SizeAmostra'] = len(ax)
    aux['nAlunos'] = contador
    aux['nAlunosRespondentes'] = cunta
    aux['%Respondentes'] = (cunta/contador)
    buffados.append((aux))

data = pd.DataFrame(buffados)

# data.to_excel('geralProfessores.xlsx',index=False)



