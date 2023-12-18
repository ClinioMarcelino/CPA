import pandas as pd
import numpy as np
import matplotlib.pyplot as mp

rootAlunos = 'C:/Users/Clinio.freitas/Documents/CPA/AvaliacaoDisciplina2023/ALUNOS.xlsx'
rootRespostas = 'C:/Users/Clinio.freitas/PycharmProjects/pythonCPA2023-2/excel/qAaDrespostas.xlsx'
rootCursos = 'C:/Users/Clinio.freitas/PycharmProjects/pythonCPA2023-2/excel/CursosAlunos.xlsx'

dfAlunos = pd.read_excel(rootAlunos)
dfRespostas = pd.read_excel(rootRespostas)
dfCursos = pd.read_excel(rootCursos)

cursos = dfCursos['CURSO'].unique().tolist()
# print(cursos)

ra = []
linhas = list()

for key, values in dfRespostas.iterrows():
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

respostasDisciplina = dict()
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
        respostasDisciplina[value['codDisc']] = respostas

dfDisciplinas = pd.DataFrame(respostasDisciplina)
# print(dfDisciplinas)


geral = []
for curso in cursos:
    aux = dict()
    teste = dfAlunos[dfAlunos['CURSO'] == curso]
    aux['Curso'] = curso
    a = teste['IDTURMADISC'].unique().tolist()
    # ax = []
    contador = 0
    cunta=0
    # print(max(alunoPTurma[alunoPTurma['IDTURMADISC'] == i]["nAlunos"].values))


    for j in range(4,20):
        listaResposta = []
        for codTurma in a:
            try:
                for resp in respostasDisciplina[codTurma][f"r{j}"]:
                    listaResposta.append(resp)
            except:
                continue
        aux[f'P{j-3}'] = np.nanmean(listaResposta)



    # aux['SizeAmostra'] = len(ax)
    # aux['nAlunos'] = contador
    # aux['nAlunosRespondentes'] = cunta
    # aux['%Respondentes'] = (cunta/contador)
    geral.append(aux)

data = pd.DataFrame(geral)
print(data)
# data.to_excel('geralCursos.xlsx',index=False)
# data.set_index('Curso', inplace=True)
# data = data.dropna()
# data = data.T



# # mp.style.use('fivethirtyeight')
# for key,value in data.iterrows():
#     value.plot(kind='bar')
#     mp.title(key)
#     mp.show()
#
# # data.plot() #criando gráfico
# # # mp.title('SEU TÍTULO LINDO') #adicionando o título
# # # mp.ylabel('Cursos')
# # # mp.legend(loc='best') #colocando a legenda no melhor lugar
# mp.show() #mostrando gráfico

