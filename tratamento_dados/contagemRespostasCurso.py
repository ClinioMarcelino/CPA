import pandas as pd

df = pd.read_excel('C:/Users/Clinio.freitas/PycharmProjects/pythonCPA2023-2/excel/CursosAlunos.xlsx')

totalCursos = dict()

for curso in df['CURSO'].unique().tolist():
    # print(f'{curso} :{df["CURSO"].tolist().count(curso)}')
    totalCursos[curso] = df["CURSO"].tolist().count(curso)

# print(totalCursos)

data = pd.DataFrame.from_dict(totalCursos, orient='index', columns=['nAlunos'])
data = data.rename_axis('Cursos').reset_index()
data['nRespostas'] = 0

raRespostas = pd.read_excel('C:/Users/Clinio.freitas/PycharmProjects/pythonCPA2023-2/excel/raRespondidosAdD.xlsx')['ra'].unique().tolist()

ra = []

for key, value in df.iterrows(): # AVALIAÇÃO DE DISCIPLINAS
    if value['RA'] not in ra:
        ra.append(value['RA'])
        if value['RA'] in raRespostas:
            for chave,valor in data[data['Cursos'] == value['CURSO']].iterrows():
                data['nRespostas'].iat[chave] += 1

print(data)

data.to_excel('C:/Users/Clinio.freitas/PycharmProjects/pythonCPA2023-2/saidas/contagemRespostasPCurso.xlsx', index=False)