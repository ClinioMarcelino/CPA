import pandas as pd
import difflib

# INSERINDO O LINK BASE GERADO PARA PESQUISA
linkBase = ''

# LENDO O XLSX DOS DISCENTES RECEBIDO DO DBA
df = pd.DataFrame(pd.read_excel('.xlsx'))

# LISTAS E VARIAVEIS AUXILIARES
listaRAaux = []
dataSaida = []
contagemAlunos = 0

# PRINT PARA AJUDAR NA VIZUALIZAÇÃO DO CABEÇALHO DA TABELA
# print(df.columns.tolist())
''' 'PERIODO_LETIVO', 'CURSO', 'CODDISC', 'IDTURMADISC', 'DISCIPLINA', 'TURNO', 'TURMA', 'STATUS_PLETIVO',
'STATUS_DISCIPLINA', 'RA', 'ALUNO', 'PERIODO_ALUNO', 'TELEFONE', 'EMAIL', 'EMAILCORP' '''

# GERAÇÃO DA LISTAGEM DE MATÉRIAS NÃO PARTICIPANTES NA PESQUISA. XLSX RECEBIDO DAS COORDENAÇÕES
materiasNaoIclusasNaPesquisa = pd.DataFrame(pd.read_excel('.xlsx'))['DISCIPLINA'].tolist()

# FUNÇÃO PARA VERIFICAÇÃO SE A DISCIPLINA ENTRA OU NÃO NA PESQUISA
def verificaDisciplinasValidas(nomeDisciplina):
    # USANDO A BIBLIOTECA DIFFLIB PARA ELIMINAR AS PEQUENAS VARIAÇÕES NO SEU CADASTRO, DEPENDEDO DO CURSO
    if difflib.get_close_matches(nomeDisciplina, materiasNaoIclusasNaPesquisa,1,0.8) == []:
        return False
    return True


# FUNÇÃO PARA GERAR LINK INDIVIDUAL DA PESQUISA
def gerarLinkAluno(dt1, linkBase):
    link=''

    for i in range(1,int((len(dt1.keys())-4)/2)+1):
        link += f'cd{i}={dt1[f"CD{i}"]}&'

    for i in range(1,int((len(dt1.keys())-4)/2)+1):
        link += f'nd{i}={dt1[f"ND{i}"]}&'

    link += f'name={dt1["NOME"]}&ra=00{dt1["RA"]}'
    linkBase += link
    return linkBase




for line, value in df.iterrows():
    if value["RA"] not in listaRAaux:
        listaRAaux.append(value["RA"])
        Aluno = dict.fromkeys(['RA', 'NOME', 'EMAIL', 'LINK'])  # Iniciam como None
        Aluno['RA'] = value['RA']
        Aluno['NOME'] = value['ALUNO'].title()
        Aluno['EMAIL'] = value['EMAILCORP']
        contagemAlunos += 1
        # Contador de codigos de disciplinas
        i = 1
        disc=0
        for linha, valor in df[df['RA'] == value['RA']].iterrows():
            if not verificaDisciplinasValidas(valor['DISCIPLINA']):
                Aluno[f'CD{i}'] = valor['IDTURMADISC']
                Aluno[f'ND{i}'] = valor['DISCIPLINA']
                i += 1
        # SE NÃO HOUVER SOMA O DISCENTE NÃO TEVE MATÉRIA A SER CONSIDERADA NA AVALIAÇÃO E NÃO ENTRA NA LISTAGEM FINAL
        if i == 1:
            continue

        Aluno["LINK"] = gerarLinkAluno(Aluno,linkBase)
        dataSaida.append(Aluno) # INCLUSÃO DO DICIONÁRIO DE CADA DISCENTE NA LISTA DE SAÍDA

        # Print de acompanhamento e verificação de quantidade de docentes
        print(contagemAlunos)

# GERAÇÃO DE DATA FRAME DA LISTAGEM DE DISCENTES
data = pd.DataFrame(dataSaida)

# GERAÇÃO DO XLSX DE SAÍDA ENVIADO PARA O SETOR DE COMUNICAÇÃO FAZER A DISTRIBUIÇÃO PARA PESQUISA
data.to_excel('.xlsx', index=False)