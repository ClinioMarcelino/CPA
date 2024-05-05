import polars as pl
import pandas as pd
class nao_avaliadas():
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