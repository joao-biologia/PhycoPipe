import os, subprocess, pandas as pd, platform, ezodf, re, seaborn as sns, numpy as np
from skbio.stats.distance import permanova, DistanceMatrix
from skbio.diversity import beta_diversity
from skbio.stats.ordination import pcoa
from pathlib import Path
import matplotlib.pyplot as plt
from matplotlib_venn import venn2, venn3
from skbio.diversity.alpha import shannon, simpson
from scipy.stats import entropy

def tool_6():
    
    caminho_pasta = Path.home() / "Documents" / "PhycoPipe"
    caminho_input = caminho_pasta / "abundancia2.ods"
    caminho_indice = caminho_pasta / "indice2.ods"
    df = pd.read_excel(str(caminho_input), sheet_name='Área de cobertura', header=1, engine='odf')
    print(df)

    def shannon_index(row):
        data = row[row > 0]
        total = data.sum()
        proportions = data / total
        return entropy(proportions, base=np.e)

    def riqueza(row):
        return (row > 0).sum()

    def calcular_diversidades(df):
        zonas = df['Estrato_(Zona)'].unique()
        resultados = []

        for zona in zonas:
            df_zona = df[df['Estrato_(Zona)'] == zona]
            dados_abundancia = df_zona.iloc[:, 2:]  # Dados das espécies

            # Diversidade alfa (média de Shannon por linha)
            alfa = dados_abundancia.apply(shannon_index, axis=1).mean()

            # Diversidade gama (Shannon da soma total da zona)
            soma_total = dados_abundancia.sum()
            gama = shannon_index(soma_total)

            # Diversidade beta (Whittaker: beta = gama / alfa)
            beta = gama / alfa if alfa != 0 else np.nan

            resultados.append({
                'Zona': zona,
                'Diversidade_Alfa': round(alfa, 4),
                'Diversidade_Gama': round(gama, 4),
                'Diversidade_Beta': round(beta, 4)
            })

        return pd.DataFrame(resultados)

    # Calcular as diversidades
    df_resultado = calcular_diversidades(df)

    # Salvar em .ods
    from pandas import ExcelWriter

    output_path = 'diversidades.ods'
    with ExcelWriter(output_path, engine='odf') as writer:
        df_resultado.to_excel(writer, index=False, sheet_name='Diversidade')

    print(f"Arquivo '{output_path}' salvo com sucesso.")
