import pandas as pd
import numpy as np
from pathlib import Path
from skbio.diversity.alpha import shannon, simpson
from skbio.diversity import beta_diversity
from skbio.stats.ordination import pcoa
from skbio.stats.distance import DistanceMatrix
from scipy.spatial.distance import pdist, squareform
import ezodf

def calcular_indices_por_zona(df, colunas_especies):
    
    zonas_ordenadas = ["zona superior oeste", "zona superior leste", "zona intermediaria oeste", "zona intermediaria leste", "zona inferior oeste", "zona inferior leste"]
    df["Estrato_(Zona)"] = pd.Categorical(df["Estrato_(Zona)"], categories=zonas_ordenadas, ordered=True)
    df_agg = df.groupby("Estrato_(Zona)")[colunas_especies].sum()
    resultados = []

    for zona in zonas_ordenadas:
        
        abundancias = df_agg.loc[zona].values
        abundancias = abundancias[abundancias > 0]
        p = abundancias / abundancias.sum()
        shannon = -np.sum(p * np.log(p))
        S = len(abundancias)
        equitabilidade = shannon / np.log(S) if S > 1 else np.nan
        dominancia = np.sum(p ** 2)
        simpson = 1 - dominancia

        resultados.append({
            "Comunidade": zona,
            "Shannon (H')": shannon,
            "Equitabilidade (E)": equitabilidade,
            "Simpson (1-D)": simpson,
            "Dominância (D)": dominancia
        })

    return pd.DataFrame(resultados)

def tool_6():
    
    caminho_pasta = Path.home() / "Documents" / "PhycoPipe"
    caminho_input = caminho_pasta / "Inputs3.ods"
    caminho_saida = caminho_pasta / "indice_diversidade.ods"

    df = pd.read_excel(str(caminho_input), sheet_name='Área de cobertura', header=2, engine='odf')

    colunas_especies = df.columns[3:]

    df["Riqueza"] = (df[colunas_especies] > 0).sum(axis=1)

    riqueza_zona = (df[colunas_especies] > 0).groupby(df["Estrato_(Zona)"]).any().sum(axis=1)
    df["Riqueza (S)"] = df["Estrato_(Zona)"].map(riqueza_zona)

    gama_total = (df[colunas_especies] > 0).any().sum()
    df["Div. Gama"] = gama_total
    
    zonas_ordenadas = ["zona superior oeste", "zona superior leste", "zona intermediaria oeste", "zona intermediaria leste", "zona inferior oeste", "zona inferior leste"]
    df["Estrato_(Zona)"] = pd.Categorical(df["Estrato_(Zona)"], categories=zonas_ordenadas, ordered=True)

    df["Div. Alfa"] = df["Riqueza"]
    df["Comunidade"] = df["Estrato_(Zona)"]
    
    df_zona = df.groupby("Comunidade").agg({
    "Riqueza (S)": "first",
    "Div. Gama": "first",
    "Div. Alfa": "mean"
    }).reset_index()
    df_zona["Div. Beta"] = df_zona["Div. Gama"] / df_zona["Div. Alfa"]

    indices_zonas = calcular_indices_por_zona(df, colunas_especies)

    print(df[["Comunidade", "Repetição", "Riqueza", "Riqueza (S)", "Div. Gama"]])
    print(df_zona[["Comunidade", "Riqueza (S)", "Div. Gama", "Div. Alfa", "Div. Beta"]])
    print(indices_zonas)
    
    with pd.ExcelWriter(caminho_saida, engine='odf') as writer:
        df.to_excel(writer, sheet_name='Dados_Completos', index=False)
        df_zona.to_excel(writer, sheet_name='Resumo_Zonas', index=False)
        indices_zonas.to_excel(writer, sheet_name='Indices_Zonas', index=False)
        
    print(f'\nResultados salvos em {caminho_saida}\n')

