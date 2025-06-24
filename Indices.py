import os, subprocess, pandas as pd, platform, ezodf, re, seaborn as sns, numpy as np
from skbio.stats.distance import permanova, DistanceMatrix
from skbio.diversity import beta_diversity
from skbio.stats.ordination import pcoa
from pathlib import Path
import matplotlib.pyplot as plt
from matplotlib_venn import venn2, venn3
from skbio.diversity.alpha import shannon, simpson

def shannon(values, base=np.e):
    proportions = values[values > 0] / np.sum(values)
    return -np.sum(proportions * np.log(proportions)) / np.log(base)

def simpson(values):
    proportions = values[values > 0] / np.sum(values)
    return 1 - np.sum(proportions ** 2)

def dominance(values):
    proportions = values[values > 0] / np.sum(values)
    return np.sum(proportions ** 2)

def riqueza(values):
    return np.sum(values > 0)

def tool_5():
    caminho_pasta = Path.home() / "Documents" / "PhycoPipe"
    caminho_input = caminho_pasta / "abundancia.ods"
    caminho_indice = caminho_pasta / "indice.txt"

    # Lê a planilha .ods
    df = pd.read_excel(str(caminho_input), sheet_name='Área de cobertura', header=1, engine='odf')
    print(df)

    # Identifica a coluna da zona (estrato)
    zona_col = next((col for col in df.columns if "estrato" in col.lower()), df.columns[0])
    colunas_excluir = [zona_col, 'Repetição', 'Área_amostrada']
    colunas_taxons = [col for col in df.columns if col not in colunas_excluir]

    resultados = []

    for idx, linha in df.iterrows():
        abundancias = linha[colunas_taxons].fillna(0).astype(float).values
        zona = linha[zona_col]

        shannon_h = shannon(abundancias)
        simpson_1d = simpson(abundancias)
        dominancia_d = dominance(abundancias)
        riqueza_s = riqueza(abundancias)

        resultados.append({
            "Comunidade": zona,
            "Shannon (H)": round(shannon_h, 3),
            "Simpson 1-D": round(simpson_1d, 3),
            "Dominance (D)": round(dominancia_d, 3),
            "Riqueza (S)": round(riqueza_s, 3)
        })

    resultados_df = pd.DataFrame(resultados)

    # Impressão no terminal
    print("\n✅ Índices de diversidade por comunidade:\n")
    print(resultados_df.to_string(index=False))

    # Salvando no arquivo indice.txt
    try:
        with open(caminho_indice, "w", encoding="utf-8") as f:
            f.write("Índices de diversidade por comunidade\n\n")
            f.write(resultados_df.to_string(index=False))
        print(f"\n📄 Resultados salvos com sucesso em: {caminho_indice}")
    except Exception as e:
        print(f"\n❌ Erro ao salvar o arquivo: {e}")
