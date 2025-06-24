import os, subprocess, pandas as pd, platform, ezodf, re, seaborn as sns, numpy as np
from skbio.stats.distance import permanova, DistanceMatrix
from skbio.diversity import beta_diversity
from skbio.stats.ordination import pcoa
from pathlib import Path
import matplotlib.pyplot as plt
from matplotlib_venn import venn2, venn3
from skbio.diversity.alpha import shannon, simpson

caminho_pasta = Path.home() / "Documents" / "PhycoPipe"
caminho_input = caminho_pasta / "Inputs.ods"
df = pd.read_excel(str(caminho_input), sheet_name='Área de cobertura', header=2, engine='odf')

def tool_4():

    zona_col = next((col for col in df.columns if "estrato" in col.lower()), df.columns[0])

    colunas_excluir = [zona_col, 'Repetição', 'Área_amostrada']
    colunas_taxons = [col for col in df.columns if col not in colunas_excluir]

    abundancia = df[colunas_taxons].fillna(0).copy()
    grupos = df[zona_col].astype(str).copy()

    abundancia.index = df.index.astype(str)
    grupos.index = df.index.astype(str)

    try:
        dist_matrix = beta_diversity("braycurtis", abundancia, ids=abundancia.index)
        pcoa_result = pcoa(dist_matrix)
        coords = pcoa_result.samples.iloc[:, 0:2]  # PC1 e PC2

        coords = coords.copy()
        coords[zona_col] = grupos.values


        plt.figure(figsize=(8, 6))
        sns.scatterplot(data=coords, x=coords.columns[0], y=coords.columns[1],
                        hue=zona_col, palette="Set2", s=100, edgecolor='k')
        plt.xlabel(f"PCoA 1 ({pcoa_result.proportion_explained.iloc[0]*100:.1f}%)")
        plt.ylabel(f"PCoA 2 ({pcoa_result.proportion_explained.iloc[1]*100:.1f}%)")
        plt.title("PCoA - Bray-Curtis")
        plt.axhline(0, color='gray', lw=0.5)
        plt.axvline(0, color='gray', lw=0.5)
        plt.tight_layout()
        plt.savefig(caminho_pasta / "PCoA_plot.png", dpi=300)

    except Exception as e:
        print(f"Erro ao executar PCoA: {e}")