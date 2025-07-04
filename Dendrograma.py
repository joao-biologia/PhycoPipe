import pandas as pd
import numpy as np
from pathlib import Path
from skbio.diversity.alpha import shannon, simpson
from skbio.diversity import beta_diversity
from skbio.stats.ordination import pcoa
from skbio.stats.distance import DistanceMatrix
from scipy.spatial.distance import pdist, squareform
import ezodf
import matplotlib.pyplot as plt
from scipy.cluster.hierarchy import linkage, dendrogram

def tool_8():
    
    caminho_pasta = Path.home() / "Documents" / "PhycoPipe"
    caminho_input = caminho_pasta / "Inputs.ods"
    caminho_saida = caminho_pasta / "Dendrograma.png"
    df = pd.read_excel(str(caminho_input), sheet_name='Área de cobertura', header=0, engine='odf')
    print(df)

    colunas_especies = df.columns[3:]
    df_zonas = df.groupby("Amostra")[colunas_especies].sum()

    presenca_ausencia = (df_zonas > 0).astype(int)
    dist_presenca = pdist(presenca_ausencia, metric='jaccard')
    Z_presenca = linkage(dist_presenca, method='average')

    dist_abundancia = pdist(df_zonas, metric='braycurtis')
    Z_abundancia = linkage(dist_abundancia, method='average')

    fig, axes = plt.subplots(1, 2, figsize=(14, 6))

    dendrogram(Z_presenca, labels=presenca_ausencia.index.tolist(), leaf_rotation=90, ax=axes[0])
    axes[0].set_title("Dendrograma por Presença/Ausência")
    axes[0].set_ylabel("Distância (Jaccard)")

    dendrogram(Z_abundancia, labels=df_zonas.index.tolist(), leaf_rotation=90, ax=axes[1])
    axes[1].set_title("Dendrograma por Abundância")
    axes[1].set_ylabel("Distância (Bray-Curtis)")

    plt.tight_layout()
    plt.savefig(caminho_saida, dpi=300)
    plt.show()

    print("\nMatriz de presença/ausência (binária):")
    print(presenca_ausencia)

    print("\nMatriz de distâncias (Jaccard - Presença/Ausência):")
    print(squareform(dist_presenca))

    print("\nMatriz de abundância (soma das espécies por zona):")
    print(df_zonas)

    print("\nMatriz de distâncias (Bray-Curtis - Abundância):")
    print(squareform(dist_abundancia))

    print("\nMatriz de linkage (Presença/Ausência):")
    print(Z_presenca)
    print()
    print(df_zonas.index)
    
    print("\nMatriz de linkage (Abundância):")
    print(Z_abundancia)
    print()
    print(df_zonas.index)

tool_8()