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

def tool_3():
    zona_col = next((col for col in df.columns if "estrato" in col.lower()), df.columns[0])
    colunas_excluir = [zona_col, 'Repetição', 'Área_amostrada']
    colunas_taxons = [col for col in df.columns if col not in colunas_excluir]
    abundancia = df[colunas_taxons].fillna(0).copy()
    grupos = df[zona_col].astype(str).copy()
    abundancia.index = df.index.astype(str)
    grupos.index = df.index.astype(str)
    try:
        dist_matrix = beta_diversity("braycurtis", abundancia, ids=abundancia.index)
        resultado = permanova(distance_matrix=dist_matrix, grouping=grupos, permutations=999)

        with open(caminho_pasta / "permanova_resultado.txt", "w") as f:
            f.write(str(resultado))
    except Exception as e:
        print(f"Erro ao executar PERMANOVA: {e}")