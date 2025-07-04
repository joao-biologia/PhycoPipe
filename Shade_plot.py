import os, subprocess, pandas as pd, platform, ezodf, re, seaborn as sns, numpy as np
from skbio.stats.distance import permanova, DistanceMatrix
from skbio.diversity import beta_diversity
from skbio.stats.ordination import pcoa
from pathlib import Path
import matplotlib.pyplot as plt
from matplotlib_venn import venn2, venn3

def tool_2():
    
    caminho_pasta = Path.home() / "Documents" / "PhycoPipe"
    caminho_input = caminho_pasta / "Inputs.ods"
    caminho_shade = caminho_pasta / "Shade_plot.png"
    df = pd.read_excel(str(caminho_input), sheet_name='Área de cobertura', header=0, engine='odf')
    print(df)
    
    zona_col = next((col for col in df.columns if "Amostra" in col.lower()), df.columns[0])
    if zona_col not in df.columns:
        print("Coluna de zona não encontrada.")
    for col in df.columns:
        if col not in ["Repetição", "Área_amostrada", zona_col]:
            df[col] = (df[col] / df["Área_amostrada"]) * 100
    colunas_excluir = [zona_col, 'Repetição', 'Área_amostrada']
    colunas_generos = [col for col in df.columns if col not in colunas_excluir]
    df_meltado = df.melt(id_vars=[zona_col], value_vars=colunas_generos,
                        var_name="Gênero", value_name="Área")
    df_meltado["Área"] = pd.to_numeric(df_meltado["Área"], errors='coerce')
    df_meltado = df_meltado.dropna(subset=["Área"])
    if df_meltado.empty:
        print("Nenhum dado numérico válido encontrado para gerar o gráfico.")
    ordem_zonas = ["zona superior", "zona intermediaria", "zona inferior"]
    df_meltado[zona_col] = pd.Categorical(df_meltado[zona_col], categories=ordem_zonas, ordered=True)
    df_resumo = df_meltado.groupby(["Gênero", zona_col], observed=False)["Área"].mean().unstack(fill_value=0)
    df_resumo = df_resumo[ordem_zonas]
    df_resumo = df_resumo.loc[df_resumo.mean(axis=1).sort_values(ascending=False).index]
    df_resumo = df_resumo.loc[df_resumo.sum(axis=1).sort_values(ascending=False).index]
    df_resumo.index = [rf"$\it{{{nome.replace('_', '\\ ')}}}$" for nome in df_resumo.index]
    plt.figure(figsize=(10, 6))
    ax = sns.heatmap(df_resumo, annot=True, fmt=".1f", cmap="YlGnBu",
                    cbar_kws={'label': '\nCobertura média relativa (%)'})
    ax.set_xlabel("")
    ax.xaxis.set_label_position('top')
    ax.xaxis.tick_top()

    plt.title("\n")
    plt.tight_layout()
    plt.savefig(caminho_shade, dpi=300)
    plt.show()

    print(f"\nShadeplot salvo em: {caminho_shade}")
