import matplotlib.pyplot as plt
import seaborn as sns
from skbio.diversity import beta_diversity
from skbio.stats.ordination import pcoa
from pathlib import Path
import pandas as pd

caminho_pasta = Path.home() / "Documents" / "PhycoPipe"
caminho_input = caminho_pasta / "Inputs.ods"
df = pd.read_excel(str(caminho_input), sheet_name='Área de cobertura', header=0, engine='odf')

def tool_4():
    print(df)
    zona_col = next((col for col in df.columns if "Amostra" in col.lower()), df.columns[0])

    colunas_excluir = [zona_col, 'Repetição', 'Área_amostrada']
    colunas_taxons = [col for col in df.columns if col not in colunas_excluir]

    abundancia = df[colunas_taxons].fillna(0).copy()
    grupos = df[zona_col].astype(str).copy()

    abundancia.index = df.index.astype(str)
    grupos.index = df.index.astype(str)

    try:
        dist_matrix = beta_diversity("braycurtis", abundancia, ids=abundancia.index)
        pcoa_result = pcoa(dist_matrix)
        coords = pcoa_result.samples.iloc[:, 0:2]

        coords = coords.copy()
        coords[zona_col] = grupos.values

        # Paleta de cores personalizada (verde, azul, laranja/vermelho)
        cores_personalizadas = {
            "zona superior oeste": "#2ca02c",       # verde escuro
            "zona superior leste": "#98df8a",       # verde claro
            "zona intermediaria oeste": "#1f77b4",  # azul escuro
            "zona intermediaria leste": "#aec7e8",  # azul claro
            "zona inferior oeste": "#d64b19",       # laranja
            "zona inferior leste": "#ef997d",       # vermelho
        }

        plt.figure(figsize=(8, 6))
        sns.scatterplot(
            data=coords,
            x=coords.columns[0],
            y=coords.columns[1],
            hue=zona_col,
            palette=cores_personalizadas,
            s=100,
            edgecolor='k'
        )
        plt.legend(title=None)
        plt.xlabel(f"PCoA 1 ({pcoa_result.proportion_explained.iloc[0]*100:.1f}%)")
        plt.ylabel(f"PCoA 2 ({pcoa_result.proportion_explained.iloc[1]*100:.1f}%)")
        plt.title("PCoA - Bray-Curtis")
        plt.axhline(0, color='gray', lw=0.5)
        plt.axvline(0, color='gray', lw=0.5)
        plt.tight_layout()
        plt.savefig(caminho_pasta / "PCoA_plot.png", dpi=300)
        plt.show()

    except Exception as e:
        print(f"Erro ao executar PCoA: {e}")
