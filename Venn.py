import os, subprocess, pandas as pd, platform, ezodf, re, seaborn as sns, numpy as np
from skbio.stats.distance import permanova, DistanceMatrix
from skbio.diversity import beta_diversity
from skbio.stats.ordination import pcoa
from pathlib import Path
import matplotlib.pyplot as plt
from matplotlib_venn import venn2, venn3

def tool_1():
    
    caminho_pasta = Path.home() / "Documents" / "PhycoPipe"
    caminho_input = caminho_pasta / "Inputs.ods"
    caminho_venn = caminho_pasta / "Diagrama_venn.png"
    df = pd.read_excel(str(caminho_input), sheet_name='Área de cobertura', header=2, engine='odf')
    
    nomes_taxons = [re.sub(r'\s*[_\.]?[Ss][Pp]\.?$', '', col).strip() for col in df.columns[3:]]
    estratos_taxons = {}
    for _, linha in df.iterrows():
        estrato = linha.iloc[0]
        if pd.isna(estrato) or str(estrato).strip() == "":
            continue
        if estrato not in estratos_taxons:
            estratos_taxons[estrato] = set()
        for i, valor in enumerate(linha.iloc[3:]):
            try:
                valor_float = float(str(valor).replace(",", "."))
                if valor_float > 0:
                    estratos_taxons[estrato].add(nomes_taxons[i])
            except:
                continue
    if not estratos_taxons:
        print("Nenhum dado válido encontrado para gerar o diagrama.")
        return
    for k, v in estratos_taxons.items():
        print(f"{k}: {sorted(v)}")
    conjuntos = list(estratos_taxons.items())
    nomes_estratos = [k for k, _ in conjuntos]
    sets = [v for _, v in conjuntos]
    diversidades_alfa = [len(s) for s in sets]
    nomes_com_alpha = [f"{nome}\nα = {alpha}" for nome, alpha in zip(nomes_estratos, diversidades_alfa)]
    if len(sets) == 2:
        venn = venn2(sets, set_labels=nomes_com_alpha)
        region_keys = ['10', '01', '11']
    elif len(sets) == 3:
        venn = venn3(sets, set_labels=nomes_com_alpha)
        region_keys = ['100', '010', '001', '110', '101', '011', '111']
    else:
        print("O diagrama de Venn só suporta 2 ou 3 estratos.")
        return
    def get_region_taxa(key, sets):
        taxa = set()
        todos_taxons = set.union(*sets)
        for taxon in todos_taxons:
            presentes = [taxon in s for s in sets]
            binario = ''.join(['1' if p else '0' for p in presentes])
            if binario == key:
                taxa.add(taxon)
        return taxa
    for key in region_keys:
        label = venn.get_label_by_id(key)
        if label:
            taxa_na_regiao = get_region_taxa(key, sets)
            texto = '\n'.join(sorted(taxa_na_regiao))
            label.set_text(texto)
    plt.title("Distribuição dos Táxons por Estrato\n")
    plt.tight_layout()
    plt.savefig(caminho_venn, dpi=300)
    print(f"\nDiagrama de Venn salvo em: {caminho_venn}")