
## ========================================== ##
##                                            ##
##   Intellectual Property Rights (IPR):      ##
##                                            ##
##   João Paulo Andrade Barbosa (creator)     ##
##   2025-01-07 (date)                        ##
##                                            ##
## ========================================== ##


# ============================ #
#                              #
#      PhycoPipe (v1.0.0)      #
#                              #
# ============================ #

# Lybrary imports
import os, subprocess, pandas as pd, platform, ezodf, re, seaborn as sns, numpy as np
from skbio.stats.distance import permanova, DistanceMatrix
from skbio.diversity import beta_diversity
from pathlib import Path
import matplotlib.pyplot as plt
from matplotlib_venn import venn2, venn3

print(f'\nBem vindo ao PhycoPipe, uma ferramenta simples para trabalhos com comunidades de macroalgas\n')

# Main menu
def tools_menu():
    print(f'\n---- Tools Menu ----')
    print(f'(0) Sair')
    print(f'(1) Tabela de entradas')
    print(f'(2) Visualizar DataFrame')
    print(f'(3) Diagrama de Venn')
    print(f'(4) Shade plot')

def tools_menu_loop():    
    while True:
        tools_menu()
        choice = input("\nChoose one of my tools: ").strip()
        
        if  choice == '0':
            print(f'\nEncerrando...\n')
            break

        elif choice == '1':
            criar_planilha_ods()
            print("Preencha a tabela no LibreOffice e salve para uso posterior.\n")
            
        elif choice == '2':
            print(df)

        elif choice == '3':
            print(f'\nGerando resultados...\n')
            tool_1()
            
        elif choice == '4':
            print(f'\nGerando resultados...\n')
            tool_2()
        
        elif choice == '5':
            print(f'\nRealizando análise PERMANOVA...\n')
            tool_3()
            
        else:
            print(f'\n\nEscolha inválida.\n')

caminho_pasta = Path.home() / "Documents" / "PhycoPipe"
caminho_input = caminho_pasta / "Inputs.ods"
caminho_venn = caminho_pasta / "Venn_diagram.png"
caminho_shade = caminho_pasta / "Shade_plot.png"
df = pd.read_excel(str(caminho_input), sheet_name='Área de cobertura', header=2)

def criar_planilha_ods():
    caminho_pasta.mkdir(parents=True, exist_ok=True)
    ezodf.config.set_table_expand_strategy('all')
    planilha = ezodf.newdoc(doctype="ods", filename=str(caminho_input))
    folha = ezodf.Sheet('Área de cobertura', size=(30, 15))
    planilha.sheets += folha
    folha[0, 0].set_value("Área de cobertura (m²)")
    grupos = ["", "", "", "Chlorophyta", "Rhodophyta", "Phaeophyta"]
    for col, valor in enumerate(grupos):
        folha[1, col].set_value(valor)
    cabecalhos = [
        "Estrato (...)", "Repetição", "Área_amostrada",
        "Taxon 1", "Taxon 2", "Taxon 3"]
    for col, valor in enumerate(cabecalhos):
        folha[2, col].set_value(valor)

    planilha.save()

def tool_1():
    if df.shape[1] < 4:
        print("A tabela deve conter pelo menos 4 colunas.")
        return
    nomes_taxons = df.columns[3:]
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
        print("⚠️ O diagrama de Venn só suporta 2 ou 3 estratos.")
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

def tool_2():
    zona_col = next((col for col in df.columns if "estrato" in col.lower()), df.columns[0])
    if zona_col not in df.columns:
        print("Coluna de zona não encontrada.")
        return
    colunas_excluir = [zona_col, 'Repetição', 'Área_amostrada']
    colunas_generos = [col for col in df.columns if col not in colunas_excluir]
    df_meltado = df.melt(id_vars=[zona_col], value_vars=colunas_generos,
                         var_name="Gênero", value_name="Área")
    df_meltado["Área"] = pd.to_numeric(df_meltado["Área"], errors='coerce')
    df_meltado = df_meltado.dropna(subset=["Área"])
    if df_meltado.empty:
        print("Nenhum dado numérico válido encontrado para gerar o gráfico.")
        return
    df_resumo = df_meltado.groupby(["Gênero", zona_col])["Área"].sum().unstack(fill_value=0)
    df_resumo = df_resumo.loc[df_resumo.sum(axis=1).sort_values(ascending=False).index]
    plt.figure(figsize=(10, 6))
    sns.heatmap(df_resumo, annot=True, fmt=".1f", cmap="YlGnBu",
                cbar_kws={'label': '\nÁrea de cobertura (cm²)'})
    plt.title("\nShade Plot - Cobertura de algas por zona\n")
    plt.xlabel("Zona")
    plt.tight_layout()
    plt.savefig(caminho_shade, dpi=300)
    print(f"\nShadeplot salvo em: {caminho_shade}")

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
        print("\n🔬 Resultado da PERMANOVA:")
        print(resultado)
    except Exception as e:
        print(f"❌ Erro ao executar PERMANOVA: {e}")

if __name__ == "__main__":
    tools_menu_loop()
