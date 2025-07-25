
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
from skbio.stats.ordination import pcoa
from pathlib import Path
import matplotlib.pyplot as plt
from matplotlib_venn import venn2, venn3
from skbio.diversity.alpha import shannon, simpson

print(f'\nBem vindo ao PhycoPipe, uma ferramenta simples para trabalhos com comunidades de macroalgas\n')

caminho_pasta = Path.home() / "Documents" / "PhycoPipe"
caminho_input = caminho_pasta / "Inputs.ods"

make_dirs = True
if make_dirs:
    try:
        os.makedirs(os.path.join(Path.home(), "Documents", "PhycoPipe"))
    except FileExistsError:
        pass

caminho_pasta.mkdir(parents=True, exist_ok=True)
ezodf.config.set_table_expand_strategy('all')

if not caminho_input.exists():
    planilha = ezodf.newdoc(doctype="ods", filename=str(caminho_input))
    folha = ezodf.Sheet('Área de cobertura', size=(30, 15))
    planilha.sheets += folha

    cabecalhos = [
        "Amostra", "Repetição", "Área_amostrada",
        "Taxon 1", "Taxon 2", "Taxon 3"
    ]
    for col, valor in enumerate(cabecalhos):
        folha[0, col].set_value(valor)

    planilha.save()

df = pd.read_excel(str(caminho_input), sheet_name='Área de cobertura', header=0, engine='odf')

# Main menu
def tools_menu():
    print(f'\n---- Tools Menu ----')
    print(f'(0) Sair')
    print(f'(1) Tabela de entradas')
    print(f'(2) Visualizar DataFrame')
    print(f'(3) Diagrama de Venn')
    print(f'(4) Shade plot')
    print(f'(6) PCoA')
    print(f'(7) Indices de Diversidade e Heterogeneidade')
    print(f'(8) Dendrograma')

def tools_menu_loop():    
    
    while True:
        
        tools_menu()
        choice = input("\nChoose one of my tools: ").strip()
        
        if  choice == '0':
            print(f'\nEncerrando...\n')
            exit()

        elif choice == '1':
            print("Preencha a tabela no LibreOffice e salve para uso posterior.\n")
            
        elif choice == '2':
            print("\n\n")
            print(df)

        elif choice == '3':
            print(f'\nPresente apenas para terminal R...\n')
            
        elif choice == '4':
            print(f'\nGerando resultados...\n')
            import Shade_plot
            Shade_plot.tool_2()
        
        elif choice == '6':
            print(f'Performando PCoA...\n')
            import PCoA
            PCoA.tool_4()
            
        elif choice == '7':
            print(f'\nCalculando índices de diversidade e heterogeneidade...\n')
            import Indices2
            Indices2.tool_6()
            
        elif choice == '8':
            print(f'\nGerando dendograma...\n')
            import Dendrograma
            Dendrograma.tool_7()

        else:
            print(f'\n\nEscolha inválida.\n')
    
if __name__ == "__main__":
    tools_menu_loop()

