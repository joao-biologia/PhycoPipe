
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
import os, subprocess, pandas as pd, platform, ezodf
from pathlib import Path

print(f'\nBem vindo ao PhycoPipe, uma ferramenta simples para trabalhos com comunidades de macroalgas\n')

# Main menu
def tools_menu():
    print(f'\n---- Tools Menu ----')
    print(f'(0) Sair')
    print(f'(1) Tabela de entradas')
    print(f'(2) Rodar todos os processos')
    print(f'(3) Diagrama de Venn')

def tools_menu_loop():    
    while True:
        tools_menu()
        choice = input("\nChoose one of my tools: ").strip()
        
        if  choice == '0':
            print(f'\nExiting...\n')
            break

        elif choice == '1':
            nome_pasta = "PhycoPipe"
            nome_arquivo = "Inputs.ods"
            caminho_pasta = criar_diretorio_em_documentos(nome_pasta)
            caminho_arquivo = caminho_pasta / nome_arquivo
            criar_planilha_ods(caminho_arquivo)
            abrir_pasta_no_explorer(caminho_pasta) 
            abrir_arquivo(caminho_arquivo)          
            print(f"\nTabela criada em: {caminho_arquivo}")
            print("Preencha a tabela no LibreOffice e salve para uso posterior.\n")

        elif choice == '2':
            print(f'\nGerando resultados...\n')
            nome_pasta = "PhycoPipe"
            nome_arquivo = "Inputs.ods"
            caminho_arquivo = criar_diretorio_em_documentos(nome_pasta) / nome_arquivo
            tool_1(caminho_arquivo)

        elif choice == '3':
            print(f'\nGerando resultados...\n')
            nome_pasta = "PhycoPipe"
            nome_arquivo = "Inputs.ods"
            caminho_arquivo = criar_diretorio_em_documentos(nome_pasta) / nome_arquivo
            tool_1(caminho_arquivo)
        
        else:
            print(f'\n\nInvalid choice. Please, try again.\n')

def obter_diretorio_documentos():
    one_drive_docs = Path(os.path.expandvars(r'%USERPROFILE%\OneDrive\Documents'))
    if one_drive_docs.exists():
        return one_drive_docs
    return Path(os.path.expanduser('~')) / 'Documents'

def criar_diretorio_em_documentos(nome_pasta):
    documentos = obter_diretorio_documentos()
    caminho = documentos / nome_pasta
    caminho.mkdir(parents=True, exist_ok=True)
    return caminho

def abrir_pasta_no_explorer(caminho: Path):
    if platform.system() == 'Windows':
        os.startfile(str(caminho))

def criar_planilha_ods(caminho_arquivo):
    ezodf.config.set_table_expand_strategy('all')

    planilha = ezodf.newdoc(doctype="ods", filename=str(caminho_arquivo))
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

libreoffice_path = r"C:\Program Files\LibreOffice\program\soffice.exe"
    
def abrir_arquivo(caminho):
    if Path(libreoffice_path).exists():
        subprocess.Popen([libreoffice_path, str(caminho)])
    else:
        print("LibreOffice não encontrado no caminho padrão. Tente abrir manualmente ou verifique a instalação.")

from matplotlib import pyplot as plt
from matplotlib_venn import venn2, venn3
import ezodf
import re

def limpar_taxon(taxon):
    if not isinstance(taxon, str):
        return ""
    return re.sub(r'\s*[_\.]?sp\.?', '', taxon, flags=re.IGNORECASE)

def tool_1(caminho_arquivo_ods):
    print(f"\n🔍 Lendo dados da planilha: {caminho_arquivo_ods}")
    doc = ezodf.opendoc(str(caminho_arquivo_ods))
    sheet = doc.sheets[0]

    dados = []
    for row in sheet.rows():
        linha = [cell.value if cell.value is not None else "" for cell in row]
        dados.append(linha)

    for i, linha in enumerate(dados):
        if linha[0] and "strato" in linha[0].lower():
            idx_cabecalho = i
            break
    else:
        print("❌ Cabeçalho com nomes de colunas não encontrado.")
        return

    nomes_colunas = dados[idx_cabecalho]
    nomes_taxons_originais = nomes_colunas[3:]
    nomes_taxons = [limpar_taxon(t) for t in nomes_taxons_originais]

    estratos_taxons = {}

    for linha in dados[idx_cabecalho+1:]:
        if len(linha) < 4:
            continue
        estrato = linha[0]
        if not estrato:
            continue
        if estrato not in estratos_taxons:
            estratos_taxons[estrato] = set()
        for i, valor in enumerate(linha[3:]):
            try:
                valor_float = float(str(valor).replace(",", "."))
                if valor_float > 0:
                    estratos_taxons[estrato].add(nomes_taxons[i])
            except:
                continue

    for k, v in estratos_taxons.items():
        print(f"{k}: {sorted(v)}")

    conjuntos = list(estratos_taxons.items())
    nomes_estratos = [k for k, _ in conjuntos]
    sets = [v for _, v in conjuntos]

    if len(sets) == 2:
        venn = venn2(sets, set_labels=nomes_estratos)
        region_keys = ['10', '01', '11']
    elif len(sets) == 3:
        venn = venn3(sets, set_labels=nomes_estratos)
        region_keys = ['100', '010', '001', '110', '101', '011', '111']
    else:
        print("⚠️ O diagrama de Venn só suporta até 3 estratos.")
        return

    # Lógica para descobrir os táxons de cada região do diagrama
    def get_region_taxa(key, sets):
        taxa = set()
        for i in range(len(sets[0])):  # máximo número de táxons
            # Vamos testar cada táxon individualmente
            pass
        todos_taxons = set.union(*sets)
        for taxon in todos_taxons:
            presentes = [taxon in s for s in sets]
            binario = ''.join(['1' if p else '0' for p in presentes])
            if binario == key:
                taxa.add(taxon)
        return taxa

    # Substitui os números pelas listas de táxons
    for key in region_keys:
        label = venn.get_label_by_id(key)
        if label:
            taxa_na_regiao = get_region_taxa(key, sets)
            texto = '\n'.join(sorted(taxa_na_regiao))
            label.set_text(texto)

    plt.title("Distribuição dos Táxons por Estrato")
    plt.tight_layout()
    plt.show()

import pandas as pd
import ezodf
import re
from upsetplot import from_memberships, UpSet
import matplotlib.pyplot as plt

def limpar_taxon(taxon):
    return re.sub(r'\s*[_\.]?sp\.?', '', str(taxon), flags=re.IGNORECASE)

def tool_2(caminho_arquivo_ods):
    # Abrir planilha ODS
    ezodf.config.set_table_expand_strategy('all')
    doc = ezodf.opendoc(str(caminho_arquivo_ods))
    sheet = doc.sheets[0]

    # Extrair os dados
    dados = []
    for row in sheet.rows():
        dados.append([cell.value if cell.value is not None else "" for cell in row])

    if len(dados) < 2:
        print("❌ Planilha sem dados suficientes.")
        return

    cabecalhos = dados[0]
    nomes_taxons = [limpar_taxon(t) for t in cabecalhos[3:]]

    df = pd.DataFrame(dados[1:], columns=cabecalhos)

    # Mapear presença de táxons por estrato
    presencas = {}
    for _, row in df.iterrows():
        estrato = row[0]
        if not estrato:
            continue
        for taxon, valor in zip(nomes_taxons, row[3:]):
            if isinstance(valor, (int, float)) and valor > 0:
                presencas.setdefault(taxon, set()).add(estrato)

    # Gerar memberships para o UpSet plot
    memberships = [tuple(sorted(v)) for v in presencas.values()]
    upset_data = from_memberships(memberships)

    # Plot
    plt.figure(figsize=(10, 6))
    UpSet(upset_data, show_counts=True).plot()
    plt.title("Distribuição dos Táxons entre os Estratos")
    plt.tight_layout()
    plt.savefig(caminho_pasta, dpi=300)
    plt.show()

    
if __name__ == "__main__":
    tools_menu_loop()
