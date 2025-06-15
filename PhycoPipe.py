
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
import os, subprocess, pandas as pd, platform, ezodf, re
from pathlib import Path
import matplotlib.pyplot as plt
from matplotlib_venn import venn2, venn3

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

        elif choice == '3':
            print(f'\nGerando resultados...\n')
            nome_pasta = "PhycoPipe"
            nome_arquivo_ods = "Inputs.ods"
            nome_arquivo_saida = "Diagrama_de_Venn.png"

            caminho_pasta = criar_diretorio_em_documentos(nome_pasta)
            caminho_arquivo_ods = caminho_pasta / nome_arquivo_ods
            caminho_arquivo_saida = caminho_pasta / nome_arquivo_saida

            tool_1(caminho_arquivo_ods, caminho_arquivo_saida)
        
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

def limpar_taxon(taxon):
    if not isinstance(taxon, str):
        return ""
    return re.sub(r'\s*[_\.]?sp\.?', '', taxon, flags=re.IGNORECASE)

def tool_1(caminho_arquivo_ods, caminho_arquivo):
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

    diversidades_alfa = [len(s) for s in sets]
    nomes_com_alpha = [f"{nome}\nα = {alpha}" for nome, alpha in zip(nomes_estratos, diversidades_alfa)]

    if len(sets) == 2:
        venn = venn2(sets, set_labels=nomes_com_alpha)
        region_keys = ['10', '01', '11']
    elif len(sets) == 3:
        venn = venn3(sets, set_labels=nomes_com_alpha)
        region_keys = ['100', '010', '001', '110', '101', '011', '111']
    else:
        print("⚠️ O diagrama de Venn só suporta até 3 estratos.")
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
    plt.savefig(caminho_arquivo, dpi=300)
    plt.show()

if __name__ == "__main__":
    tools_menu_loop()
