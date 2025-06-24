tool_1 <- function() {
  library(readODS)
  library(dplyr)
  library(stringr)
  library(ggVennDiagram)
  library(ggplot2)
  
  # Caminhos
  caminho_pasta <- file.path(Sys.getenv("HOME"), "Documents", "PhycoPipe")
  caminho_input <- file.path(caminho_pasta, "Inputs.ods")
  caminho_venn <- file.path(caminho_pasta, "Diagrama_venn.png")
  
  # Leitura com cabeçalho técnico ignorado
  df_raw <- read_ods(caminho_input, sheet = "Área de cobertura", col_names = FALSE)
  
  # A terceira linha (linha 3) contém o nome real das colunas
  nomes_colunas <- as.character(df_raw[3, ])
  nomes_colunas <- str_trim(str_replace_all(nomes_colunas, "\\s*[_\\.]?[Ss][Pp]\\.?$", ""))
  df <- df_raw[-c(1:3), ]
  colnames(df) <- nomes_colunas
  
  # Conversão de números e limpeza
  df <- df %>%
    mutate(across(where(is.character), ~ str_replace_all(., ",", "."))) %>%
    mutate(across(-1:-3, as.numeric))  # ignora colunas 1 a 3
  
  # Agrupamento por zona e verificação de presença (>0)
  zonas <- unique(df[[1]])
  nomes_taxons <- colnames(df)[4:ncol(df)]
  estratos_taxons <- list()
  
  for (zona in zonas) {
    subdf <- df[df[[1]] == zona, nomes_taxons, drop = FALSE]
    presentes <- colSums(subdf > 0, na.rm = TRUE)
    taxons_presentes <- nomes_taxons[presentes > 0]
    estratos_taxons[[as.character(zona)]] <- taxons_presentes
  }
  
  # Verifica quantidade de zonas válidas
  if (length(estratos_taxons) < 2 || length(estratos_taxons) > 3) {
    cat("O diagrama de Venn só suporta 2 ou 3 zonas.\n")
    return()
  }
  
  # Cálculo de diversidade alfa
  diversidades_alfa <- sapply(estratos_taxons, length)
  nomes_com_alpha <- paste0(names(estratos_taxons), "\nα = ", diversidades_alfa)
  names(estratos_taxons) <- nomes_com_alpha
  
  # Gera o diagrama com nomes dos táxons nas interseções
  plot <- ggVennDiagram(
    estratos_taxons,
    label = "both",  # nomes dos táxons dentro das interseções
    label_geom = "label",
    label_alpha = 0
  ) +
    scale_fill_gradient(low = "#F0F8FF", high = "#4682B4") +
    theme_void() +
    labs(title = "Distribuição dos Táxons por Estrato")
  
  # Exporta como PNG
  ggsave(caminho_venn, plot = plot, width = 8, height = 8, dpi = 300)
  cat(sprintf("\n✅ Diagrama de Venn salvo em: %s\n", caminho_venn))
}

tool_1()