tool_1 <- function() {
  library(readODS)
  library(dplyr)
  library(stringr)
  library(ggVennDiagram)
  library(ggplot2)
  
  caminho_pasta <- file.path(Sys.getenv("HOME"), "Documents", "PhycoPipe")
  caminho_input <- file.path(caminho_pasta, "Inputs.ods")
  caminho_venn <- file.path(caminho_pasta, "Diagrama_venn.png")
  
  df_raw <- read_ods(caminho_input, sheet = "Área de cobertura", col_names = FALSE)
  
  nomes_colunas <- as.character(df_raw[3, ])
  nomes_colunas <- str_trim(str_replace_all(nomes_colunas, "\\s*[_\\.]?[Ss][Pp]\\.?$", ""))
  df <- df_raw[-c(1:3), ]
  colnames(df) <- nomes_colunas
  
  df <- df %>%
    mutate(across(where(is.character), ~ str_replace_all(., ",", "."))) %>%
    mutate(across(-1:-3, as.numeric))

  zonas <- unique(df[[1]])
  nomes_taxons <- colnames(df)[4:ncol(df)]
  estratos_taxons <- list()
  
  for (zona in zonas) {
    subdf <- df[df[[1]] == zona, nomes_taxons, drop = FALSE]
    presentes <- colSums(subdf > 0, na.rm = TRUE)
    taxons_presentes <- nomes_taxons[presentes > 0]
    estratos_taxons[[as.character(zona)]] <- taxons_presentes
  }
  
  if (length(estratos_taxons) < 2 || length(estratos_taxons) > 3) {
    cat("O diagrama de Venn só suporta 2 ou 3 zonas.\n")
    return()
  }
  
  diversidades_alfa <- sapply(estratos_taxons, length)
  nomes_com_alpha <- paste0(names(estratos_taxons), "\nα = ", diversidades_alfa)
  names(estratos_taxons) <- nomes_com_alpha
  
  plot <- ggVennDiagram(
    estratos_taxons,
    label = "both",
    label_geom = "label",
    label_alpha = 0
  ) +
    scale_fill_gradient(low = "#F0F8FF", high = "#4682B4") +
    theme_void() +
    labs(title = "Distribuição dos Táxons por Estrato")
  
  ggsave(caminho_venn, plot = plot, width = 8, height = 8, dpi = 300)
  cat(sprintf("\n✅ Diagrama de Venn salvo em: %s\n", caminho_venn))
}

tool_1()