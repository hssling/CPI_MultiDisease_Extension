# =============================================================================
# CPI Multi-Disease Extension: Core Signature Analysis
# Identifies the "Universal Primed Signature" shared across TB, Sepsis, and Dengue
# =============================================================================

library(dplyr)
library(ggplot2)
library(data.table)

# Configuration
BASE_DIR <- "d:/research-automation/TB multiomics/TB Chromatin Priming Multiomics/CPI_MultiDisease_Extension"
TB_DIR <- "d:/research-automation/TB multiomics/TB Chromatin Priming Multiomics/v2_extracted/TB-Chromatin-Priming-Multiomics_v2"
setwd(BASE_DIR)

RESULTS_DIR <- file.path(BASE_DIR, "3_results")
dir.create(file.path(RESULTS_DIR, "core_signature"), recursive = TRUE, showWarnings = FALSE)

message("=", paste(rep("=", 60), collapse = ""), "=")
message("CPI CORE SIGNATURE ANALYSIS")
message("=", paste(rep("=", 60), collapse = ""), "=")

# =============================================================================
# [1/4] Load Data
# =============================================================================
message("\n[1/4] Loading Data...")

# 1. ATAC Reference (Priming Map)
peak_links <- fread(file.path(BASE_DIR, "config/peak_gene_links.csv"))
primed_genes_ref <- unique(peak_links$gene)
message(sprintf("  Loaded Reference: %d potential primed genes", length(primed_genes_ref)))

# 2. TB DEGs (BAL + PBMC)
tb_files <- c(
  file.path(TB_DIR, "4_results/tables/DEG_all_celltypes.csv"), # BAL
  file.path(TB_DIR, "4_results/tables/DEG_GSE287288_DTB.csv")   # PBMC
)

tb_genes <- c()
for (f in tb_files) {
  if (file.exists(f)) {
    dt <- fread(f)
    # Check column names (gene/feature)
    gene_col <- if("gene" %in% names(dt)) "gene" else "feature"
    if (gene_col %in% names(dt)) {
      # Filter sig
      sig <- dt %>% filter(p_val_adj < 0.05) %>% pull(!!sym(gene_col))
      tb_genes <- c(tb_genes, sig)
    }
  }
}
tb_primed <- intersect(unique(tb_genes), primed_genes_ref)
message(sprintf("  TB (BAL+PBMC): %d primed genes found", length(tb_primed)))

# 3. Sepsis DEGs
sepsis_file <- file.path(RESULTS_DIR, "tables/DEG_Sepsis.csv")
sepsis_primed <- c()
if (file.exists(sepsis_file)) {
  dt <- fread(sepsis_file)
  sig <- dt %>% filter(p_val_adj < 0.05)
  sepsis_primed <- intersect(unique(sig$gene), primed_genes_ref)
  message(sprintf("  Sepsis: %d primed genes found", length(sepsis_primed)))
} else {
  message("  Warning: Sepsis DEG file missing")
}

# 4. Dengue DEGs
dengue_file <- file.path(RESULTS_DIR, "tables/DEG_Dengue.csv")
dengue_primed <- c()
if (file.exists(dengue_file)) {
  dt <- fread(dengue_file)
  sig <- dt %>% filter(p_val_adj < 0.05)
  dengue_primed <- intersect(unique(sig$gene), primed_genes_ref)
  message(sprintf("  Dengue: %d primed genes found", length(dengue_primed)))
} else {
  message("  Warning: Dengue DEG file missing")
}

# =============================================================================
# [2/4] Intersection Analysis (Venn)
# =============================================================================
message("\n[2/4] Identifying Core Signature...")

# Find overlap
core_signature <- intersect(intersect(tb_primed, sepsis_primed), dengue_primed)
message(sprintf("  CORE SIGNATURE (Shared by All 3): %d genes", length(core_signature)))

# Pairwise
tb_sepsis <- intersect(tb_primed, sepsis_primed)
tb_dengue <- intersect(tb_primed, dengue_primed)
sepsis_dengue <- intersect(sepsis_primed, dengue_primed)

message(sprintf("  TB-Sepsis Overlap: %d genes", length(tb_sepsis)))
message(sprintf("  TB-Dengue Overlap: %d genes", length(tb_dengue)))
message(sprintf("  Sepsis-Dengue Overlap: %d genes", length(sepsis_dengue)))

# Export Core Genes
core_df <- data.frame(gene = core_signature)
fwrite(core_df, file.path(RESULTS_DIR, "core_signature", "Core_Primed_Signature_Genes.csv"))

# Create Venn Diagram (Simple implementation without heavy deps)
create_venn_data <- function(tb, sepsis, dengue) {
  all_genes <- unique(c(tb, sepsis, dengue))
  res <- data.frame(
    gene = all_genes,
    TB = all_genes %in% tb,
    Sepsis = all_genes %in% sepsis,
    Dengue = all_genes %in% dengue
  )
  return(res)
}

venn_df <- create_venn_data(tb_primed, sepsis_primed, dengue_primed)
fwrite(venn_df, file.path(RESULTS_DIR, "core_signature", "Gene_Disease_Matrix.csv"))

# Try to plot if ggvenn installed, else skip
if (require("ggvenn", quietly=TRUE)) {
  venn_list <- list(TB = tb_primed, Sepsis = sepsis_primed, Dengue = dengue_primed)
  p_venn <- ggvenn(venn_list, fill_color = c("#377EB8", "#4DAF4A", "#FF7F00")) +
    ggtitle("Universal Primed Signature")
  ggsave(file.path(RESULTS_DIR, "core_signature", "Fig_Core_Signature_Venn.png"), p_venn)
} else {
  message("  Note: 'ggvenn' not installed. Skipping Venn plot (data saved).")
}

# =============================================================================
# [3/4] Analyze Core Function (Top Genes)
# =============================================================================
message("\n[3/4] Analyzing Core Genes...")
print(head(core_signature, 20))

# Retrieve expression stats (LFC) for these genes to make a heatmap
get_lfc <- function(file, genes, name) {
  if (!file.exists(file)) return(NULL)
  dt <- fread(file)
  # Filter for core genes
  dt_sub <- dt %>% filter(gene %in% genes)
  # Average LFC across cell types/comparisons if multiple entries per gene
  dt_summ <- dt_sub %>% group_by(gene) %>% summarise(mean_LFC = mean(avg_log2FC), .groups="drop")
  names(dt_summ)[2] <- name
  return(dt_summ)
}

# Extract LFCs
tb_lfc <- get_lfc(tb_files[2], core_signature, "LFC_TB_PBMC") # Use PBMC for better comparison
sepsis_lfc <- get_lfc(sepsis_file, core_signature, "LFC_Sepsis")
dengue_lfc <- get_lfc(dengue_file, core_signature, "LFC_Dengue")

# Merge
if (!is.null(tb_lfc)) {
  lfc_matrix <- tb_lfc
  if (!is.null(sepsis_lfc)) lfc_matrix <- full_join(lfc_matrix, sepsis_lfc, by="gene")
  if (!is.null(dengue_lfc)) lfc_matrix <- full_join(lfc_matrix, dengue_lfc, by="gene")

  # Replace NA with 0
  lfc_matrix[is.na(lfc_matrix)] <- 0
  
  # Save Matrix
  fwrite(lfc_matrix, file.path(RESULTS_DIR, "core_signature", "Core_Signature_LFC_Matrix.csv"))
  
  # Heatmap
  lfc_long <- melt(data.table(lfc_matrix), id.vars="gene", variable.name="Disease", value.name="LFC")
  lfc_long$Disease <- gsub("LFC_", "", lfc_long$Disease)
  
  # Select top 30 variable genes for plot clarity
  top_genes <- lfc_matrix$gene[order(apply(lfc_matrix[,-1], 1, var), decreasing=TRUE)][1:min(30, nrow(lfc_matrix))]
  lfc_plot <- lfc_long %>% filter(gene %in% top_genes)
  
  p_heat <- ggplot(lfc_plot, aes(x = Disease, y = gene, fill = LFC)) +
    geom_tile(color="white") +
    scale_fill_gradient2(low = "blue", mid = "white", high = "red") +
    labs(title = "Core Primed Signature Expression", subtitle = "Top shared genes (Log2FC)") +
    theme_minimal() +
    theme(axis.text.y = element_text(size=8))
  
  ggsave(file.path(RESULTS_DIR, "core_signature", "Fig_Core_Signature_Heatmap.png"), p_heat)
}

# =============================================================================
# [4/4] Summary and Pathway Suggestion
# =============================================================================
message("\n", paste(rep("=", 60), collapse = ""))
message("CORE SIGNATURE ANALYSIS COMPLETE")
message(sprintf("  Universal Genes Identified: %d", length(core_signature)))
message("  See results in: ", file.path(RESULTS_DIR, "core_signature"))
message(paste(rep("=", 60), collapse = ""))

# Try simple enrichment if gprofiler2 exists
if (require("gprofiler2", quietly=TRUE)) {
  message("  Running Pathway Enrichment (g:Profiler)...")
  gost_res <- gost(query = core_signature, organism = "hsapiens", sources = c("GO:BP", "KEGG", "REAC"))
  if (!is.null(gost_res$result)) {
    result_table <- gost_res$result %>% select(term_id, term_name, p_value, source) %>% arrange(p_value)
    print(head(result_table, 10))
    fwrite(result_table, file.path(RESULTS_DIR, "core_signature", "Enrichment_Core_Signature.csv"))
  }
} else {
  message("  Recommend running GO/KEGG enrichment on 'Core_Primed_Signature_Genes.csv'.")
}
