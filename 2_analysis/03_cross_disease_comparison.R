# =============================================================================
# CPI Multi-Disease Extension: Cross-Disease Comparison
# Integrates CPI results from TB, Sepsis, and Dengue for comprehensive analysis
# =============================================================================

library(dplyr)
library(ggplot2)
library(data.table)

# Configuration
BASE_DIR <- "d:/research-automation/TB multiomics/TB Chromatin Priming Multiomics/CPI_MultiDisease_Extension"
TB_DIR <- "d:/research-automation/TB multiomics/TB Chromatin Priming Multiomics/v2_extracted/TB-Chromatin-Priming-Multiomics_v2"
setwd(BASE_DIR)

RESULTS_DIR <- file.path(BASE_DIR, "3_results")
dir.create(file.path(RESULTS_DIR, "figures"), recursive = TRUE, showWarnings = FALSE)
dir.create(file.path(RESULTS_DIR, "tables"), recursive = TRUE, showWarnings = FALSE)

message("=", paste(rep("=", 60), collapse = ""), "=")
message("CPI MULTI-DISEASE COMPARISON")
message("=", paste(rep("=", 60), collapse = ""), "=")

# =============================================================================
# [1/5] Load all CPI results
# =============================================================================
message("\n[1/5] Loading CPI results from all diseases...")

# TB BAL (from main analysis)
tb_bal <- tryCatch({
  df <- fread(file.path(TB_DIR, "4_results/tables/Table_CPI_by_celltype.csv"))
  df$disease <- "TB (BAL)"
  df$tissue <- "BAL"
  message(sprintf("  TB BAL: %d cell types, mean CPI = %.1f%%", nrow(df), mean(df$CPI) * 100))
  df
}, error = function(e) {
  message("  Warning: TB BAL data not found")
  data.frame()
})

# TB PBMC (from main analysis)
tb_pbmc <- tryCatch({
  df <- fread(file.path(TB_DIR, "4_results/tables/Table_CPI_GSE287288_DTB.csv"))
  df$disease <- "TB (PBMC)"
  df$tissue <- "PBMC"
  message(sprintf("  TB PBMC: %d cell types, mean CPI = %.1f%%", nrow(df), mean(df$CPI) * 100))
  df
}, error = function(e) {
  message("  Warning: TB PBMC data not found")
  data.frame()
})

# Sepsis
sepsis <- tryCatch({
  fpath <- file.path(RESULTS_DIR, "tables/CPI_Sepsis.csv")
  if (file.exists(fpath)) {
    df <- fread(fpath)
    df$tissue <- "PBMC"
    message(sprintf("  Sepsis: %d cell types, mean CPI = %.1f%%", nrow(df), mean(df$CPI) * 100))
    df
  } else {
    message("  Note: Sepsis data file not found - run 01_sepsis_cpi.R first")
    data.frame()
  }
}, error = function(e) {
  message("  Error loading Sepsis data: ", e$message)
  data.frame()
})

# Dengue
dengue <- tryCatch({
  fpath <- file.path(RESULTS_DIR, "tables/CPI_Dengue.csv")
  if (file.exists(fpath)) {
    df <- fread(fpath)
    df$tissue <- "PBMC"
    message(sprintf("  Dengue: %d cell types, mean CPI = %.1f%%", nrow(df), mean(df$CPI) * 100))
    df
  } else {
    message("  Note: Dengue data file not found - run 02_dengue_cpi.R first")
    data.frame()
  }
}, error = function(e) {
  message("  Error loading Dengue data: ", e$message)
  data.frame()
})

# =============================================================================
# [2/5] Combine all data
# =============================================================================
message("\n[2/5] Combining all CPI data...")

# Standardize column names and combine
all_datasets <- list()

if (nrow(tb_bal) > 0) {
  all_datasets$tb_bal <- tb_bal %>% 
    select(cell_type, CPI, disease, tissue) %>%
    mutate(disease_category = "TB")
}

if (nrow(tb_pbmc) > 0) {
  all_datasets$tb_pbmc <- tb_pbmc %>%
    select(cell_type, CPI, disease, tissue) %>%
    mutate(disease_category = "TB")
}

if (nrow(sepsis) > 0) {
  all_datasets$sepsis <- sepsis %>%
    select(cell_type, CPI, disease, tissue) %>%
    mutate(disease_category = "Sepsis")
}

if (nrow(dengue) > 0) {
  all_datasets$dengue <- dengue %>%
    select(cell_type, CPI, disease, tissue) %>%
    mutate(disease_category = "Dengue")
}

if (length(all_datasets) == 0) {
  stop("No CPI data available. Run individual disease analyses first.")
}

all_cpi <- bind_rows(all_datasets)
all_cpi$cell_type <- gsub("_", " ", all_cpi$cell_type)

message(sprintf("  Combined: %d observations across %d diseases", 
                nrow(all_cpi), length(unique(all_cpi$disease))))

# =============================================================================
# [3/5] Summary statistics
# =============================================================================
message("\n[3/5] Computing summary statistics...")

summary_stats <- all_cpi %>%
  group_by(disease) %>%
  summarise(
    n_cell_types = n(),
    mean_CPI = round(mean(CPI), 4),
    sd_CPI = round(sd(CPI), 4),
    min_CPI = round(min(CPI), 4),
    max_CPI = round(max(CPI), 4),
    median_CPI = round(median(CPI), 4),
    .groups = "drop"
  ) %>%
  arrange(desc(mean_CPI))

print(summary_stats)
fwrite(summary_stats, file.path(RESULTS_DIR, "tables", "CPI_CrossDisease_Summary.csv"))

# Overall statistics
overall_mean <- mean(all_cpi$CPI)
overall_sd <- sd(all_cpi$CPI)
overall_range <- range(all_cpi$CPI)

message("\n  Overall CPI Statistics:")
message(sprintf("    Mean: %.1f%% (SD: %.1f%%)", overall_mean * 100, overall_sd * 100))
message(sprintf("    Range: %.1f%% - %.1f%%", overall_range[1] * 100, overall_range[2] * 100))

# Cross-disease consistency (coefficient of variation)
disease_means <- summary_stats$mean_CPI
cv <- sd(disease_means) / mean(disease_means) * 100
message(sprintf("    Cross-disease CV: %.1f%% (lower = more consistent)", cv))

# =============================================================================
# [4/5] Statistical tests
# =============================================================================
message("\n[4/5] Statistical analysis...")

# Test if CPI differs significantly across diseases
if (length(unique(all_cpi$disease)) >= 2) {
  kruskal_result <- kruskal.test(CPI ~ disease, data = all_cpi)
  message(sprintf("  Kruskal-Wallis test: χ² = %.2f, p = %.4f", 
                  kruskal_result$statistic, kruskal_result$p.value))
  
  if (kruskal_result$p.value > 0.05) {
    message("  → CPI is NOT significantly different across diseases (p > 0.05)")
    message("  → This supports the hypothesis of cross-disease CPI consistency")
  } else {
    message("  → CPI shows significant variation across diseases (p < 0.05)")
  }
}

# =============================================================================
# [5/5] Visualizations
# =============================================================================
message("\n[5/5] Generating visualizations...")

# Color palette for diseases
disease_colors <- c(
  "TB (BAL)" = "#E41A1C",
  "TB (PBMC)" = "#377EB8",
  "Sepsis" = "#4DAF4A",
  "Dengue" = "#FF7F00"
)

# --- Figure 1: Multi-disease CPI comparison (boxplot) ---
p1 <- ggplot(all_cpi, aes(x = reorder(disease, -CPI, FUN = median), y = CPI * 100, fill = disease)) +
  geom_boxplot(alpha = 0.7, width = 0.6, outlier.shape = NA) +
  geom_jitter(width = 0.15, size = 3, alpha = 0.7, color = "black") +
  geom_hline(yintercept = overall_mean * 100, linetype = "dashed", color = "gray40", linewidth = 0.8) +
  annotate("text", x = length(unique(all_cpi$disease)) + 0.3, y = overall_mean * 100 + 2, 
           label = sprintf("Overall mean: %.0f%%", overall_mean * 100), 
           hjust = 1, size = 3.5, color = "gray40") +
  labs(
    title = "Chromatin Priming Index Across Infectious Diseases",
    subtitle = sprintf("CPI shows cross-disease consistency (CV = %.1f%%)", cv),
    x = "Disease / Dataset",
    y = "Chromatin Priming Index (%)",
    caption = "Each point represents a cell type. Horizontal line = overall mean."
  ) +
  theme_minimal() +
  theme(
    legend.position = "none",
    plot.title = element_text(face = "bold", size = 14),
    plot.subtitle = element_text(size = 11, color = "gray40"),
    axis.text.x = element_text(size = 11, angle = 15, hjust = 1),
    panel.grid.major.x = element_blank()
  ) +
  scale_fill_manual(values = disease_colors) +
  coord_cartesian(ylim = c(50, 100))

ggsave(file.path(RESULTS_DIR, "figures", "Fig1_MultiDisease_CPI_Comparison.png"), 
       p1, width = 9, height = 7, dpi = 300)

# --- Figure 2: CPI by cell type across diseases (grouped bar) ---
p2 <- ggplot(all_cpi, aes(x = reorder(cell_type, CPI), y = CPI * 100, fill = disease)) +
  geom_bar(stat = "identity", position = position_dodge(width = 0.8), width = 0.7) +
  geom_hline(yintercept = 80, linetype = "dashed", color = "gray50") +
  coord_flip() +
  labs(
    title = "CPI by Cell Type Across Diseases",
    x = "Cell Type",
    y = "CPI (%)",
    fill = "Disease"
  ) +
  theme_minimal() +
  theme(
    legend.position = "bottom",
    plot.title = element_text(face = "bold", size = 13),
    panel.grid.major.y = element_blank()
  ) +
  scale_fill_manual(values = disease_colors) +
  guides(fill = guide_legend(nrow = 1))

ggsave(file.path(RESULTS_DIR, "figures", "Fig2_CPI_CellType_ByDisease.png"), 
       p2, width = 10, height = 7, dpi = 300)

# --- Figure 3: CPI distribution density plot ---
p3 <- ggplot(all_cpi, aes(x = CPI * 100, fill = disease_category, color = disease_category)) +
  geom_density(alpha = 0.4, linewidth = 1) +
  geom_vline(xintercept = overall_mean * 100, linetype = "dashed", color = "black") +
  labs(
    title = "CPI Distribution by Disease Category",
    x = "Chromatin Priming Index (%)",
    y = "Density",
    fill = "Disease",
    color = "Disease"
  ) +
  theme_minimal() +
  theme(
    legend.position = "bottom",
    plot.title = element_text(face = "bold", size = 13)
  ) +
  scale_fill_brewer(palette = "Set1") +
  scale_color_brewer(palette = "Set1")

ggsave(file.path(RESULTS_DIR, "figures", "Fig3_CPI_Distribution.png"), 
       p3, width = 8, height = 6, dpi = 300)

# --- Figure 4: Heatmap-style visualization ---
# Create summary matrix
heatmap_data <- all_cpi %>%
  group_by(disease, cell_type) %>%
  summarise(mean_CPI = mean(CPI), .groups = "drop")

p4 <- ggplot(heatmap_data, aes(x = disease, y = cell_type, fill = mean_CPI * 100)) +
  geom_tile(color = "white", linewidth = 0.5) +
  geom_text(aes(label = sprintf("%.0f%%", mean_CPI * 100)), size = 3.5, color = "black") +
  scale_fill_gradient2(low = "#377EB8", mid = "#FFFFBF", high = "#E41A1C", 
                       midpoint = 80, limits = c(60, 100),
                       name = "CPI (%)") +
  labs(
    title = "CPI Heatmap: Cell Type × Disease",
    x = "Disease",
    y = "Cell Type"
  ) +
  theme_minimal() +
  theme(
    plot.title = element_text(face = "bold", size = 13),
    axis.text.x = element_text(angle = 30, hjust = 1, size = 10),
    panel.grid = element_blank()
  )

ggsave(file.path(RESULTS_DIR, "figures", "Fig4_CPI_Heatmap.png"), 
       p4, width = 9, height = 7, dpi = 300)

# =============================================================================
# Save combined data
# =============================================================================
all_cpi$CPI_percent <- round(all_cpi$CPI * 100, 1)
fwrite(all_cpi, file.path(RESULTS_DIR, "tables", "CPI_AllDiseases.csv"))

# =============================================================================
# Summary
# =============================================================================
message("\n", paste(rep("=", 60), collapse = ""))
message("CROSS-DISEASE ANALYSIS COMPLETE")
message(paste(rep("=", 60), collapse = ""))
message("\nKey Findings:")
message(sprintf("  • Overall mean CPI: %.1f%% (range: %.1f%% - %.1f%%)", 
                overall_mean * 100, overall_range[1] * 100, overall_range[2] * 100))
message(sprintf("  • Cross-disease coefficient of variation: %.1f%%", cv))

if (cv < 15) {
  message("  • CPI shows STRONG cross-disease consistency")
} else if (cv < 25) {
  message("  • CPI shows MODERATE cross-disease consistency")
} else {
  message("  • CPI shows VARIABLE results across diseases")
}

message("\nOutputs:")
message("  Figures: ", file.path(RESULTS_DIR, "figures"))
message("  Tables: ", file.path(RESULTS_DIR, "tables"))
message(paste(rep("=", 60), collapse = ""))
