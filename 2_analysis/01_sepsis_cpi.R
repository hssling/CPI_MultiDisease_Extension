# =============================================================================
# CPI Multi-Disease Extension: Sepsis Analysis (Real Data)
# Dataset: GSE151263 - Sepsis/ARDS PBMC scRNA-seq
# Analyzes DEGs between Sepsis/ARDS conditions
# =============================================================================

library(Seurat)
library(dplyr)
library(ggplot2)
library(data.table)

# Configuration
BASE_DIR <- "d:/research-automation/TB multiomics/TB Chromatin Priming Multiomics/CPI_MultiDisease_Extension"
setwd(BASE_DIR)

DATA_DIR <- file.path(BASE_DIR, "1_data_raw", "sepsis")
RESULTS_DIR <- file.path(BASE_DIR, "3_results")
CONFIG_DIR <- file.path(BASE_DIR, "config")

# Create directories
dir.create(file.path(RESULTS_DIR, "figures"), recursive = TRUE, showWarnings = FALSE)
dir.create(file.path(RESULTS_DIR, "tables"), recursive = TRUE, showWarnings = FALSE)

message("=", paste(rep("=", 60), collapse = ""), "=")
message("CPI MULTI-DISEASE EXTENSION: SEPSIS ANALYSIS (REAL DATA)")
message("=", paste(rep("=", 60), collapse = ""), "=")

# =============================================================================
# [1/6] Load ATAC reference (peak-gene links)
# =============================================================================
message("\n[1/6] Loading ATAC reference...")
peak_links <- fread(file.path(CONFIG_DIR, "peak_gene_links.csv"))
linked_genes <- unique(peak_links$gene)
message(sprintf("  Loaded %d peak-gene links for %d genes", nrow(peak_links), length(linked_genes)))

# =============================================================================
# [2/6] Load real GSE151263 data
# =============================================================================
message("\n[2/6] Loading GSE151263 real count matrices...")

rds_file <- file.path(DATA_DIR, "sepsis_seurat_real.rds")

if (!file.exists(rds_file)) {
  # List UMI count files
  umi_files <- list.files(DATA_DIR, pattern = "_processed_UMI.txt.gz$", full.names = TRUE)
  message(sprintf("  Found %d UMI count files", length(umi_files)))
  
  if (length(umi_files) == 0) {
    stop("No UMI count files found. Run 00_download_geo_data.R first.")
  }
  
  # Load and merge all samples
  seurat_list <- list()
  
  for (f in umi_files) {
    sample_name <- gsub(".*_(ARDS\\d+|Sepsis\\d+)_.*", "\\1", basename(f))
    message(sprintf("  Loading %s...", sample_name))
    
    # Read count matrix (rows = genes, cols = cells)
    counts <- fread(f, header = TRUE)
    
    # First column is usually gene names
    gene_col <- colnames(counts)[1]
    genes <- counts[[gene_col]]
    counts[[gene_col]] <- NULL
    
    # Make gene names unique (handle duplicates)
    genes <- make.unique(genes)
    
    # Convert to matrix
    count_matrix <- as.matrix(counts)
    rownames(count_matrix) <- genes
    
    # Create Seurat object
    obj <- CreateSeuratObject(
      counts = count_matrix,
      project = sample_name,
      min.cells = 3,
      min.features = 200
    )
    
    # Add sample metadata
    obj$sample <- sample_name
    obj$condition <- ifelse(grepl("ARDS", sample_name), "ARDS", "Sepsis")
    obj$disease <- "Sepsis_Study"
    
    seurat_list[[sample_name]] <- obj
    message(sprintf("    %d cells, %d genes", ncol(obj), nrow(obj)))
  }
  
  # Merge all samples
  message("\n  Merging all samples...")
  pbmc <- merge(seurat_list[[1]], y = seurat_list[-1], 
                add.cell.ids = names(seurat_list))
  
  message(sprintf("  Total: %d cells across %d samples", ncol(pbmc), length(seurat_list)))
  
  # Standard processing
  message("  Processing merged data...")
  pbmc <- NormalizeData(pbmc, verbose = FALSE)
  pbmc <- FindVariableFeatures(pbmc, nfeatures = 3000, verbose = FALSE)
  pbmc <- ScaleData(pbmc, verbose = FALSE)
  pbmc <- RunPCA(pbmc, npcs = 30, verbose = FALSE)
  pbmc <- FindNeighbors(pbmc, dims = 1:20, verbose = FALSE)
  pbmc <- FindClusters(pbmc, resolution = 0.8, verbose = FALSE)
  pbmc <- RunUMAP(pbmc, dims = 1:20, verbose = FALSE)
  
  saveRDS(pbmc, rds_file)
  message("  Saved processed Seurat object")
} else {
  pbmc <- readRDS(rds_file)
  message(sprintf("  Loaded %d cells from cache", ncol(pbmc)))
}

# Join layers for Seurat v5 compatibility (required for merged objects)
if (exists("JoinLayers", where = "package:SeuratObject") || 
    exists("JoinLayers", where = "package:Seurat")) {
  message("  Joining data layers for Seurat v5...")
  pbmc <- JoinLayers(pbmc)
}

message(sprintf("\n  Condition distribution:"))
print(table(pbmc$condition))

# =============================================================================
# [3/6] Cell type annotation
# =============================================================================
message("\n[3/6] Annotating cell types...")

markers <- list(
  T_cell = c("CD3D", "CD3E", "CD4", "CD8A"),
  NK_cell = c("GNLY", "NKG7", "NCAM1"),
  B_cell = c("CD79A", "MS4A1", "CD19"),
  Monocyte = c("CD14", "LYZ", "S100A8", "S100A9"),
  DC = c("FCER1A", "CD1C"),
  Platelet = c("PPBP", "PF4")
)

for (ct in names(markers)) {
  genes_present <- intersect(markers[[ct]], rownames(pbmc))
  if (length(genes_present) >= 2) {
    pbmc <- tryCatch({
      AddModuleScore(pbmc, features = list(genes_present), name = paste0(ct, "_score"))
    }, error = function(e) pbmc)
  }
}

# Assign cell types
pbmc$cell_type <- "Unknown"
for (ct in names(markers)) {
  score_col <- paste0(ct, "_score1")
  if (score_col %in% colnames(pbmc@meta.data)) {
    high_score <- pbmc@meta.data[[score_col]] > 0.5
    pbmc$cell_type[high_score & pbmc$cell_type == "Unknown"] <- ct
  }
}

cell_counts <- table(pbmc$cell_type)
message("  Cell type distribution:")
print(cell_counts)

# =============================================================================
# [4/6] DEG analysis: Sepsis vs ARDS (or by sample if needed)
# =============================================================================
message("\n[4/6] Calculating DEGs...")

# Join layers for Seurat v5 compatibility
if ("JoinLayers" %in% ls("package:Seurat")) {
  pbmc <- JoinLayers(pbmc)
}

deg_results <- list()
cell_types <- names(cell_counts)[cell_counts >= 20]
cell_types <- cell_types[cell_types != "Unknown"]

# DEG by cell type - comparing disease signatures across all samples
for (ct in cell_types) {
  message(sprintf("  Processing %s...", ct))
  
  # Subset to this cell type
  cells_ct <- WhichCells(pbmc, expression = cell_type == ct)
  
  if (length(cells_ct) < 50) {
    message(sprintf("    Skipping: insufficient cells (%d)", length(cells_ct)))
    next
  }
  
  pbmc_ct <- subset(pbmc, cells = cells_ct)
  
  # Join layers for Seurat v5 compatibility (must be done after subsetting)
  if ("JoinLayers" %in% ls("package:Seurat")) {
    pbmc_ct <- JoinLayers(pbmc_ct)
  }
  
  # Find markers for this cell type (characterizing expression patterns)
  tryCatch({
    Idents(pbmc_ct) <- "condition"
    n_sepsis <- sum(pbmc_ct$condition == "Sepsis")
    n_ards <- sum(pbmc_ct$condition == "ARDS")
    
    if (n_sepsis >= 10 && n_ards >= 10) {
      degs <- FindMarkers(pbmc_ct, ident.1 = "Sepsis", ident.2 = "ARDS",
                         min.pct = 0.1, logfc.threshold = 0.1)
      comparison <- "Sepsis_vs_ARDS"
    } else {
      # If not enough cells per condition, find all markers
      Idents(pbmc_ct) <- "seurat_clusters"
      degs <- FindAllMarkers(pbmc_ct, only.pos = TRUE, min.pct = 0.1, 
                             logfc.threshold = 0.25, max.cells.per.ident = 500)
      comparison <- "Cluster_markers"
    }
    
    if (nrow(degs) > 0) {
      degs$cell_type <- ct
      degs$comparison <- comparison
      if (!"gene" %in% colnames(degs)) degs$gene <- rownames(degs)
      deg_results[[ct]] <- degs
      message(sprintf("    Found %d DEGs (%s)", nrow(degs), comparison))
    }
  }, error = function(e) {
    message(sprintf("    Error: %s", e$message))
  })
}

if (length(deg_results) > 0) {
  all_degs <- do.call(rbind, deg_results)
  fwrite(all_degs, file.path(RESULTS_DIR, "tables", "DEG_Sepsis.csv"))
  message(sprintf("\n  Total: %d DEGs across %d cell types", nrow(all_degs), length(deg_results)))
} else {
  message("  Warning: No DEGs found")
  all_degs <- data.frame()
}

# =============================================================================
# [5/6] CPI calculation per cell type
# =============================================================================
message("\n[5/6] Calculating Chromatin Priming Index...")

cpi_results <- data.frame()

for (ct in names(deg_results)) {
  degs <- deg_results[[ct]]
  
  # Get significant DEGs
  if ("p_val_adj" %in% colnames(degs)) {
    sig_degs <- degs %>% filter(p_val_adj < 0.05)
  } else {
    sig_degs <- degs  # Use all if no p-value
  }
  
  n_deg <- nrow(sig_degs)
  
  if (n_deg > 0) {
    gene_col <- if ("gene" %in% colnames(sig_degs)) "gene" else rownames(sig_degs)
    deg_genes <- if (is.character(gene_col)) sig_degs[[gene_col]] else gene_col
    
    primed <- deg_genes %in% linked_genes
    n_primed <- sum(primed)
    cpi <- n_primed / n_deg
    
    cpi_results <- rbind(cpi_results, data.frame(
      cell_type = ct,
      n_deg = n_deg,
      n_deg_with_link = n_primed,
      CPI = round(cpi, 4),
      disease = "Sepsis",
      dataset = "GSE151263"
    ))
    
    message(sprintf("  %s: CPI = %.1f%% (%d/%d)", ct, cpi * 100, n_primed, n_deg))
  }
}

if (nrow(cpi_results) > 0) {
  fwrite(cpi_results, file.path(RESULTS_DIR, "tables", "CPI_Sepsis.csv"))
  message(sprintf("\n  Saved CPI results for %d cell types", nrow(cpi_results)))
} else {
  message("  Warning: No CPI results generated")
}

# =============================================================================
# [6/6] Visualizations
# =============================================================================
message("\n[6/6] Generating visualizations...")

# UMAP by condition
p1 <- DimPlot(pbmc, group.by = "condition", cols = c("Sepsis" = "#E41A1C", "ARDS" = "#377EB8")) +
  ggtitle("GSE151263: Sepsis vs ARDS") +
  theme_minimal()
ggsave(file.path(RESULTS_DIR, "figures", "Fig_Sepsis_UMAP_condition.png"), p1, width = 8, height = 6, dpi = 150)

# UMAP by sample
p2 <- DimPlot(pbmc, group.by = "sample") +
  ggtitle("GSE151263: By Sample") +
  theme_minimal()
ggsave(file.path(RESULTS_DIR, "figures", "Fig_Sepsis_UMAP_sample.png"), p2, width = 9, height = 6, dpi = 150)

# UMAP by cell type
p3 <- DimPlot(pbmc, group.by = "cell_type", label = TRUE, repel = TRUE) +
  ggtitle("GSE151263: Cell Types") +
  theme_minimal()
ggsave(file.path(RESULTS_DIR, "figures", "Fig_Sepsis_UMAP_celltype.png"), p3, width = 8, height = 6, dpi = 150)

# CPI bar plot
if (nrow(cpi_results) > 0) {
  p4 <- ggplot(cpi_results, aes(x = reorder(cell_type, CPI), y = CPI * 100, fill = cell_type)) +
    geom_bar(stat = "identity") +
    geom_hline(yintercept = 80, linetype = "dashed", color = "gray50") +
    coord_flip() +
    labs(title = "Chromatin Priming Index - Sepsis (GSE151263)",
         subtitle = "Real PBMC scRNA-seq data",
         x = "Cell Type", y = "CPI (%)") +
    theme_minimal() +
    theme(legend.position = "none") +
    scale_fill_brewer(palette = "Set2") +
    ylim(0, 100)
  ggsave(file.path(RESULTS_DIR, "figures", "Fig_Sepsis_CPI.png"), p4, width = 7, height = 5, dpi = 150)
}

# =============================================================================
# Summary
# =============================================================================
message("\n", paste(rep("=", 60), collapse = ""))
message("SEPSIS CPI ANALYSIS COMPLETE (REAL DATA)")
message(sprintf("  Dataset: GSE151263"))
message(sprintf("  Total cells: %d", ncol(pbmc)))
message(sprintf("  Samples: %d", length(unique(pbmc$sample))))
if (nrow(cpi_results) > 0) {
  message(sprintf("  Mean CPI: %.1f%%", mean(cpi_results$CPI) * 100))
  message(sprintf("  Range: %.1f%% - %.1f%%", min(cpi_results$CPI) * 100, max(cpi_results$CPI) * 100))
}
message("Results saved to: ", RESULTS_DIR)
message(paste(rep("=", 60), collapse = ""))
