# =============================================================================
# CPI Multi-Disease Extension: Dengue Analysis (Real Data)
# Dataset: GSE154386 - Dengue PBMC scRNA-seq
# Analyzes DEGs between Dengue and Pre-infection conditions
# =============================================================================

library(Seurat)
library(dplyr)
library(ggplot2)
library(data.table)
library(Matrix)

# Configuration
BASE_DIR <- "d:/research-automation/TB multiomics/TB Chromatin Priming Multiomics/CPI_MultiDisease_Extension"
setwd(BASE_DIR)

DATA_DIR <- file.path(BASE_DIR, "1_data_raw", "dengue")
RESULTS_DIR <- file.path(BASE_DIR, "3_results")
CONFIG_DIR <- file.path(BASE_DIR, "config")

# Create directories
dir.create(file.path(RESULTS_DIR, "figures"), recursive = TRUE, showWarnings = FALSE)
dir.create(file.path(RESULTS_DIR, "tables"), recursive = TRUE, showWarnings = FALSE)

message("=", paste(rep("=", 60), collapse = ""), "=")
message("CPI MULTI-DISEASE EXTENSION: DENGUE ANALYSIS (REAL DATA)")
message("=", paste(rep("=", 60), collapse = ""), "=")

# =============================================================================
# [1/6] Load ATAC reference
# =============================================================================
message("\n[1/6] Loading ATAC reference...")
peak_links <- fread(file.path(CONFIG_DIR, "peak_gene_links.csv"))
linked_genes <- unique(peak_links$gene)
message(sprintf("  Loaded %d peak-gene links for %d genes", nrow(peak_links), length(linked_genes)))

# =============================================================================
# [2/6] Load real GSE154386 data
# =============================================================================
message("\n[2/6] Loading GSE154386 real count matrices...")

rds_file <- file.path(DATA_DIR, "dengue_seurat_real.rds")
tar_file <- file.path(DATA_DIR, "GSE154386_RAW.tar")

if (!file.exists(rds_file)) {
    # Check/Extract TAR
    geo_files <- list.files(DATA_DIR, pattern = "matrix.mtx.gz$", full.names = TRUE)
    if (length(geo_files) == 0) {
        if (!file.exists(tar_file)) stop("GSE154386_RAW.tar not found.")
        message("  Extracting RAW tar file...")
        untar(tar_file, exdir = DATA_DIR)
        geo_files <- list.files(DATA_DIR, pattern = "matrix.mtx.gz$", full.names = TRUE)
    }
    
    if (length(geo_files) == 0) stop("No matrix files found.")

    # Load samples
    seurat_list <- list()
    target_cells <- 20000 
    
    # Identify unique samples from matrix files
    # Format: GMxxxx_SampleName_matrix.mtx.gz
    sample_prefixes <- gsub("_matrix.mtx.gz", "", basename(geo_files))
    
    for (prefix in sample_prefixes) {
        # Determine condition
        condition <- "Unknown"
        # Logic: D0, D2, D4, D6 = Acute; Dneg = Pre-infection
        if (grepl("_D0|_D2|_D4|_D6", prefix, ignore.case=TRUE)) condition <- "Dengue_Acute"
        if (grepl("_Dneg", prefix, ignore.case=TRUE)) condition <- "Pre_infection"
        
        if (condition == "Unknown") {
            # Check for "Healthy"
            if (grepl("Healthy", prefix, ignore.case=TRUE)) condition <- "Pre_infection" 
            else {
                message(sprintf("  Skipping %s (Unknown condition)", prefix))
                next
            }
        }

        # Optimization: Skip if we already have enough of this condition
        current_n_acute <- sum(sapply(seurat_list, function(x) x$condition == "Dengue_Acute"))
        current_n_pre <- sum(sapply(seurat_list, function(x) x$condition == "Pre_infection"))
        
        if (condition == "Dengue_Acute" && current_n_acute >= 4) {
             message(sprintf("  Skipping %s (Already have %d Acute samples)", prefix, current_n_acute))
             next
        }
        
        message(sprintf("  Loading %s (%s)...", prefix, condition))
        
        # Paths
        mat_path <- file.path(DATA_DIR, paste0(prefix, "_matrix.mtx.gz"))
        genes_path <- file.path(DATA_DIR, paste0(prefix, "_genes.tsv.gz"))
        if (!file.exists(genes_path)) genes_path <- file.path(DATA_DIR, paste0(prefix, "_features.tsv.gz"))
        barcodes_path <- file.path(DATA_DIR, paste0(prefix, "_barcodes.tsv.gz"))
        
        if (!file.exists(mat_path) || !file.exists(genes_path) || !file.exists(barcodes_path)) {
            message("    Missing triplet files, skipping.")
            next
        }
        
        # Read 10x
        counts <- ReadMtx(mtx = mat_path, features = genes_path, cells = barcodes_path)
        
        # Create object
        obj <- CreateSeuratObject(counts = counts, project = prefix, min.cells = 3, min.features = 200)
        obj$condition <- condition
        obj$disease <- "Dengue_Study"
        
        # Subsample immediately 
        if (ncol(obj) > 2000) {
            obj <- subset(obj, cells = sample(Cells(obj), 2000))
        }
        
        seurat_list[[prefix]] <- obj
        
        # Check if we have enough
        n_acute <- sum(sapply(seurat_list, function(x) x$condition == "Dengue_Acute"))
        n_pre <- sum(sapply(seurat_list, function(x) x$condition == "Pre_infection"))
        
        if (n_acute >= 3 && n_pre >= 3 && length(seurat_list) >= 8) {
             message("  Loaded sufficient samples.")
             break
        }
    }
    
    if (length(seurat_list) == 0) stop("No valid samples loaded.")
    
    message("  Merging samples...")
    pbmc <- merge(seurat_list[[1]], y = seurat_list[-1], add.cell.ids = names(seurat_list))
    
    # Subsample to target size
    if (ncol(pbmc) > target_cells) {
        message(sprintf("  Subsampling from %d to %d cells...", ncol(pbmc), target_cells))
        pbmc <- subset(pbmc, cells = sample(Cells(pbmc), target_cells))
    }
    
    # Process
    if (exists("JoinLayers", where = "package:SeuratObject") || exists("JoinLayers", where = "package:Seurat")) {
        pbmc <- JoinLayers(pbmc)
    }
    
    pbmc <- NormalizeData(pbmc, verbose = FALSE)
    pbmc <- FindVariableFeatures(pbmc, verbose = FALSE)
    pbmc <- ScaleData(pbmc, verbose = FALSE)
    pbmc <- RunPCA(pbmc, verbose = FALSE)
    pbmc <- FindNeighbors(pbmc, dims = 1:15, verbose = FALSE)
    pbmc <- FindClusters(pbmc, resolution = 0.5, verbose = FALSE)
    pbmc <- RunUMAP(pbmc, dims = 1:15, verbose = FALSE)
    
    saveRDS(pbmc, rds_file)
    
} else {
    pbmc <- readRDS(rds_file)
    if (exists("JoinLayers", where = "package:SeuratObject") || exists("JoinLayers", where = "package:Seurat")) {
        pbmc <- JoinLayers(pbmc)
    }
}

message(sprintf("  Total cells: %d", ncol(pbmc)))
print(table(pbmc$condition))

# =============================================================================
# [3/6] Cell type annotation
# =============================================================================
message("\n[3/6] Annotating cell types...")
# (Using same markers as Sepsis)
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

pbmc$cell_type <- "Unknown"
for (ct in names(markers)) {
  score_col <- paste0(ct, "_score1")
  if (score_col %in% colnames(pbmc@meta.data)) {
    high_score <- pbmc@meta.data[[score_col]] > 0.5
    pbmc$cell_type[high_score & pbmc$cell_type == "Unknown"] <- ct
  }
}

# =============================================================================
# [4/6] DEG & [5/6] CPI
# =============================================================================
message("\n[4/6] Calculating DEGs & CPI...")
cpi_results <- data.frame()
deg_results <- list()
cell_types <- unique(pbmc$cell_type)
cell_types <- cell_types[cell_types != "Unknown"]

for (ct in cell_types) {
    message(sprintf("  Processing %s...", ct))
    cells_ct <- WhichCells(pbmc, expression = cell_type == ct)
    if (length(cells_ct) < 50) next
    
    pbmc_ct <- subset(pbmc, cells = cells_ct)
    if (exists("JoinLayers", where = "package:SeuratObject") || exists("JoinLayers", where = "package:Seurat")) {
        pbmc_ct <- JoinLayers(pbmc_ct)
    }
    
    Idents(pbmc_ct) <- "condition"
    if (sum(pbmc_ct$condition == "Dengue_Acute") < 10 || sum(pbmc_ct$condition == "Pre_infection") < 10) next
    
    tryCatch({
        degs <- FindMarkers(pbmc_ct, ident.1 = "Dengue_Acute", ident.2 = "Pre_infection", 
                            min.pct = 0.1, logfc.threshold = 0.1)
        degs$gene <- rownames(degs)
        degs$cell_type <- ct
        deg_results[[ct]] <- degs
        
        # Calculate CPI
        sig_degs <- degs %>% filter(p_val_adj < 0.05)
        n_deg <- nrow(sig_degs)
        if (n_deg > 0) {
            primed <- sig_degs$gene %in% linked_genes
            cpi <- sum(primed) / n_deg
            cpi_results <- rbind(cpi_results, data.frame(
                cell_type = ct, n_deg = n_deg, n_deg_with_link = sum(primed),
                CPI = round(cpi, 4), disease = "Dengue", dataset = "GSE154386"
            ))
            message(sprintf("    %s CPI: %.1f%%", ct, cpi*100))
        }
    }, error = function(e) message(e$message))
}

if (nrow(cpi_results) > 0) fwrite(cpi_results, file.path(RESULTS_DIR, "tables", "CPI_Dengue.csv"))
if (length(deg_results) > 0) fwrite(do.call(rbind, deg_results), file.path(RESULTS_DIR, "tables", "DEG_Dengue.csv"))

# [6/6] Viz
p1 <- DimPlot(pbmc, group.by = "condition", cols = c("Dengue_Acute"="#FF7F00", "Pre_infection"="#4DAF4A"))
ggsave(file.path(RESULTS_DIR, "figures", "Fig_Dengue_UMAP_condition.png"), p1)
p2 <- DimPlot(pbmc, group.by = "cell_type", label=TRUE)
ggsave(file.path(RESULTS_DIR, "figures", "Fig_Dengue_UMAP_celltype.png"), p2)

message("DENGUE ANALYSIS COMPLETE")
