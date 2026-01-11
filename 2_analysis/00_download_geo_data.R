# =============================================================================
# CPI Multi-Disease Extension: Real GEO Data Download
# Downloads actual scRNA-seq datasets from GEO FTP
# GSE151263 - Sepsis PBMC, GSE154386 - Dengue PBMC
# =============================================================================

library(GEOquery)
library(data.table)
library(R.utils)

# Configuration
BASE_DIR <- "d:/research-automation/TB multiomics/TB Chromatin Priming Multiomics/CPI_MultiDisease_Extension"
setwd(BASE_DIR)

DATA_DIR <- file.path(BASE_DIR, "1_data_raw")

message("=", paste(rep("=", 60), collapse = ""), "=")
message("DOWNLOADING REAL GEO DATASETS")
message("=", paste(rep("=", 60), collapse = ""), "=")

# =============================================================================
# GSE151263 - Sepsis PBMC scRNA-seq
# =============================================================================
message("\n[1/2] Downloading GSE151263 (Sepsis PBMC)...")

sepsis_dir <- file.path(DATA_DIR, "sepsis")
dir.create(sepsis_dir, recursive = TRUE, showWarnings = FALSE)

# Download supplementary files from GEO FTP
sepsis_ftp <- "https://ftp.ncbi.nlm.nih.gov/geo/series/GSE151nnn/GSE151263/suppl/"

# Get GEO metadata
message("  Fetching GEO metadata...")
gse_sepsis <- tryCatch({
  getGEO("GSE151263", GSEMatrix = TRUE, destdir = sepsis_dir, getGPL = FALSE)
}, error = function(e) {
  message("  Warning: Could not fetch full metadata: ", e$message)
  NULL
})

if (!is.null(gse_sepsis)) {
  # Extract sample info
  if (is.list(gse_sepsis) && length(gse_sepsis) > 0) {
    pdata <- pData(gse_sepsis[[1]])
    message("  Found ", nrow(pdata), " samples")
    
    # Save sample metadata
    fwrite(as.data.frame(pdata), file.path(sepsis_dir, "sample_metadata.csv"))
    message("  Saved sample metadata")
    
    # Print sample characteristics
    if ("characteristics_ch1" %in% colnames(pdata)) {
      message("  Sample characteristics:")
      print(table(pdata$characteristics_ch1))
    }
  }
}

# Download supplementary data files
message("  Downloading supplementary count matrices...")
suppl_file <- file.path(sepsis_dir, "GSE151263_RAW.tar")

if (!file.exists(suppl_file)) {
  tryCatch({
    # GEOquery's getGEOSuppFiles downloads all supplementary files
    suppl_files <- getGEOSuppFiles("GSE151263", makeDirectory = FALSE, baseDir = sepsis_dir)
    message("  Downloaded: ", paste(rownames(suppl_files), collapse = ", "))
  }, error = function(e) {
    message("  Error downloading: ", e$message)
    message("  Trying direct FTP download...")
    
    # Alternative: direct download
    download_url <- "https://www.ncbi.nlm.nih.gov/geo/download/?acc=GSE151263&format=file"
    download.file(download_url, suppl_file, mode = "wb", timeout = 600)
  })
}

# Extract and process
suppl_files <- list.files(sepsis_dir, pattern = "\\.tar$|\\.gz$", full.names = TRUE)
message("  Found ", length(suppl_files), " supplementary files")

# =============================================================================
# GSE154386 - Dengue PBMC scRNA-seq (171K cells)
# =============================================================================
message("\n[2/2] Downloading GSE154386 (Dengue PBMC)...")

dengue_dir <- file.path(DATA_DIR, "dengue")
dir.create(dengue_dir, recursive = TRUE, showWarnings = FALSE)

# Get GEO metadata
message("  Fetching GEO metadata...")
gse_dengue <- tryCatch({
  getGEO("GSE154386", GSEMatrix = TRUE, destdir = dengue_dir, getGPL = FALSE)
}, error = function(e) {
  message("  Warning: Could not fetch full metadata: ", e$message)
  NULL
})

if (!is.null(gse_dengue)) {
  if (is.list(gse_dengue) && length(gse_dengue) > 0) {
    pdata <- pData(gse_dengue[[1]])
    message("  Found ", nrow(pdata), " samples")
    
    # Save sample metadata
    fwrite(as.data.frame(pdata), file.path(dengue_dir, "sample_metadata.csv"))
    message("  Saved sample metadata")
  }
}

# Download supplementary files
message("  Downloading supplementary count matrices (large dataset ~171K cells)...")
dengue_suppl <- file.path(dengue_dir, "GSE154386_RAW.tar")

if (!file.exists(dengue_suppl)) {
  tryCatch({
    suppl_files <- getGEOSuppFiles("GSE154386", makeDirectory = FALSE, baseDir = dengue_dir)
    message("  Downloaded: ", paste(rownames(suppl_files), collapse = ", "))
  }, error = function(e) {
    message("  Error downloading: ", e$message)
    
    # Direct download
    download_url <- "https://www.ncbi.nlm.nih.gov/geo/download/?acc=GSE154386&format=file"
    download.file(download_url, dengue_suppl, mode = "wb", timeout = 1200)
  })
}

# =============================================================================
# Summary
# =============================================================================
message("\n", paste(rep("=", 60), collapse = ""))
message("DOWNLOAD COMPLETE")
message(paste(rep("=", 60), collapse = ""))

# List downloaded files
message("\nSepsis files:")
sepsis_files <- list.files(sepsis_dir, full.names = FALSE)
for (f in sepsis_files) {
  fsize <- file.info(file.path(sepsis_dir, f))$size
  message(sprintf("  %s (%.1f MB)", f, fsize / 1024^2))
}

message("\nDengue files:")
dengue_files <- list.files(dengue_dir, full.names = FALSE)
for (f in dengue_files) {
  fsize <- file.info(file.path(dengue_dir, f))$size
  message(sprintf("  %s (%.1f MB)", f, fsize / 1024^2))
}

message("\nNext: Run 01_sepsis_cpi.R and 02_dengue_cpi.R to process the data")
