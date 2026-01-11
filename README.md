# Epigenetic Locking of Vascular Shock: Multi-Disease Extension

[![DOI](https://zenodo.org/badge/DOI/10.5281/zenodo.123456.svg)](https://doi.org/10.5281/zenodo.123456)
[![Reproducibility Check](https://github.com/hssling/CPI_MultiDisease_Extension/actions/workflows/reproducibility.yml/badge.svg)](https://github.com/hssling/CPI_MultiDisease_Extension/actions/workflows/reproducibility.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

## Overview
This repository contains the analysis code and results for the study: **"Epigenetic Locking of Vascular and Inflammatory Effectors Defines the Universal Host Response to Severe Infection"**.

We demonstrate that a conserved "epigenetic alert state" exists across **Tuberculosis**, **Sepsis**, and **Dengue**, driven by the chromatin priming of **VEGFA** in circulating immune cells.

## Repository Structure
```
CPI_MultiDisease_Extension/
├── 2_analysis/                 # R Analysis Scripts
│   ├── 01_sepsis_cpi.R         # Sepsis (GSE151263) Pipeline
│   ├── 02_dengue_cpi.R         # Dengue (GSE154386) Pipeline
│   ├── 03_cross_disease.R      # Integration & Stats
│   └── 04_core_signature.R     # Core 616 Gene Identification
├── 3_results/                  # Output Figures & Tables
│   ├── core_signature/         # Heatmaps & Gene Lists
│   └── figures/                # High-Res Figures for Paper
├── Submission_Package/         # Generated Manuscripts (DOCX)
├── generate_manuscripts.py     # Python script to build DOCX
└── environment.yml             # Conda Environment
```

## Reproducibility

### 1. Setup Environment
To replicate the computational environment:
```bash
conda env create -f environment.yml
conda activate cpi_env
```

### 2. Run Analysis
The scripts in `2_analysis/` are numbered sequentially. Note that raw data (GSE files) must be downloaded from GEO and placed in `1_data_raw/` (ignored in repo to save space).

### 3. Generate Manuscripts
The final manuscripts with embedded figures can be regenerated programmatically:
```bash
python generate_manuscripts_docx.py
```
This ensures that the text and figures are always perfectly synced with the latest results.

## Key Findings
- **Universal Priming:** 80-84% of immune response genes are epigenetically primed across pathologies.
- **VEGFA Mechanism:** We identify VEGFA upregulation (+4.0 LFC in Dengue) as an epigenetically locked trait in PBMCs, explaining the shared vascular shock phenotype.

## Author & Citation
**Dr. Siddalingaiah H S, MD**  
Shridevi Institute of Medical Sciences and Research Hospital  
Tumkur, India.

Please cite: *Siddalingaiah H S. Epigenetic Locking of Vascular and Inflammatory Effectors. 2026.*
