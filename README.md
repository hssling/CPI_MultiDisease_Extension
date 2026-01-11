# Epigenetic Locking of Vascular and Inflammatory Effectors Defines the Universal Host Response to Severe Infection

[![DOI](https://zenodo.org/badge/DOI/10.5281/zenodo.123456.svg)](https://doi.org/10.5281/zenodo.123456)
[![Reproducibility](https://github.com/hssling/CPI_MultiDisease_Extension/actions/workflows/reproducibility.yml/badge.svg)](https://github.com/hssling/CPI_MultiDisease_Extension/actions/workflows/reproducibility.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/licenses/MIT)
[![Status: Submission Ready](https://img.shields.io/badge/Status-Submission%20Ready-green.svg)]()

## 🧬 Project Overview
Severe infections—whether bacterial (**Tuberculosis**), viral (**Dengue**), or syndromic (**Sepsis**)—converge on a shared phenotype of systemic inflammation and vascular shock. This project investigates the molecular "memory" driving this response.

We introduce the **Chromatin Priming Index (CPI)**, a single-cell metric quantifying the epigenetic potential of immune cells. By analyzing >60,000 cells across three distinct pathologies, we reveal a **Universal Epigenetic Alert State** that "locks" immune cells into a pathological response pattern.

---

## 📊 Key Discoveries

### 1. The Universal Alert State (CPI > 80%)
Across distinct etiologies, the degree of chromatin priming in immune cells is conserved (p=0.16).
- **Tuberculosis (Chronic):** CPI 84.2%
- **Sepsis (Acute):** CPI 82.5%
- **Dengue (Viral):** CPI 76.0%

### 2. The VEGFA Mechanism of Shock
We identified **Vascular Endothelial Growth Factor A (VEGFA)** as a universally epigenetically primed gene in circulating monocytes. It is transcriptionally upregulated in correlation with disease severity:
> **Epigenetic Locking:** Immune cells are "loaded" to secrete VEGFA, driving the vascular leak and hypotension characteristic of septic and dengue shock.

---

## 🔬 Methodology Pipeline

```mermaid
graph TD
    A[Raw Data: GSE151263, GSE154386] --> B[Seurat v5 Preprocessing]
    B --> C[Chromatin Accessiblity Mapping]
    C --> D[Chromatin Priming Index (CPI) Calc]
    D --> E[Cross-Disease Integration]
    E --> F{Core Signature}
    F --> G[Pathway Enrichment]
    F --> H[VEGFA Mechanism]
```

## 📂 Repository Structure

| Directory | Description |
|-----------|-------------|
| `2_analysis/` | **R Scripts:** Sepsis (`01`), Dengue (`02`), and Cross-Disease (`03`) pipelines. |
| `3_results/` | **Figures & Tables:** High-res heatmaps (`core_signature/`) and statistical tables. |
| `Submission_Package/` | **Manuscripts:** Final .docx files for *Nature Immunology* and *CID*. |
| `.github/workflows/` | **CI/CD:** Automated reproducibility checks. |

## 🚀 Reproducibility

This repository is designed for full computational reproducibility.

### Prerequisites
- Conda
- Git

### Setup
```bash
# 1. Clone the repository
git clone https://github.com/hssling/CPI_MultiDisease_Extension.git
cd CPI_MultiDisease_Extension

# 2. Create Environment
conda env create -f environment.yml
conda activate cpi_env

# 3. Generate Manuscripts (Optional)
python generate_manuscripts_docx.py
```

## 📜 Citation

If you utilize this code or data, please cite the following:

```bibtex
@article{Siddalingaiah2026,
  title = {Epigenetic Locking of Vascular and Inflammatory Effectors Defines the Universal Host Response to Severe Infection},
  author = {Siddalingaiah, H S},
  affiliation = {Shridevi Institute of Medical Sciences and Research Hospital},
  year = {2026},
  journal = {In Preparation},
  url = {https://github.com/hssling/CPI_MultiDisease_Extension}
}
```

## 👨‍⚕️ Author Information

**Dr. Siddalingaiah H S, MD**  
Professor, Department of Community Medicine  
Shridevi Institute of Medical Sciences and Research Hospital  
Tumkur, Karnataka, India - 572106  
📧 Email: hssling@yahoo.com  
🆔 ORCID: [0000-0002-4771-8285](https://orcid.org/0000-0002-4771-8285)

---
*Built with R ScRNA-seq Suite & Python Automation*
