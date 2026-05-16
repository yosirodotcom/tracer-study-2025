# Tracer Study 2025

Project for analyzing and visualizing alumni tracer study data. It processes raw survey data into cleaned datasets and generates comprehensive reports with charts and maps.

## Project Overview

- **Purpose:** Analyze alumni career paths, waiting times, income, and competencies.
- **Main Technologies:** Python, Pandas, Matplotlib, Geopandas, Folium.
- **Data Source:** Excel files (`.xlsx`) located in `data/raw`.
- **Outputs:** 
    - Cleaned data in `data/processed`.
    - HTML reports and static charts in `reports`.
    - Map visualizations in `assets/gambar`.

## Architecture & Workflow

1.  **Cleaning (`src/cleaning.py`):**
    - Loads `data/raw/data.xlsx`.
    - Performs deduplication (keeping latest entry based on Timestamp).
    - Fixes inconsistent categorical data (e.g., mismatched Jurusan/Prodi).
    - Maps various columns to standardized categories (Company types, Competencies, Funding sources).
    - Saves result to `data/processed/cleaned_data.xlsx`.

2.  **Analysis & Visualization (`src/viz_*.py`):**
    - Each script focuses on a specific metric (e.g., `viz_distribusi_masa_tunggu.py`).
    - Uses `src/viz_utils.py` for shared logic like HTML report generation, bar/pie/line chart creation, and geographical mapping.

3.  **Utilities (`src/viz_utils.py`):**
    - Contains robust functions for generating static maps using GeoPandas.
    - Provides base64 encoding for embedding charts in HTML reports.
    - Includes helper functions for table formatting and sorting.

## Key Directories

- `src/`: Core logic and analysis scripts.
- `scripts/`: Maintenance, verification, and debugging scripts.
- `data/`: Raw and processed datasets.
- `reports/`: Generated HTML reports and analysis summaries.
- `assets/`: Static images and final visualization outputs.
- `resources/`: Reference files like column lists and prodi mappings.

## Building and Running

### Prerequisites
Ensure Python is installed. Dependencies are often auto-installed by scripts, but recommended to have:
- `pandas`, `numpy`, `matplotlib`, `geopandas`, `folium`, `shapely`, `openpyxl`, `tabulate`.

### Key Commands
- **Clean Data:** `python src/cleaning.py`
- **Generate Reports:** Run specific visualization scripts in `src/`, e.g., `python src/viz_serapan_lulusan.py`.
- **Verify Data:** Use scripts in `scripts/` to check for data integrity, e.g., `python scripts/verify_columns.py`.

## Development Conventions

- **Data Safety:** Never modify `data/raw/` directly; always work from `data/processed/`.
- **Visualization:** Use `src/viz_utils.py` for consistent styling across charts and reports.
- **Error Handling:** Scripts often include auto-install checks for missing libraries.
- **Naming:** Follow `snake_case` for filenames and variables.
