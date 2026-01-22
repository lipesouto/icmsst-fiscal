# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

OmniAI Fiscal is a Streamlit web application for calculating PIS/COFINS tax credits from ICMS-ST exclusion from the tax base. It processes Brazilian SPED (Sistema Publico de Escrituracao Digital) fiscal files and generates consolidated reports.

## Commands

```bash
# Install dependencies
pip install -r requirements.txt

# Run the application
streamlit run app.py

# Docker build and run
docker build -t omniai-fiscal .
docker run -p 8501:8501 omniai-fiscal
```

The app runs on `http://localhost:8501`.

## Architecture

The entire application is contained in a single file `app.py` with the following structure:

### Data Classes (lines 206-287)
- `SpedHeader`: SPED file header information (block 0000)
- `ProductInfo`: Product data from block 0200
- `C870Record`: Individual sales record from block C870
- `CalculationResult`: Result of ICMS-ST exclusion calculation
- `MonthSummary`: Monthly aggregation of results

### Processing Classes (lines 293-613)
- `SpedParser`: Parses SPED Contribuicoes files (Latin-1 encoded TXT), extracts header (0000), products (0200), and sales records (C870)
- `ProductBaseLoader`: Loads Excel product base with NCM codes and MVA (Margem de Valor Agregado) percentages
- `IcmsStCalculator`: Core calculation logic - computes ICMS-ST values and new PIS/COFINS bases
- `SpedWriter`: Generates rectified SPED files with updated tax bases

### Calculation Logic (`IcmsStCalculator.calculate`)
1. Filter by eligible CFOPs (default: 5405)
2. Match product NCM to MVA from product base (via REG 0200 COD_ITEM)
3. Calculate exclusion: `Exclusão = VL_BC_PIS/COFINS * MVA% * ALIQ_ICMS%`
4. New tax base: `BC Nova = VL_BC_PIS/COFINS - Exclusão`
5. New tax values: `VL_PIS/COFINS_novo = BC Nova * ALIQ_PIS/COFINS`
6. Credit = difference between original and recalculated taxes

### Output Generators (lines 620-834)
- `generate_excel()`: Creates multi-sheet Excel with summary and per-month details
- `generate_pdf()`: Creates executive PDF report using ReportLab

### Streamlit Interface (lines 840-1211)
- Sidebar: CFOP selection (5405, 5403, 5401, 5102)
- Main: File uploaders for product base (Excel) and SPED files (TXT)
- Results: Metrics, monthly table, download buttons (Excel, PDF, ZIP with rectified SPEDs, JSON)

## Key Technical Details

- SPED files use Latin-1 encoding
- All monetary calculations use `Decimal` with `ROUND_HALF_UP` to 2 decimal places
- NCM codes are 8 digits, can be provided as full NCM or as Capitulo (4 digits) + Item (4 digits)
- Month/year extracted from filename pattern `MM_YYYY` or `MM-YYYY`
