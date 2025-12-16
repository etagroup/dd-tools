# Repeat Customer Purchase-Pattern Analysis

Analyzes customer purchase patterns from monthly revenue data, focusing on **repeat customers with discontinuous ("lumpy") spending** where a single fiscal-year view can misstate customer importance and revenue concentration.

**Key features:**
- Automated duplicate customer detection (fuzzy matching)
- Manual consolidation workflow via editable Excel file
- Rolling concentration metrics (TTM, 24M, 36M)
- Customer-level behavioral metrics (tenure, gaps, reactivations)

## Quick Start

See [INSTRUCTIONS.md](INSTRUCTIONS.md) for detailed workflow documentation.

### 1. Generate Initial Analysis

```bash
python src/customer_analysis.py \
  --input "path/to/Monthly Customer Revenue.xlsx" \
  --output output/customer_analysis.xlsx
```

**Output:**
- `customer_analysis.xlsx` - Analysis workbook
- `customer_master.xlsx` - Editable duplicate consolidation template

### 2. Consolidate Duplicates (Optional)

Edit `customer_master.xlsx` to merge duplicate customer names, then regenerate:

```bash
python src/customer_analysis.py \
  --input "path/to/Monthly Customer Revenue.xlsx" \
  --output output/customer_analysis.xlsx \
  --master output/customer_master.xlsx
```

## Analysis Workbook Contents

- **Customer_Summary** – customer-level metrics (tenure, gaps, reactivations, peak concentration)
- **Monthly_Long** – long-format month-by-customer revenue
- **Monthly_Matrix** – pivoted matrix (months × customers)
- **Rolling_Concentration** – time series of top customer concentration (TTM/24M/36M, Top1/5/10 share)
- **Top25_Lifetime** / **Top25_TTM** / **Top25_PeakTTMShare** – ranked customer lists

## Duplicate Detection

The script automatically detects potential duplicate customer names using:
- Legal suffix variations (Inc./Corp./Ltd.)
- Fuzzy string matching with distinctive name analysis
- Confidence levels: HIGH (auto-group) / MEDIUM (review) / LOW (possible false positive)

## Setup

```bash
pip install -r src/requirements.txt
```

**Requirements:** Python 3.7+, pandas, openpyxl, numpy

## Input Data Format

- Excel workbook with sheets named by year: `2020`, `2021`, etc.
- Each sheet contains a "Name of Customer" header row with month columns
- Customer names are auto-normalized (uppercase + whitespace collapse)
- Data automatically trimmed to last month with non-zero revenue

