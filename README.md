# Repeat Customer Purchase-Pattern Analysis

Analyzes customer purchase patterns from monthly revenue data, focusing on **repeat customers with discontinuous ("lumpy") spending** where a single fiscal-year view can misstate customer importance and revenue concentration.

**Key features:**
- Automated duplicate customer detection (fuzzy matching)
- Manual consolidation workflow via editable Excel file
- Rolling concentration metrics (TTM, 24M, 36M)
- Customer-level behavioral metrics (tenure, gaps, reactivations)

## Architecture

The analysis is split into three phases:

1. **Preparation** (monthly) - Data loading, normalization, customer master management
2. **Analytics** (as needed) - Summaries, rolling metrics, concentration analysis
3. **Reporting** (frequent) - Console reports from analytics workbook

```
Raw Input Excel
      │
      ▼
┌─────────────────┐
│   prepare.py    │ ──► {basename}.prepared.xlsx
│                 │ ──► {basename}.customers.xlsx
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│  analytics.py   │ ──► customer_analytics.xlsx
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│    Reports      │ ──► Console output
│ - high_value    │
│ - segment_matrix│
└─────────────────┘
```

## Quick Start

See [INSTRUCTIONS.md](INSTRUCTIONS.md) for detailed workflow documentation.

### 1. Prepare Data

```bash
python src/prepare.py --input "4.0.5 Financial - Month-wise Customer Revenue.xlsx"
```

**Output:**
- `4.0.5 Financial...prepared.xlsx` - Flattened, normalized revenue data
- `4.0.5 Financial...customers.xlsx` - Customer master with duplicate detection

### 2. Review Customer Master (Optional)

Edit the `.customers.xlsx` file to consolidate duplicates, then re-run prepare with the master:

```bash
python src/prepare.py \
  --input "4.0.5 Financial - Month-wise Customer Revenue.xlsx" \
  --master "4.0.4 Financial...customers.xlsx"
```

### 3. Generate Analytics

```bash
python src/analytics.py \
  --input "4.0.5 Financial...prepared.xlsx" \
  --output "customer_analytics.xlsx"
```

### 4. Run Reports

```bash
python src/high_value_report.py customer_analytics.xlsx
python src/customer_segment_matrix.py customer_analytics.xlsx
```

## Analytics Workbook Contents

- **Metadata** - Data date range and generation timestamp
- **Customer_Summary** - Customer-level metrics (tenure, gaps, reactivations, peak concentration)
- **Monthly_Matrix** - Pivoted matrix (months x customers)
- **Rolling_Concentration** - Time series of top customer concentration (TTM/24M/36M, Top1/5/10 share)
- **Top25_Lifetime** / **Top25_TTM** / **Top25_PeakTTMShare** - Ranked customer lists

## Duplicate Detection

The script automatically detects potential duplicate customer names using:
- Legal suffix variations (Inc./Corp./Ltd.)
- Fuzzy string matching with distinctive name analysis
- Confidence levels: HIGH (auto-group) / MEDIUM (review) / LOW (possible false positive)

## Setup

```bash
pip install pandas openpyxl numpy
```

**Requirements:** Python 3.7+, pandas, openpyxl, numpy

## Input Data Format

- Excel workbook with sheets named by year: `2020`, `2021`, etc.
- Each sheet contains a "Name of Customer" header row with month columns
- Customer names are auto-normalized (uppercase + whitespace collapse)
- Data automatically trimmed to last month with non-zero revenue
