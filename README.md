# Repeat Customer Purchase-Pattern Analysis

Analyzes customer purchase patterns from monthly revenue data, focusing on **repeat customers with discontinuous ("lumpy") spending** where a single fiscal-year view can misstate customer importance and revenue concentration.

**Key features:**
- Automated duplicate customer detection (fuzzy matching)
- Manual consolidation workflow via editable Excel file
- Rolling concentration metrics (TTM, 24M, 36M)
- Customer-level behavioral metrics (tenure, gaps, reactivations)
- Reproducible output (deterministic checksums)

## Quick Start

```bash
# Setup
./setup.sh

# Run full pipeline
./run.sh all --input "Monthly Revenue.xlsx" --outdir output/

# With reports and charts
./run.sh all --input "Monthly Revenue.xlsx" --outdir output/ --reports --charts
```

## Architecture

```
Raw Input Excel
      │
      ▼
┌─────────────────┐
│ prepare-customers│ ──► customers.xlsx
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│ prepare-revdata │ ──► revdata.xlsx
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│  run-analytics  │ ──► analytics.xlsx
└────────┬────────┘
         │
    ┌────┴────┐
    ▼         ▼
┌────────┐ ┌────────┐
│reports │ │ charts │
│ (PDF)  │ │ (PNG)  │
└────────┘ └────────┘
```

## Setup

```bash
./setup.sh
# or manually:
pip install pandas openpyxl numpy matplotlib fpdf2 pyyaml
```

**Requirements:** Python 3.7+

## CLI Reference

```bash
./run.sh <command> [options]

Commands:
  all               Run full pipeline (prepare → analytics)
  prepare-customers Create/update customer master from revenue data
  prepare-revdata   Flatten revenue data and apply customer mappings
  run-analytics     Generate analytics workbook from prepared data
  gen-reports       Generate console or PDF reports
  gen-charts        Generate PNG visualizations

Run './run.sh <command> --help' for command-specific options.
```

### all - Full Pipeline

```bash
./run.sh all --input <file> [options]

Options:
  --input <file>     Raw revenue Excel file (required)
  --existing <file>  Existing customer master (auto-detects {outdir}/customers.xlsx)
  --outdir <dir>     Output directory (default: current dir)
  --merge            Auto-apply HIGH confidence consolidations
  --reports          Also generate PDF reports
  --charts           Also generate PNG charts
```

### prepare-customers

```bash
./run.sh prepare-customers --input <file> [options]

Options:
  --input <file>     Raw revenue Excel file (required)
  --existing <file>  Existing customer master to update
  --merge            Auto-apply HIGH confidence consolidations
  --outdir <dir>     Output directory (default: current dir)

Output: {outdir}/customers.xlsx
```

### prepare-revdata

```bash
./run.sh prepare-revdata --input <file> [options]

Options:
  --input <file>     Raw revenue Excel file (required)
  --customers <file> Customer master to apply (auto-detects if not specified)
  --outdir <dir>     Output directory (default: current dir)

Output: {outdir}/revdata.xlsx
```

### run-analytics

```bash
./run.sh run-analytics --revdata <file> [options]

Options:
  --revdata <file>   Prepared revenue data file (required)
  --outdir <dir>     Output directory (default: current dir)

Output: {outdir}/analytics.xlsx
```

### gen-reports

```bash
./run.sh gen-reports --analytics <file> [options]

Options:
  --analytics <file>  Analytics workbook (required)
  --outdir <dir>      Output directory for PDFs (default: current dir)
  --pdf               Generate PDF files instead of console output
  --all               Show all customers (high-value and other)
  --high-value        Show high-value customers only (default)
  --low-value         Show low-value customers only

Output (with --pdf): {outdir}/churn_report.pdf, {outdir}/segment_matrix.pdf
```

### gen-charts

```bash
./run.sh gen-charts --analytics <file> [options]

Options:
  --analytics <file>  Analytics workbook (required)
  --outdir <dir>      Output directory (default: ./charts)

Output: pareto_curve.png, concentration_trend.png, segment_heatmap.png
```

## Input Data Format

- Excel workbook with sheets named by year: `2020`, `2021`, etc.
- Each sheet contains a "Name of Customer" header row with month columns
- Customer names are auto-normalized (uppercase + whitespace collapse)
- Data automatically trimmed to last month with non-zero revenue

## Workflow

### First Run

```bash
./run.sh all --input "Monthly Revenue.xlsx" --outdir output/
```

### Review Duplicates

Open `output/customers.xlsx` in Excel. The file has two sections:

**Section 1: Potential Duplicates** (top)
- Review `confidence` column: HIGH, MEDIUM, LOW
- Edit `customer_master` column to consolidate duplicates

**Section 2: No Duplicates** (below separator)
- Alphabetically sorted for reference

**To consolidate duplicates**, set the same `customer_master` value for both entries:
```
customer_normalized    customer_master    confidence
ACME                   ACME INC           HIGH
ACME INC               ACME INC           HIGH
```

### Subsequent Runs

```bash
# Auto-detects existing customers.xlsx in outdir
./run.sh all --input "Monthly Revenue.xlsx" --outdir output/

# Or specify explicitly
./run.sh all --input "Monthly Revenue.xlsx" --outdir output/ --existing prev/customers.xlsx
```

## Output Files

| File | Description |
|------|-------------|
| `customers.xlsx` | Customer master with duplicate detection |
| `revdata.xlsx` | Flattened revenue data (Revenue_Detail sheet) |
| `analytics.xlsx` | Full analytics workbook (6 sheets) |
| `churn_report.pdf` | Customer churn report by status |
| `segment_matrix.pdf` | 3x2 customer segment matrix |
| `*.png` | Visualization charts |

### Analytics Workbook Sheets

1. **Metadata** - Data date range
2. **Customer_Summary** - Customer-level metrics with segmentation formulas
3. **Monthly_Matrix** - Pivot table of customers x months
4. **Rolling_Concentration** - Time series of top customer concentration
5. **Top25_Lifetime** - Top 25 customers by lifetime revenue
6. **Top25_TTM** - Top 25 customers by trailing twelve months
7. **Top25_PeakTTMShare** - Top 25 customers by peak concentration

## Customer Summary Columns

**Revenue Metrics:**
- `lifetime_revenue` - Total revenue across all time
- `ttm_revenue_last` - Trailing 12-month revenue at most recent month
- `t24m_revenue_last`, `t36m_revenue_last` - 24/36-month trailing revenue

**Activity Metrics:**
- `first_purchase_month`, `last_purchase_month` - Purchase date range
- `active_months`, `active_quarters`, `active_years` - Activity counts
- `tenure_months` - Months from first to last purchase
- `activity_ratio` - Active months / tenure months

**Behavioral Metrics:**
- `avg_gap_months`, `max_gap_months` - Purchase gap statistics
- `reactivations` - Times customer returned after a gap
- `is_repeat_customer` - Active in 2+ distinct years
- `peak_ttm_share` - Highest monthly share during any TTM period
- `top10_ttm_persistence` - Fraction of TTM periods in top 10

**Segmentation (Formula-based):**
- `months_since_last_purchase` - Months from last purchase to data end
- `status` - Active (<=6 months), Inactive (7-18), Churned (>18)
- `is_high_value` - Lifetime >= $1M OR peak share >= 2%
- `segment` - Strategic / High Value / Mid Value (Active customers only)

## Configuration

Thresholds are defined in `etc/config.yaml`:

```yaml
high_value:
  lifetime_revenue_min: 1000000    # Lifetime revenue >= $1M
  peak_ttm_share_min: 0.02         # Peak TTM share >= 2%

status:
  active_max: 6          # Active: <= 6 months since last purchase
  inactive_max: 18       # Inactive: 7-18 months
                         # Churned: > 18 months
```

## Tips

- **Start with HIGH confidence matches** - These are almost always correct
- **Review MEDIUM confidence** - Usually correct but worth verifying
- **Be cautious with LOW confidence** - May be false positives
- **Filter by is_new** - Focus on newly added customers each period
- **Iterative approach** - Re-run as needed, refining mappings over time
- **Use `--merge`** - Auto-apply HIGH confidence consolidations to save time
