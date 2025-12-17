# Repeat Customer Purchase-Pattern Analysis

Analyzes customer purchase patterns from monthly revenue data, focusing on **repeat customers with discontinuous ("lumpy") spending** where a single fiscal-year view can misstate customer importance and revenue concentration.

**Key features:**
- Automated duplicate customer detection (fuzzy matching)
- Manual consolidation workflow via editable Excel file
- Rolling concentration metrics (TTM, 24M, 36M)
- Customer-level behavioral metrics (tenure, gaps, reactivations)

## Architecture

The analysis is split into three phases:

1. **Preparation** (`prepare.py`) - Monthly task when books close
2. **Analytics** (`analytics.py`) - Run as needed
3. **Reporting** (`customer_churn_report.py`, `customer_segment_matrix.py`) - Run frequently
4. **Visualization** (`visualize.py`) - Generate charts

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
│ - churn_report  │
│ - segment_matrix│
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│  visualize.py   │ ──► PNG charts
└─────────────────┘
```

## Setup

```bash
./setup.sh
# or manually:
pip install pandas openpyxl numpy matplotlib fpdf2 pyyaml
```

**Requirements:** Python 3.7+, pandas, openpyxl, numpy, matplotlib, fpdf2, pyyaml

## Input Data Format

- Excel workbook with sheets named by year: `2020`, `2021`, etc.
- Each sheet contains a "Name of Customer" header row with month columns
- Customer names are auto-normalized (uppercase + whitespace collapse)
- Data automatically trimmed to last month with non-zero revenue

## Workflow

### Step 1: Prepare Data

```bash
python src/prepare.py --input "path/to/Monthly Customer Revenue.xlsx"
```

**Output files** (derived from input filename):
- `{basename}.prepared.xlsx` - Flattened revenue data (Revenue_Detail sheet)
- `{basename}.customers.xlsx` - Customer master for duplicate consolidation

### Step 2: Review and Consolidate Duplicates

Open the `.customers.xlsx` file in Excel. The file is organized into two sections:

**Section 1: Potential Duplicates (Top of file)**

Customers with potential duplicates that need your review.

Key columns:
- `customer_normalized` - The current normalized name (**don't edit this**)
- `customer_master` - **EDIT THIS COLUMN** to consolidate duplicates
- `suggested_consolidation` - Recommended master name for HIGH confidence matches
- `potential_duplicate` - What this customer might be a duplicate of
- `confidence` - Match confidence level:
  - **HIGH** - Very likely the same company (legal suffix variations, exact substring matches)
  - **MEDIUM** - Probably related, review recommended
  - **LOW** - Possibly related, manual review needed
- `is_new` - TRUE if this customer is new (not in previous master)

**Section 2: No Duplicates (Below separator)**

Customers with no detected duplicates (alphabetically sorted for reference).

**How to Consolidate:**

To merge duplicates, edit the `customer_master` column to use the same name for both entries:
```
customer_normalized              customer_master                   confidence
ACME                        →   ACME INC                         HIGH
ACME INC                        ACME INC                         HIGH
```

To keep separate, leave the `customer_master` column as-is.

### Step 3: Re-run Preparation with Master

After editing the customer master, re-run preparation to apply your consolidations:

```bash
python src/prepare.py \
  --input "path/to/Monthly Customer Revenue.xlsx" \
  --master "path/to/previous.customers.xlsx"
```

The script will:
- Preserve your existing mappings
- Mark new customers with `is_new = TRUE` for review
- Apply consolidations to the prepared data

### Step 4: Generate Analytics

```bash
python src/analytics.py \
  --input "path/to/data.prepared.xlsx" \
  --output "customer_analytics.xlsx"
```

### Step 5: Run Reports

```bash
# Customer churn report (--high-value default, or --low-value, --all)
python src/customer_churn_report.py customer_analytics.xlsx

# Customer segment matrix (3x2: Status x Value)
python src/customer_segment_matrix.py customer_analytics.xlsx

# PDF output (add --pdf flag to any report)
python src/customer_churn_report.py customer_analytics.xlsx --pdf churn_report.pdf
```

### Step 6: Generate Visualizations

```bash
python src/visualize.py customer_analytics.xlsx --output-dir ./charts
```

Generates:
- `pareto_curve.png` - Revenue concentration (80/20 analysis)
- `concentration_trend.png` - Top customer share over time
- `segment_heatmap.png` - Customer segment matrix visualization

## File Outputs

### Prepared Data (`{basename}.prepared.xlsx`)

- **Revenue_Detail** - Flattened monthly revenue data (date, year, month, customer, revenue)

### Customer Master (`{basename}.customers.xlsx`)

- **Customer_Master** - All customers with duplicate detection results

### Analytics Workbook

1. **Metadata** - Data date range and generation timestamp
2. **Customer_Summary** - Customer-level metrics with segmentation formulas
3. **Monthly_Matrix** - Pivot table of customers x months
4. **Rolling_Concentration** - Time series of top customer concentration (TTM, 24M, 36M)
5. **Top25_Lifetime** - Top 25 customers by lifetime revenue
6. **Top25_TTM** - Top 25 customers by trailing twelve months revenue
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
- `status` - Active (≤6 months), Inactive (7-18), Churned (>18)
- `is_high_value` - Lifetime ≥ $1M OR peak share ≥ 2%
- `segment` - Strategic / High Value / Mid Value (Active customers only)

## Configuration

Thresholds for customer classification are defined in `etc/config.yaml`:

```yaml
high_value:
  lifetime_revenue_min: 1000000    # Lifetime revenue >= $1M
  peak_ttm_share_min: 0.02         # Peak TTM share >= 2%

status:
  active_max: 6          # Active: <= 6 months since last purchase
  inactive_max: 18       # Inactive: 7-18 months
                         # Churned: > 18 months
```

Edit this file to adjust thresholds without modifying code.

## Tips

- **Start with HIGH confidence matches** - These are almost always correct
- **Review MEDIUM confidence** - Usually correct but worth verifying
- **Be cautious with LOW confidence** - May be false positives
- **Filter by is_new** - Focus on newly added customers each period
- **Iterative approach** - You can re-run preparation multiple times, refining mappings
