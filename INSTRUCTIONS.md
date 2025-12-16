# Repeat Customer Analysis - User Guide

This tool analyzes customer purchase patterns from monthly revenue data and helps identify and consolidate duplicate customer names.

## Architecture Overview

The analysis is split into three phases:

1. **Preparation** (`prepare.py`) - Monthly task when books close
   - Load and normalize raw revenue data
   - Detect potential duplicate customer names
   - Apply customer master mappings

2. **Analytics** (`analytics.py`) - Run as needed
   - Generate customer summary metrics
   - Calculate rolling concentration
   - Create analytics workbook

3. **Reporting** (`high_value_report.py`, `customer_segment_matrix.py`) - Run frequently
   - Generate console reports from analytics workbook

## Quick Start

### Step 1: Prepare Data

Run the preparation script with your input data:

```bash
python src/prepare.py --input "path/to/Monthly Customer Revenue.xlsx"
```

**Output files** (derived from input filename):
- `{basename}.prepared.xlsx` - Flattened revenue data (Revenue_Detail sheet)
- `{basename}.customers.xlsx` - Customer master for duplicate consolidation

### Step 2: Review and Consolidate Duplicates

Open the `.customers.xlsx` file in Excel. The file is organized into two sections:

#### Section 1: Potential Duplicates (Top of file)
Customers with potential duplicates that need your review.

**Key Columns:**
- `customer_normalized` - The current normalized name (**don't edit this**)
- `customer_master` - **EDIT THIS COLUMN** to consolidate duplicates
- `suggested_consolidation` - Recommended master name for HIGH confidence matches
- `potential_duplicate` - What this customer might be a duplicate of
- `confidence` - Match confidence level:
  - **HIGH** - Very likely the same company (legal suffix variations, exact substring matches)
  - **MEDIUM** - Probably related, review recommended
  - **LOW** - Possibly related, manual review needed
- `is_new` - TRUE if this customer is new (not in previous master)

#### Section 2: No Duplicates (Below separator)
Customers with no detected duplicates (alphabetically sorted for reference).

#### How to Consolidate

**To merge duplicates:** Edit the `customer_master` column to use the same name for both entries.

**Example:**
```
customer_normalized              customer_master                   confidence
ONTARIO POWER GENERATION    →   ONTARIO POWER GENERATION INC     HIGH
ONTARIO POWER GENERATION INC    ONTARIO POWER GENERATION INC     HIGH
```

**To keep separate:** Leave the `customer_master` column as-is.

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

Run the analytics script on the prepared data:

```bash
python src/analytics.py \
  --input "path/to/data.prepared.xlsx" \
  --output "customer_analytics.xlsx"
```

### Step 5: Run Reports

Generate reports from the analytics workbook:

```bash
# High-value customer report
python src/high_value_report.py customer_analytics.xlsx

# Customer segment matrix (3x2: Status x Value)
python src/customer_segment_matrix.py customer_analytics.xlsx
```

## File Outputs

### Prepared Data (`{basename}.prepared.xlsx`)

Single sheet:
- **Revenue_Detail** - Flattened monthly revenue data (date, year, month, customer, revenue)

### Customer Master (`{basename}.customers.xlsx`)

Single sheet:
- **Customer_Master** - All customers with duplicate detection results

### Analytics Workbook

Contains these sheets:
1. **Metadata** - Data date range and generation timestamp
2. **Customer_Summary** - Customer-level metrics with segmentation formulas
3. **Monthly_Matrix** - Pivot table of customers x months
4. **Rolling_Concentration** - Time series of top customer concentration (TTM, 24M, 36M)
5. **Top25_Lifetime** - Top 25 customers by lifetime revenue
6. **Top25_TTM** - Top 25 customers by trailing twelve months revenue
7. **Top25_PeakTTMShare** - Top 25 customers by peak concentration

## Customer Summary Columns

The Customer_Summary sheet includes:

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

## Tips

- **Start with HIGH confidence matches** - These are almost always correct
- **Review MEDIUM confidence** - Usually correct but worth verifying
- **Be cautious with LOW confidence** - May be false positives
- **Filter by is_new** - Focus on newly added customers each period
- **Iterative approach** - You can re-run preparation multiple times, refining mappings

## Requirements

- Python 3.7+
- pandas
- openpyxl
- numpy

Install dependencies:
```bash
pip install pandas openpyxl numpy
```
