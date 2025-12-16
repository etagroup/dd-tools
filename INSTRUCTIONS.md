# Repeat Customer Analysis - User Guide

This tool analyzes customer purchase patterns from monthly revenue data and helps identify and consolidate duplicate customer names.

## Quick Start

### Step 1: Generate Initial Analysis

Run the script with your input data:

```bash
python3 src/generate_repeat_customer_analysis.py \
  --input "path/to/your/Monthly Customer Revenue.xlsx" \
  --output output/customers.xlsx
```

**Output:**
- `output/customers.xlsx` - Analysis workbook with unconsolidated customer names
- `output/customer_master.xlsx` - **Editable mapping template** for consolidating duplicates

## Step 2: Review and Consolidate Duplicates

Open `customer_master.xlsx` in Excel. The file is organized into two sections:

### Section 1: Potential Duplicates (Top of file)
Rows 2-55 show customers with potential duplicates that need your review.

**Key Columns:**
- `customer_normalized` - The current normalized name (**don't edit this**)
- `customer_master` - **EDIT THIS COLUMN** to consolidate duplicates
- `suggested_consolidation` - Recommended master name for HIGH confidence matches
- `potential_duplicate` - What this customer might be a duplicate of
- `confidence` - Match confidence level:
  - **HIGH** - Very likely the same company (legal suffix variations, exact substring matches)
  - **MEDIUM** - Probably related, review recommended
  - **LOW** - Possibly related, manual review needed
- `merge_reason` - Explanation of why these might be duplicates
- `similarity_score` - Numeric similarity percentage

### Section 2: No Duplicates (Below separator)
Rows 59+ show customers with no detected duplicates (alphabetically sorted for reference).

### How to Consolidate

**To merge duplicates:** Edit the `customer_master` column to use the same name for both entries.

**Example:**
```
customer_normalized              customer_master                   confidence
ONTARIO POWER GENERATION    →   ONTARIO POWER GENERATION INC     HIGH
ONTARIO POWER GENERATION INC    ONTARIO POWER GENERATION INC     HIGH
```

**To keep separate:** Leave the `customer_master` column as-is (defaults to `customer_normalized`).

**Example:**
```
customer_normalized                      customer_master                          confidence
SUN LIFE ASSURANCE COMPANY               SUN LIFE ASSURANCE COMPANY               MEDIUM
SUN LIFE ASSURANCE COMPANY OF CANADA     SUN LIFE ASSURANCE COMPANY OF CANADA     MEDIUM
```

**Tip:** Use the `suggested_consolidation` column as a guide for HIGH confidence matches.

## Step 3: Apply Consolidations

After editing and saving `customer_master.xlsx`, regenerate the analysis with your mappings:

```bash
python3 src/generate_repeat_customer_analysis.py \
  --input "path/to/your/Monthly Customer Revenue.xlsx" \
  --output output/customers_consolidated.xlsx \
  --master output/customer_master.xlsx
```

**Result:**
- `output/customers_consolidated.xlsx` - Analysis with your consolidated customer names
- Your edited `customer_master.xlsx` is preserved (not overwritten)

All metrics (lifetime revenue, TTM, concentration, etc.) will now reflect the consolidated entities.

## Analysis Workbook Contents

The output workbook contains these sheets:

1. **Customer_Summary** - Customer-level metrics (tenure, gaps, reactivations, etc.)
2. **Monthly_Long** - Raw monthly revenue data
3. **Monthly_Matrix** - Pivot table of customers × months
4. **Rolling_Concentration** - Time series of top customer concentration (TTM, 24M, 36M)
5. **Top25_Lifetime** - Top 25 customers by lifetime revenue
6. **Top25_TTM** - Top 25 customers by trailing twelve months revenue
7. **Top25_PeakTTMShare** - Top 25 customers by peak concentration

**Note:** Duplicate detection and consolidation is handled through the separate `customer_master.xlsx` file, not inside the main workbook.

## Tips

- **Start with HIGH confidence matches** - These are almost always correct
- **Review MEDIUM confidence** - Usually correct but worth verifying
- **Be cautious with LOW confidence** - May be false positives
- **Filter by confidence** - Use Excel's filter dropdown to focus on HIGH matches first
- **Iterative approach** - You can re-run the consolidation multiple times, refining your mappings each time

## Requirements

- Python 3.7+
- pandas
- openpyxl
- numpy

Install dependencies:
```bash
pip install pandas openpyxl numpy
```
