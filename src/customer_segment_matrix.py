#!/usr/bin/env python3
"""Generate a 3x2 customer segment matrix (Status x Value)."""

import argparse
import sys
import pandas as pd


def build_segment_matrix(input_file: str) -> tuple:
    """
    Build a 3x2 matrix of customer counts and revenue.

    Columns (Status): Active, Inactive, Churned
    Rows (Value): High Value, Not High Value

    Returns: (counts, revenue, data_start, data_end)
    """
    # Read Customer_Summary sheet
    df = pd.read_excel(input_file, sheet_name="Customer_Summary")

    # Read Metadata sheet for date range
    metadata = pd.read_excel(input_file, sheet_name="Metadata")
    metadata_dict = dict(zip(metadata["property"], metadata["value"]))
    data_start = pd.to_datetime(metadata_dict["data_start_date"])
    data_end = pd.to_datetime(metadata_dict["data_end_date"])

    # Use the most recent month in the data as the reference date
    reference_date = data_end
    df['last_purchase_month'] = pd.to_datetime(df['last_purchase_month'])
    df['months_since_last_purchase_calc'] = ((reference_date - df['last_purchase_month']).dt.days / 30.44).round(0)

    # Recalculate status
    df['status_calc'] = df['months_since_last_purchase_calc'].apply(
        lambda x: 'Active' if x <= 6 else ('Inactive' if x <= 18 else 'Churned')
    )

    # Recalculate is_high_value
    df['is_high_value_calc'] = (df['lifetime_revenue'] >= 1_000_000) | (df['peak_ttm_share'] >= 0.02)

    # Create cross-tabulation for counts
    counts = pd.crosstab(
        index=df['is_high_value_calc'].map({True: 'High Value', False: 'Not High Value'}),
        columns=df['status_calc'],
        margins=True,
        margins_name='Total'
    )

    # Create cross-tabulation for revenue
    revenue = pd.crosstab(
        index=df['is_high_value_calc'].map({True: 'High Value', False: 'Not High Value'}),
        columns=df['status_calc'],
        values=df['lifetime_revenue'],
        aggfunc='sum',
        margins=True,
        margins_name='Total'
    )

    # Reorder columns: Active, Inactive, Churned, Total
    col_order = ['Active', 'Inactive', 'Churned', 'Total']
    counts = counts[[c for c in col_order if c in counts.columns]]
    revenue = revenue[[c for c in col_order if c in revenue.columns]]

    # Clean up index/column names
    counts.index.name = None
    counts.columns.name = None
    revenue.index.name = None
    revenue.columns.name = None

    return counts, revenue, data_start, data_end


def format_currency(val):
    """Format value as currency (millions)."""
    if pd.isna(val):
        return '-'
    return f'${val/1_000_000:.1f}M'


def generate_report(input_file: str):
    """Generate report lines."""
    counts, revenue, data_start, data_end = build_segment_matrix(input_file)

    lines = []
    lines.append("")
    lines.append("=" * 70)
    lines.append("CUSTOMER SEGMENT MATRIX")
    lines.append(f"Data Range: {data_start.strftime('%b %Y')} - {data_end.strftime('%b %Y')}")
    lines.append("=" * 70)

    lines.append("")
    lines.append("CUSTOMER COUNTS:")
    lines.append(counts.to_string())

    lines.append("")
    lines.append("")
    lines.append("LIFETIME REVENUE:")
    revenue_formatted = revenue.map(format_currency)
    lines.append(revenue_formatted.to_string())

    lines.append("")
    lines.append("=" * 70)

    # Calculate some key metrics
    total_customers = counts.loc['Total', 'Total']
    total_revenue = revenue.loc['Total', 'Total']
    high_value_customers = counts.loc['High Value', 'Total']
    high_value_revenue = revenue.loc['High Value', 'Total']

    lines.append("")
    lines.append("KEY INSIGHTS:")
    lines.append(f"  Total Customers: {total_customers:.0f}")
    lines.append(f"  Total Revenue: {format_currency(total_revenue)}")
    lines.append(f"  High Value Customers: {high_value_customers:.0f} ({100*high_value_customers/total_customers:.1f}%)")
    lines.append(f"  High Value Revenue: {format_currency(high_value_revenue)} ({100*high_value_revenue/total_revenue:.1f}%)")

    # Print status breakdowns if columns exist
    for status in ['Active', 'Inactive', 'Churned']:
        if status in counts.columns:
            hv_count = counts.loc['High Value', status] if 'High Value' in counts.index else 0
            hv_rev = revenue.loc['High Value', status] if 'High Value' in revenue.index else 0
            lines.append(f"  {status} High Value: {hv_count:.0f} customers, {format_currency(hv_rev)}")

    lines.append("=" * 70)
    lines.append("")

    return lines


def write_pdf(lines, output_path):
    """Write report lines to PDF."""
    from fpdf import FPDF
    from fpdf.enums import XPos, YPos

    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.set_font('Courier', size=9)

    for line in lines:
        # Handle special characters
        safe_line = line.encode('latin-1', errors='replace').decode('latin-1')
        pdf.cell(0, 4, safe_line, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.output(output_path)
    print(f"Wrote: {output_path}", file=sys.stderr)


def main(input_file: str, pdf_output: str = None):
    lines = generate_report(input_file)

    if pdf_output:
        write_pdf(lines, pdf_output)
    else:
        for line in lines:
            print(line)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Generate a 3x2 customer segment matrix (Status x Value)"
    )
    parser.add_argument("input_file", help="Path to customer analysis Excel file")
    parser.add_argument("--pdf", metavar="FILE", help="Output to PDF file instead of console")
    args = parser.parse_args()
    main(args.input_file, args.pdf)
