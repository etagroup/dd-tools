#!/usr/bin/env python3
"""Generate a 3x2 customer segment matrix (Status x Value)."""

import argparse
import pandas as pd
from datetime import datetime


def build_segment_matrix(input_file: str) -> tuple:
    """
    Build a 3x2 matrix of customer counts and revenue.

    Columns (Status): Active, Inactive, Churned
    Rows (Value): High Value, Not High Value

    Returns: (counts, revenue, data_start, data_end)
    """
    # Read Customer_Summary sheet
    df = pd.read_excel(input_file, sheet_name="Customer_Summary")

    # Read Revenue_Detail sheet to determine data range
    df_detail = pd.read_excel(input_file, sheet_name="Revenue_Detail")
    df_detail['date'] = pd.to_datetime(df_detail['date'])
    data_start = df_detail['date'].min()
    data_end = df_detail['date'].max()

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

    return counts, revenue, data_start, data_end


def format_currency(val):
    """Format value as currency (millions)."""
    if pd.isna(val):
        return '-'
    return f'${val/1_000_000:.1f}M'


def main(input_file: str):
    counts, revenue, data_start, data_end = build_segment_matrix(input_file)

    print("\n" + "="*70)
    print("CUSTOMER SEGMENT MATRIX")
    print(f"Data Range: {data_start.strftime('%b %Y')} – {data_end.strftime('%b %Y')}")
    print("="*70)

    print("\nCUSTOMER COUNTS:")
    print(counts.to_string())

    print("\n\nLIFETIME REVENUE:")
    print(revenue.to_string())

    print("\n\nLIFETIME REVENUE (Formatted):")
    revenue_formatted = revenue.map(format_currency)
    print(revenue_formatted.to_string())

    print("\n" + "="*70)

    # Calculate some key metrics
    total_customers = counts.loc['Total', 'Total']
    total_revenue = revenue.loc['Total', 'Total']
    high_value_customers = counts.loc['High Value', 'Total']
    high_value_revenue = revenue.loc['High Value', 'Total']

    print(f"\nKEY INSIGHTS:")
    print(f"  Total Customers: {total_customers:.0f}")
    print(f"  Total Revenue: {format_currency(total_revenue)}")
    print(f"  High Value Customers: {high_value_customers:.0f} ({100*high_value_customers/total_customers:.1f}%)")
    print(f"  High Value Revenue: {format_currency(high_value_revenue)} ({100*high_value_revenue/total_revenue:.1f}%)")

    # Print status breakdowns if columns exist
    for status in ['Active', 'Inactive', 'Churned']:
        if status in counts.columns:
            hv_count = counts.loc['High Value', status] if 'High Value' in counts.index else 0
            hv_rev = revenue.loc['High Value', status] if 'High Value' in revenue.index else 0
            print(f"  {status} High Value: {hv_count:.0f} customers, {format_currency(hv_rev)}")

    print("="*70 + "\n")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Generate a 3x2 customer segment matrix (Status x Value)"
    )
    parser.add_argument("input_file", help="Path to customer analysis Excel file")
    args = parser.parse_args()
    main(args.input_file)
