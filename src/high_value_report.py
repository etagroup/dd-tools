#!/usr/bin/env python3
"""List high value customers grouped by status."""

import argparse
import pandas as pd
from datetime import datetime


def format_currency(val):
    """Format value as currency (millions)."""
    if pd.isna(val):
        return '-'
    if val >= 1_000_000:
        return f'${val/1_000_000:.1f}M'
    else:
        return f'${val/1_000:,.0f}K'

def main(input_file: str):
    # Read Customer_Summary sheet
    df = pd.read_excel(input_file, sheet_name="Customer_Summary")

    # Read Revenue_Detail sheet to determine data range
    df_monthly = pd.read_excel(input_file, sheet_name="Revenue_Detail")
    df_monthly['date'] = pd.to_datetime(df_monthly['date'])
    data_start = df_monthly['date'].min()
    data_end = df_monthly['date'].max()

    # Use the most recent month in the data as the reference date
    reference_date = data_end
    df['last_purchase_month'] = pd.to_datetime(df['last_purchase_month'])
    df['months_since'] = ((reference_date - df['last_purchase_month']).dt.days / 30.44).round(0)

    # Calculate status
    df['status'] = df['months_since'].apply(
        lambda x: 'Active' if x <= 6 else ('Inactive' if x <= 18 else 'Churned')
    )

    # Calculate is_high_value
    df['is_high_value'] = (df['lifetime_revenue'] >= 1_000_000) | (df['peak_ttm_share'] >= 0.02)

    # Filter for high value customers only
    hv = df[df['is_high_value']].copy()

    # Sort by status (Active first) then by lifetime revenue descending
    status_order = {'Active': 0, 'Inactive': 1, 'Churned': 2}
    hv['status_sort'] = hv['status'].map(status_order)
    hv = hv.sort_values(['status_sort', 'lifetime_revenue'], ascending=[True, False])

    print("\n" + "="*90)
    print("HIGH VALUE CUSTOMERS (Lifetime ≥ $1M OR Peak TTM Share ≥ 2%)")
    print(f"Data Range: {data_start.strftime('%b %Y')} – {data_end.strftime('%b %Y')}")
    print("="*90)

    for status in ['Active', 'Inactive', 'Churned']:
        status_df = hv[hv['status'] == status]

        if len(status_df) == 0:
            continue

        total_rev = status_df['lifetime_revenue'].sum()
        avg_rev = status_df['lifetime_revenue'].mean()

        print(f"\n{status.upper()} ({len(status_df)} customers, {format_currency(total_rev)} total)")
        print("-" * 90)
        print(f"{'Customer':<50} {'Lifetime':>12} {'TTM':>12} {'Peak Share':>10} {'Months Since':>10}")
        print("-" * 90)

        for _, row in status_df.iterrows():
            print(f"{row['customer'][:48]:<50} "
                  f"{format_currency(row['lifetime_revenue']):>12} "
                  f"{format_currency(row['ttm_revenue_last']):>12} "
                  f"{row['peak_ttm_share']*100:>9.1f}% "
                  f"{row['months_since']:>10.0f}")

        print(f"\n{'SUBTOTAL':<50} {format_currency(total_rev):>12} "
              f"(Avg: {format_currency(avg_rev)})")

    print("\n" + "="*90)
    print(f"TOTAL HIGH VALUE: {len(hv)} customers, {format_currency(hv['lifetime_revenue'].sum())}")
    print("="*90 + "\n")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="List high value customers grouped by status"
    )
    parser.add_argument("input_file", help="Path to customer analysis Excel file")
    args = parser.parse_args()
    main(args.input_file)
