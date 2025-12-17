#!/usr/bin/env python3
"""Customer churn report grouped by status (Active/Inactive/Churned)."""

import argparse
import sys
import pandas as pd

from config import get_high_value_thresholds, get_status_thresholds


def format_currency(val):
    """Format value as currency (millions)."""
    if pd.isna(val):
        return '-'
    if val >= 1_000_000:
        return f'${val/1_000_000:.1f}M'
    else:
        return f'${val/1_000:,.0f}K'


def generate_section(df, title, subtitle, data_start, data_end, active_max=6, inactive_max=18):
    """Generate report lines for a subset of customers."""
    lines = []
    if len(df) == 0:
        return lines

    status_order = {'Active': 0, 'Inactive': 1, 'Churned': 2}
    df = df.copy()
    df['status_sort'] = df['status'].map(status_order)
    df = df.sort_values(['status_sort', 'lifetime_revenue'], ascending=[True, False])

    lines.append("")
    lines.append("=" * 72)
    lines.append(title)
    if subtitle:
        lines.append(subtitle)
    lines.append(f"Data Range: {data_start.strftime('%b %Y')} - {data_end.strftime('%b %Y')}")
    lines.append("=" * 72)

    status_descriptions = {
        'Active': f'Last purchase within {active_max} months',
        'Inactive': f'Last purchase {active_max + 1}-{inactive_max} months ago',
        'Churned': f'Last purchase over {inactive_max} months ago',
    }

    for status in ['Active', 'Inactive', 'Churned']:
        status_df = df[df['status'] == status]

        if len(status_df) == 0:
            continue

        total_rev = status_df['lifetime_revenue'].sum()
        avg_rev = status_df['lifetime_revenue'].mean()

        lines.append("")
        lines.append("")
        lines.append(f"{status.upper()} ({len(status_df)} customers, {format_currency(total_rev)} total)")
        lines.append(status_descriptions[status])
        lines.append("-" * 72)
        lines.append(f"{'Customer':<40} {'Lifetime':>8} {'TTM':>8} {'Peak':>6} {'Months':>6}")
        lines.append("-" * 72)

        for _, row in status_df.iterrows():
            lines.append(f"{row['customer'][:40]:<40} "
                        f"{format_currency(row['lifetime_revenue']):>8} "
                        f"{format_currency(row['ttm_revenue_last']):>8} "
                        f"{row['peak_ttm_share']*100:>5.1f}% "
                        f"{row['months_since']:>6.0f}")

        lines.append("")
        lines.append(f"{'SUBTOTAL':<40} {format_currency(total_rev):>8} "
                    f"(Avg: {format_currency(avg_rev)})")

    lines.append("")
    lines.append("")
    lines.append("=" * 72)
    lines.append(f"TOTAL: {len(df)} customers, {format_currency(df['lifetime_revenue'].sum())}")
    lines.append("=" * 72)
    lines.append("")

    return lines


def write_pdf(lines, output_path):
    """Write report lines to PDF."""
    from datetime import datetime, timezone
    from fpdf import FPDF
    from fpdf.enums import XPos, YPos

    pdf = FPDF()
    # Fixed timestamp for reproducible output
    pdf.creation_date = datetime(2020, 1, 1, tzinfo=timezone.utc)
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.set_font('Courier', size=9)

    for line in lines:
        # Handle special characters
        safe_line = line.encode('latin-1', errors='replace').decode('latin-1')
        pdf.cell(0, 4, safe_line, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.output(output_path)
    print(f"Wrote: {output_path}", file=sys.stderr)


def main(input_file: str, filter_type: str, pdf_output: str = None):
    # Load thresholds from config
    ltv_min, share_min = get_high_value_thresholds()
    active_max, inactive_max = get_status_thresholds()

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
    df['months_since'] = ((reference_date - df['last_purchase_month']).dt.days / 30.44).round(0)

    # Calculate status using config thresholds
    def calc_status(months):
        if months <= active_max:
            return 'Active'
        elif months <= inactive_max:
            return 'Inactive'
        else:
            return 'Churned'

    df['status'] = df['months_since'].apply(calc_status)

    # Calculate is_high_value using config thresholds
    df['is_high_value'] = (df['lifetime_revenue'] >= ltv_min) | (df['peak_ttm_share'] >= share_min)

    # Filter based on filter_type
    high_value_df = df[df['is_high_value']]
    low_value_df = df[~df['is_high_value']]

    # Build subtitle with actual thresholds
    ltv_fmt = f"${ltv_min/1_000_000:.0f}M" if ltv_min >= 1_000_000 else f"${ltv_min/1_000:.0f}K"
    share_fmt = f"{share_min*100:.0f}%"
    hv_subtitle = f"Lifetime >= {ltv_fmt} OR Peak TTM Share >= {share_fmt}"
    other_subtitle = f"Lifetime < {ltv_fmt} AND Peak TTM Share < {share_fmt}"

    lines = []
    if filter_type == 'high-value':
        lines.extend(generate_section(
            high_value_df,
            "HIGH VALUE CUSTOMERS",
            hv_subtitle,
            data_start, data_end, active_max, inactive_max
        ))
    elif filter_type == 'low-value':
        lines.extend(generate_section(
            low_value_df,
            "OTHER CUSTOMERS",
            other_subtitle,
            data_start, data_end, active_max, inactive_max
        ))
    elif filter_type == 'all':
        lines.extend(generate_section(
            high_value_df,
            "HIGH VALUE CUSTOMERS",
            hv_subtitle,
            data_start, data_end, active_max, inactive_max
        ))
        lines.extend(generate_section(
            low_value_df,
            "OTHER CUSTOMERS",
            other_subtitle,
            data_start, data_end, active_max, inactive_max
        ))

    # Output
    if pdf_output:
        write_pdf(lines, pdf_output)
    else:
        for line in lines:
            print(line)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Customer churn report grouped by status"
    )
    parser.add_argument("input_file", help="Path to customer analytics Excel file")
    parser.add_argument("--pdf", metavar="FILE", help="Output to PDF file instead of console")
    group = parser.add_mutually_exclusive_group()
    group.add_argument(
        "--high-value", action="store_true", default=True,
        help="Show high-value customers only (default)"
    )
    group.add_argument(
        "--low-value", action="store_true",
        help="Show non-high-value customers only"
    )
    group.add_argument(
        "--all", action="store_true",
        help="Show all customers (high-value and other in separate sections)"
    )
    args = parser.parse_args()

    if args.all:
        filter_type = 'all'
    elif args.low_value:
        filter_type = 'low-value'
    else:
        filter_type = 'high-value'

    main(args.input_file, filter_type, args.pdf)
