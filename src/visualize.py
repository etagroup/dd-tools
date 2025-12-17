#!/usr/bin/env python3
"""Generate visualizations from customer analytics data."""

import argparse
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd


def plot_pareto_curve(df: pd.DataFrame, output_path: Path) -> None:
    """
    Generate a Pareto curve showing revenue concentration.

    X-axis: Cumulative % of customers (sorted by revenue descending)
    Y-axis: Cumulative % of revenue
    """
    df_sorted = df.sort_values('lifetime_revenue', ascending=False).copy()
    total_revenue = df_sorted['lifetime_revenue'].sum()
    total_customers = len(df_sorted)

    df_sorted['cumulative_revenue'] = df_sorted['lifetime_revenue'].cumsum()
    df_sorted['cumulative_revenue_pct'] = df_sorted['cumulative_revenue'] / total_revenue * 100
    df_sorted['customer_rank'] = range(1, total_customers + 1)
    df_sorted['cumulative_customer_pct'] = df_sorted['customer_rank'] / total_customers * 100

    fig, ax = plt.subplots(figsize=(10, 7))

    # Plot the Pareto curve
    ax.plot(
        df_sorted['cumulative_customer_pct'],
        df_sorted['cumulative_revenue_pct'],
        'b-', linewidth=2, label='Actual'
    )

    # Plot the diagonal (perfect equality line)
    ax.plot([0, 100], [0, 100], 'k--', alpha=0.5, label='Perfect equality')

    # Find key points (80/20, 50%, etc.)
    pct_for_80 = df_sorted[df_sorted['cumulative_revenue_pct'] >= 80]['cumulative_customer_pct'].iloc[0]
    pct_for_50 = df_sorted[df_sorted['cumulative_revenue_pct'] >= 50]['cumulative_customer_pct'].iloc[0]

    # Add reference lines
    ax.axhline(y=80, color='r', linestyle=':', alpha=0.5)
    ax.axvline(x=pct_for_80, color='r', linestyle=':', alpha=0.5)
    ax.axhline(y=50, color='orange', linestyle=':', alpha=0.5)
    ax.axvline(x=pct_for_50, color='orange', linestyle=':', alpha=0.5)

    # Annotate key points
    ax.annotate(
        f'{pct_for_80:.0f}% of customers = 80% of revenue',
        xy=(pct_for_80, 80), xytext=(pct_for_80 + 10, 70),
        fontsize=10, color='r',
        arrowprops=dict(arrowstyle='->', color='r', alpha=0.7)
    )
    ax.annotate(
        f'{pct_for_50:.0f}% of customers = 50% of revenue',
        xy=(pct_for_50, 50), xytext=(pct_for_50 + 10, 40),
        fontsize=10, color='orange',
        arrowprops=dict(arrowstyle='->', color='orange', alpha=0.7)
    )

    ax.set_xlabel('Cumulative % of Customers', fontsize=12)
    ax.set_ylabel('Cumulative % of Revenue', fontsize=12)
    ax.set_title('Revenue Concentration (Pareto Curve)', fontsize=14)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.legend(loc='lower right')
    ax.grid(True, alpha=0.3)

    plt.tight_layout()
    plt.savefig(output_path, dpi=150, bbox_inches='tight')
    plt.close()
    print(f"Saved: {output_path}")


def plot_concentration_trend(df: pd.DataFrame, output_path: Path) -> None:
    """
    Plot top-N customer concentration over time (TTM basis).

    Shows how dependent the business is on top customers over time.
    """
    df = df.copy()
    df['date'] = pd.to_datetime(df['date'])

    # Filter to valid TTM periods (after first 12 months of data)
    df = df.dropna(subset=['top1_share_trailing_12m'])

    fig, ax = plt.subplots(figsize=(12, 7))

    # Plot top 1, 5, 10 concentration
    ax.plot(df['date'], df['top1_share_trailing_12m'] * 100,
            'r-', linewidth=2, label='Top 1 customer', marker='o', markersize=3)
    ax.plot(df['date'], df['top5_share_trailing_12m'] * 100,
            'orange', linewidth=2, label='Top 5 customers', marker='s', markersize=3)
    ax.plot(df['date'], df['top10_share_trailing_12m'] * 100,
            'g-', linewidth=2, label='Top 10 customers', marker='^', markersize=3)

    # Add reference lines
    ax.axhline(y=50, color='gray', linestyle='--', alpha=0.5, label='50% threshold')
    ax.axhline(y=25, color='gray', linestyle=':', alpha=0.5, label='25% threshold')

    ax.set_xlabel('Date', fontsize=12)
    ax.set_ylabel('Share of TTM Revenue (%)', fontsize=12)
    ax.set_title('Customer Concentration Over Time (Trailing 12 Months)', fontsize=14)
    ax.legend(loc='upper right')
    ax.grid(True, alpha=0.3)

    # Format x-axis dates
    fig.autofmt_xdate()

    # Set y-axis limits with some padding
    ymax = max(df['top10_share_trailing_12m'].max() * 100, 60)
    ax.set_ylim(0, min(ymax + 10, 100))

    plt.tight_layout()
    plt.savefig(output_path, dpi=150, bbox_inches='tight')
    plt.close()
    print(f"Saved: {output_path}")


def plot_segment_heatmap(df: pd.DataFrame, output_path: Path, data_end: pd.Timestamp) -> None:
    """
    Create a heatmap of customer segments (Status x Value).

    Color intensity represents revenue.
    """
    df = df.copy()
    df['last_purchase_month'] = pd.to_datetime(df['last_purchase_month'])
    df['months_since'] = ((data_end - df['last_purchase_month']).dt.days / 30.44).round(0)

    df['status'] = df['months_since'].apply(
        lambda x: 'Active' if x <= 6 else ('Inactive' if x <= 18 else 'Churned')
    )
    df['is_high_value'] = (df['lifetime_revenue'] >= 1_000_000) | (df['peak_ttm_share'] >= 0.02)
    df['value_segment'] = df['is_high_value'].map({True: 'High Value', False: 'Other'})

    # Create pivot tables for counts and revenue
    counts = pd.crosstab(df['value_segment'], df['status'])
    revenue = pd.crosstab(df['value_segment'], df['status'],
                          values=df['lifetime_revenue'], aggfunc='sum')

    # Reorder
    status_order = ['Active', 'Inactive', 'Churned']
    value_order = ['High Value', 'Other']
    counts = counts.reindex(index=value_order, columns=status_order, fill_value=0)
    revenue = revenue.reindex(index=value_order, columns=status_order, fill_value=0)

    fig, ax = plt.subplots(figsize=(10, 6))

    # Create heatmap based on revenue
    im = ax.imshow(revenue.values / 1_000_000, cmap='Blues', aspect='auto')

    # Add colorbar
    cbar = ax.figure.colorbar(im, ax=ax)
    cbar.ax.set_ylabel('Lifetime Revenue ($M)', rotation=-90, va='bottom', fontsize=11)

    # Set ticks and labels
    ax.set_xticks(range(len(status_order)))
    ax.set_yticks(range(len(value_order)))
    ax.set_xticklabels(status_order, fontsize=12)
    ax.set_yticklabels(value_order, fontsize=12)

    # Add text annotations (count and revenue)
    for i in range(len(value_order)):
        for j in range(len(status_order)):
            count = counts.iloc[i, j]
            rev = revenue.iloc[i, j] / 1_000_000
            text = f"{count}\n${rev:.1f}M"
            # Use white text on dark cells
            text_color = 'white' if rev > revenue.values.max() / 2_000_000 else 'black'
            ax.text(j, i, text, ha='center', va='center', fontsize=12, color=text_color)

    ax.set_title('Customer Segment Matrix\n(Count and Lifetime Revenue)', fontsize=14)
    ax.set_xlabel('Status', fontsize=12)
    ax.set_ylabel('Value Segment', fontsize=12)

    plt.tight_layout()
    plt.savefig(output_path, dpi=150, bbox_inches='tight')
    plt.close()
    print(f"Saved: {output_path}")


def main():
    parser = argparse.ArgumentParser(description="Generate visualizations from customer analytics")
    parser.add_argument("input_file", help="Path to customer analytics Excel file")
    parser.add_argument("--output-dir", "-o", default=".", help="Output directory for charts")
    args = parser.parse_args()

    input_path = Path(args.input_file)
    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    # Read data
    print(f"Loading: {input_path}")
    customer_summary = pd.read_excel(input_path, sheet_name="Customer_Summary")
    rolling_conc = pd.read_excel(input_path, sheet_name="Rolling_Concentration")

    # Read metadata for date range
    metadata = pd.read_excel(input_path, sheet_name="Metadata")
    metadata_dict = dict(zip(metadata["property"], metadata["value"]))
    data_end = pd.to_datetime(metadata_dict["data_end_date"])

    print(f"  Customers: {len(customer_summary)}")
    print(f"  Data end: {data_end.strftime('%b %Y')}")
    print()

    # Generate charts
    plot_pareto_curve(customer_summary, output_dir / "pareto_curve.png")
    plot_concentration_trend(rolling_conc, output_dir / "concentration_trend.png")
    plot_segment_heatmap(customer_summary, output_dir / "segment_heatmap.png", data_end)

    print("\nDone!")


if __name__ == "__main__":
    main()
