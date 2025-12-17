#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SRC_DIR="$SCRIPT_DIR/src"

# Activate venv if present
if [ -f "$SCRIPT_DIR/venv/bin/activate" ]; then
    source "$SCRIPT_DIR/venv/bin/activate"
fi

usage() {
    cat <<EOF
Usage: ./run.sh <command> [options]

Commands:
  all               Run full pipeline (prepare → analytics)
  prepare-customers Create/update customer master from revenue data
  prepare-revdata   Flatten revenue data and apply customer mappings
  run-analytics     Generate analytics workbook from prepared data
  gen-reports       Generate console or PDF reports
  gen-charts        Generate PNG visualizations

Run './run.sh <command> --help' for command-specific options.
EOF
    exit 1
}

show_all_help() {
    cat <<EOF
Usage: ./run.sh all --input <file> [options]

Runs the full pipeline: prepare-customers → prepare-revdata → run-analytics

Options:
  --input <file>     Raw revenue Excel file (required)
  --existing <file>  Existing customer master (auto-detects {outdir}/customers.xlsx)
  --outdir <dir>     Output directory (default: current dir)
  --merge            Auto-apply HIGH confidence consolidations
  --reports          Also generate PDF reports
  --charts           Also generate PNG charts

Output: {outdir}/customers.xlsx, revdata.xlsx, analytics.xlsx
EOF
}

cmd_all() {
    local input="" existing="" outdir="." merge_flag="" do_reports="" do_charts=""

    while [[ $# -gt 0 ]]; do
        case $1 in
            --input) [ -z "${2:-}" ] && echo "Error: --input requires a value" && echo "" && show_all_help && exit 1; input="$2"; shift 2 ;;
            --existing) [ -z "${2:-}" ] && echo "Error: --existing requires a value" && echo "" && show_all_help && exit 1; existing="$2"; shift 2 ;;
            --outdir) [ -z "${2:-}" ] && echo "Error: --outdir requires a value" && echo "" && show_all_help && exit 1; outdir="$2"; shift 2 ;;
            --merge) merge_flag="--merge"; shift ;;
            --reports) do_reports="yes"; shift ;;
            --charts) do_charts="yes"; shift ;;
            --help) show_all_help; exit 0 ;;
            *) echo "Unknown option: $1"; echo ""; show_all_help; exit 1 ;;
        esac
    done

    if [ -z "$input" ]; then
        echo "Error: --input is required"
        echo ""
        show_all_help
        exit 1
    fi

    mkdir -p "$outdir"

    # Auto-detect existing customer master if not specified
    if [ -z "$existing" ]; then
        auto_existing="$outdir/customers.xlsx"
        if [ -f "$auto_existing" ]; then
            existing="$auto_existing"
            echo "Auto-detected existing customer master: $existing"
        fi
    fi

    echo ""
    echo "=== FULL PIPELINE ==="
    echo "  Input: $input"
    echo "  Output: $outdir/"
    [ -n "$existing" ] && echo "  Existing master: $existing"
    echo ""

    # Step 1: prepare-customers
    echo "--- Step 1: Preparing customer master ---"
    cust_args=(--input "$input" --outdir "$outdir")
    [ -n "$existing" ] && cust_args+=(--existing "$existing")
    [ -n "$merge_flag" ] && cust_args+=($merge_flag)
    cmd_prepare_customers "${cust_args[@]}"
    echo ""

    # Step 2: prepare-revdata
    echo "--- Step 2: Preparing revenue data ---"
    cmd_prepare_data --input "$input" --customers "$outdir/customers.xlsx" --outdir "$outdir"
    echo ""

    # Step 3: run-analytics
    echo "--- Step 3: Generating analytics ---"
    cmd_analytics --revdata "$outdir/revdata.xlsx" --outdir "$outdir"
    echo ""

    # Optional: reports
    if [ -n "$do_reports" ]; then
        echo "--- Step 4: Generating reports ---"
        cmd_reports --analytics "$outdir/analytics.xlsx" --outdir "$outdir" --pdf --all
        echo ""
    fi

    # Optional: charts
    if [ -n "$do_charts" ]; then
        echo "--- Step 5: Generating charts ---"
        cmd_charts --analytics "$outdir/analytics.xlsx" --outdir "$outdir"
        echo ""
    fi

    echo "=== PIPELINE COMPLETE ==="
    echo "Output files in: $outdir/"
}

show_prepare_customers_help() {
    cat <<EOF
Usage: ./run.sh prepare-customers --input <file> [options]

Options:
  --input <file>     Raw revenue Excel file (required)
  --existing <file>  Existing customer master to update
  --merge            Auto-apply HIGH confidence consolidations
  --nomerge          Just flag duplicates (default)
  --outdir <dir>     Output directory (default: current dir)

Output: {outdir}/customers.xlsx
EOF
}

cmd_prepare_customers() {
    local input="" existing="" outdir="." merge_flag=""

    while [[ $# -gt 0 ]]; do
        case $1 in
            --input) [ -z "${2:-}" ] && echo "Error: --input requires a value" && echo "" && show_prepare_customers_help && exit 1; input="$2"; shift 2 ;;
            --existing) [ -z "${2:-}" ] && echo "Error: --existing requires a value" && echo "" && show_prepare_customers_help && exit 1; existing="$2"; shift 2 ;;
            --merge) merge_flag="--merge"; shift ;;
            --nomerge) merge_flag=""; shift ;;
            --outdir) [ -z "${2:-}" ] && echo "Error: --outdir requires a value" && echo "" && show_prepare_customers_help && exit 1; outdir="$2"; shift 2 ;;
            --help) show_prepare_customers_help; exit 0 ;;
            *) echo "Unknown option: $1"; echo ""; show_prepare_customers_help; exit 1 ;;
        esac
    done

    if [ -z "$input" ]; then
        echo "Error: --input is required"
        echo ""
        show_prepare_customers_help
        exit 1
    fi

    output="$outdir/customers.xlsx"

    mkdir -p "$outdir"

    args=(--input "$input" --output "$output")
    [ -n "$existing" ] && args+=(--master "$existing")
    [ -n "$merge_flag" ] && args+=($merge_flag)

    echo "Creating customer master..."
    echo "  Input: $input"
    [ -n "$existing" ] && echo "  Existing master: $existing"
    echo "  Output: $output"

    python "$SRC_DIR/prepare.py" "${args[@]}" --customers-only
}

show_prepare_revdata_help() {
    cat <<EOF
Usage: ./run.sh prepare-revdata --input <file> [options]

Options:
  --input <file>     Raw revenue Excel file (required)
  --customers <file> Customer master to apply (auto-detects if not specified)
  --outdir <dir>     Output directory (default: current dir)

Output: {outdir}/revdata.xlsx
EOF
}

cmd_prepare_data() {
    local input="" customers="" outdir="."

    while [[ $# -gt 0 ]]; do
        case $1 in
            --input) [ -z "${2:-}" ] && echo "Error: --input requires a value" && echo "" && show_prepare_revdata_help && exit 1; input="$2"; shift 2 ;;
            --customers) [ -z "${2:-}" ] && echo "Error: --customers requires a value" && echo "" && show_prepare_revdata_help && exit 1; customers="$2"; shift 2 ;;
            --outdir) [ -z "${2:-}" ] && echo "Error: --outdir requires a value" && echo "" && show_prepare_revdata_help && exit 1; outdir="$2"; shift 2 ;;
            --help) show_prepare_revdata_help; exit 0 ;;
            *) echo "Unknown option: $1"; echo ""; show_prepare_revdata_help; exit 1 ;;
        esac
    done

    if [ -z "$input" ]; then
        echo "Error: --input is required"
        echo ""
        show_prepare_revdata_help
        exit 1
    fi

    output="$outdir/revdata.xlsx"

    # Auto-detect customer master if not specified
    if [ -z "$customers" ]; then
        auto_customers="$outdir/customers.xlsx"
        if [ -f "$auto_customers" ]; then
            customers="$auto_customers"
            echo "Auto-detected customer master: $customers"
        fi
    fi

    mkdir -p "$outdir"

    args=(--input "$input" --output "$output")
    [ -n "$customers" ] && args+=(--master "$customers")

    echo "Preparing data..."
    echo "  Input: $input"
    [ -n "$customers" ] && echo "  Customer master: $customers"
    echo "  Output: $output"

    python "$SRC_DIR/prepare.py" "${args[@]}" --data-only
}

show_analytics_help() {
    cat <<EOF
Usage: ./run.sh run-analytics --revdata <file> [options]

Options:
  --revdata <file>   Prepared revenue data file (required)
  --outdir <dir>     Output directory (default: current dir)

Output: {outdir}/analytics.xlsx
EOF
}

cmd_analytics() {
    local revdata="" outdir="."

    while [[ $# -gt 0 ]]; do
        case $1 in
            --revdata) [ -z "${2:-}" ] && echo "Error: --revdata requires a value" && echo "" && show_analytics_help && exit 1; revdata="$2"; shift 2 ;;
            --outdir) [ -z "${2:-}" ] && echo "Error: --outdir requires a value" && echo "" && show_analytics_help && exit 1; outdir="$2"; shift 2 ;;
            --help) show_analytics_help; exit 0 ;;
            *) echo "Unknown option: $1"; echo ""; show_analytics_help; exit 1 ;;
        esac
    done

    if [ -z "$revdata" ]; then
        echo "Error: --revdata is required"
        echo ""
        show_analytics_help
        exit 1
    fi

    output="$outdir/analytics.xlsx"

    mkdir -p "$outdir"

    echo "Generating analytics..."
    echo "  Input: $revdata"
    echo "  Output: $output"

    python "$SRC_DIR/analytics.py" --input "$revdata" --output "$output"
}

show_reports_help() {
    cat <<EOF
Usage: ./run.sh gen-reports --analytics <file> [options]

Options:
  --analytics <file>  Analytics workbook (required)
  --outdir <dir>      Output directory for PDFs (default: current dir)
  --pdf               Generate PDF files instead of console output
  --all               Show all customers (high-value and other)
  --high-value        Show high-value customers only (default)
  --low-value         Show low-value customers only

Output (with --pdf): {outdir}/churn_report.pdf, {outdir}/segment_matrix.pdf
EOF
}

cmd_reports() {
    local analytics="" outdir="." pdf_flag="" filter="--high-value"

    while [[ $# -gt 0 ]]; do
        case $1 in
            --analytics) [ -z "${2:-}" ] && echo "Error: --analytics requires a value" && echo "" && show_reports_help && exit 1; analytics="$2"; shift 2 ;;
            --outdir) [ -z "${2:-}" ] && echo "Error: --outdir requires a value" && echo "" && show_reports_help && exit 1; outdir="$2"; shift 2 ;;
            --pdf) pdf_flag="yes"; shift ;;
            --all) filter="--all"; shift ;;
            --high-value) filter="--high-value"; shift ;;
            --low-value) filter="--low-value"; shift ;;
            --help) show_reports_help; exit 0 ;;
            *) echo "Unknown option: $1"; echo ""; show_reports_help; exit 1 ;;
        esac
    done

    if [ -z "$analytics" ]; then
        echo "Error: --analytics is required"
        echo ""
        show_reports_help
        exit 1
    fi

    mkdir -p "$outdir"

    if [ -n "$pdf_flag" ]; then
        echo "Generating PDF reports..."
        echo "  Input: $analytics"
        echo "  Output: $outdir/churn_report.pdf, $outdir/segment_matrix.pdf"

        python "$SRC_DIR/customer_churn_report.py" "$analytics" $filter --pdf "$outdir/churn_report.pdf"
        python "$SRC_DIR/customer_segment_matrix.py" "$analytics" --pdf "$outdir/segment_matrix.pdf"
    else
        echo "=== CHURN REPORT ==="
        python "$SRC_DIR/customer_churn_report.py" "$analytics" $filter
        echo ""
        echo "=== SEGMENT MATRIX ==="
        python "$SRC_DIR/customer_segment_matrix.py" "$analytics"
    fi
}

show_charts_help() {
    cat <<EOF
Usage: ./run.sh gen-charts --analytics <file> [options]

Options:
  --analytics <file>  Analytics workbook (required)
  --outdir <dir>      Output directory (default: ./charts)

Output: {outdir}/pareto_curve.png, concentration_trend.png, segment_heatmap.png
EOF
}

cmd_charts() {
    local analytics="" outdir="./charts"

    while [[ $# -gt 0 ]]; do
        case $1 in
            --analytics) [ -z "${2:-}" ] && echo "Error: --analytics requires a value" && echo "" && show_charts_help && exit 1; analytics="$2"; shift 2 ;;
            --outdir) [ -z "${2:-}" ] && echo "Error: --outdir requires a value" && echo "" && show_charts_help && exit 1; outdir="$2"; shift 2 ;;
            --help) show_charts_help; exit 0 ;;
            *) echo "Unknown option: $1"; echo ""; show_charts_help; exit 1 ;;
        esac
    done

    if [ -z "$analytics" ]; then
        echo "Error: --analytics is required"
        echo ""
        show_charts_help
        exit 1
    fi

    mkdir -p "$outdir"

    echo "Generating charts..."
    echo "  Input: $analytics"
    echo "  Output: $outdir/"

    python "$SRC_DIR/visualize.py" "$analytics" --output-dir "$outdir"
}

# Main dispatch
if [ $# -eq 0 ]; then
    usage
fi

command="$1"
shift

case "$command" in
    all) cmd_all "$@" ;;
    prepare-customers) cmd_prepare_customers "$@" ;;
    prepare-revdata) cmd_prepare_data "$@" ;;
    run-analytics) cmd_analytics "$@" ;;
    gen-reports) cmd_reports "$@" ;;
    gen-charts) cmd_charts "$@" ;;
    help|--help|-h) usage ;;
    *) echo "Unknown command: $command"; usage ;;
esac
