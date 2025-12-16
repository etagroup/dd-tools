#!/usr/bin/env bash
set -euo pipefail

python src/generate_repeat_customer_analysis.py \
  --input "tmp/4.0.5 Financial - Month-wise Customer Revenue.xlsx" \
  --output "tmp/customers.xlsx"
