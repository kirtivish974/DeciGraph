# DecisionGraph Synthetic Procurement Dataset

Synthetic, deterministic procurement dataset for **Asteron Industrial Systems Pvt Ltd** (Pune, Maharashtra, India), covering **2025-04-01 to 2026-03-31** (INR, Indian GST).

## Quick Start
```bash
pip install pandas openpyxl reportlab
python scripts/generate_dataset.py
python scripts/validate_dataset.py
```

## Contents
- `data/raw/`: 8 relational CSV files (`suppliers`, `products`, `purchase_orders`, `supplier_deliveries`, `invoices`, `payments`, `contracts`, `quality_reports`)
- `data/processed/procurement_data.xlsx`: 8-sheet Excel workbook matching all CSVs
- `documents/`: 12 Markdown documents (contracts, emails, quality reports, management reports) and 5 text-selectable PDFs
- `evaluation/`: `decision_scenarios.json` (10 test questions) and `ground_truth_anomalies.json` (12 planted anomalies)
- `reports/validation_report.json`: automated validation output from `scripts/validate_dataset.py`
