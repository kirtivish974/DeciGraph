
# DecisionGraph

A company uploads its internal business data. DecisionGraph connects structured records and documents, investigates evidence, identifies risks and contradictions, compares alternatives, and produces explainable recommendations. A human remains responsible for the final decision.

## Current Project: Synthetic Procurement Dataset

Synthetic, deterministic procurement dataset for **Asteron Industrial Systems Pvt Ltd** (Pune, Maharashtra, India), covering **2025-04-01 to 2026-03-31** (INR, Indian GST).

## Quick Start

```bash
pip install -r requirements.txt
python scripts/generate_dataset.py
python scripts/validate_dataset.py
```

## Project Contents

- `data/raw/`: 8 relational CSV files.
- `data/processed/`: Excel workbook containing 8 sheets matching the CSV files.
- `documents/`: Contracts, emails, quality reports, management reports, and PDFs.
- `evaluation/`: 10 decision scenarios and 12 ground-truth anomalies.
- `reports/`: Dataset and document validation reports.
- `scripts/`: Dataset generation, auditing, and validation scripts.