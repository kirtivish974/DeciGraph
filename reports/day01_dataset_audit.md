
# DecisionGraph — Day 1 Dataset Audit

## 1. Dataset Overview

- Company: Asteron Industrial Systems Pvt Ltd
- Financial Year: FY 2025-26
- Period: 1 April 2025 – 31 March 2026
- Currency: INR
- Dataset generation seed: 42
- Total CSV tables: 8
- Total CSV records: 135

## 2. Dataset Inventory

| Table | Records |
|---|---:|
| suppliers | 5 |
| products | 10 |
| contracts | 5 |
| purchase_orders | 25 |
| invoices | 25 |
| payments | 25 |
| supplier_deliveries | 25 |
| quality_reports | 15 |
| **Total** | **135** |

## 3. Structural Validation

- All 8 CSV files have expected columns.
- No missing values detected.
- No exact duplicate rows detected.
- 32 relationship and calculation checks passed.
- Primary and foreign key relationships validated.

## 4. Document Validation

- Markdown documents: 12
- PDF documents: 5
- All Markdown documents contain extractable text.
- All PDFs have successfully extractable text.
- Total registered source files: 26.

## 5. CSV and Excel Consistency

All 8 CSV tables match their corresponding Excel workbook sheets.

Numeric formatting differences were identified during the initial audit and resolved by using numeric-aware comparison.

## 6. Evaluation Dataset

- Decision scenarios: 10
- Ground-truth anomalies: 12
- Registered document IDs: 17

Known anomaly categories include:
- Delivery delays and quantity shortages
- Invoice and purchase order mismatches
- Duplicate invoices
- Supplier quality failures
- Expired certifications
- Contract and email contradictions
- Supplier performance claim contradictions
- Unapproved price increases
- Partial payments

## 7. Known Audit Limitations

- Evidence ID registration does not prove that every referenced ID exists in the corresponding document content.
- Structural validation does not prove that all business rules are correct.
- Business anomalies require separate validation against their source evidence.
- The partial-payment anomaly needs a dedicated decision scenario.

## 8. Final Result

Day 1 dataset and source-file audit completed successfully.

The dataset is structurally ready for the next stage: document and tabular data ingestion.
