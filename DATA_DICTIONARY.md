# Data Dictionary — DecisionGraph Dataset

| Table | Primary Key | Row Count | Foreign Keys |
|---|---|---|---|
| `suppliers.csv` | `supplier_id` | 5 | — |
| `products.csv` | `product_id` | 10 | — |
| `contracts.csv` | `contract_id` | 5 | `supplier_id` -> `suppliers` |
| `purchase_orders.csv` | `po_id` | 25 | `supplier_id` -> `suppliers`, `product_id` -> `products` |
| `supplier_deliveries.csv` | `delivery_id` | 25 | `po_id` -> `purchase_orders`, `supplier_id` -> `suppliers` |
| `invoices.csv` | `invoice_id` | 25 | `po_id` -> `purchase_orders`, `supplier_id` -> `suppliers` |
| `payments.csv` | `payment_id` | 25 | `invoice_id` -> `invoices`, `supplier_id` -> `suppliers` |
| `quality_reports.csv` | `qr_id` | 15 | `supplier_id` -> `suppliers`, `po_id` -> `purchase_orders`, `product_id` -> `products` |
