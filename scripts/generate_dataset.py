"""
DecisionGraph Synthetic Dataset Generator
=========================================
Generates a deterministic, internally consistent procurement dataset for
Asteron Industrial Systems Pvt Ltd (Pune, Maharashtra, India) covering
2025-04-01 to 2026-03-31.

Part 1: Configuration, Directory Setup, Master Data (Suppliers, Products,
Contracts), Purchase Orders, and Part 1 Validation Assertions.
"""

from __future__ import annotations

import json
import os
import random
import subprocess
import sys
import zipfile
from datetime import date, timedelta
from pathlib import Path

import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas

# ---------------------------------------------------------------------------
# 1. Deterministic Configuration & Paths
# ---------------------------------------------------------------------------
SEED = 42
random.seed(SEED)

COMPANY_NAME = "Asteron Industrial Systems Pvt Ltd"
COMPANY_LOCATION = "Pune, Maharashtra, India"
DATASET_START = date(2025, 4, 1)
DATASET_END = date(2026, 3, 31)

BASE_DIR = Path(__file__).resolve().parent.parent
DIRS = {
    "raw": BASE_DIR / "data" / "raw",
    "processed": BASE_DIR / "data" / "processed",
    "contracts_md": BASE_DIR / "documents" / "contracts",
    "emails_md": BASE_DIR / "documents" / "emails",
    "quality_md": BASE_DIR / "documents" / "quality",
    "management_md": BASE_DIR / "documents" / "management",
    "pdf": BASE_DIR / "documents" / "pdf",
    "evaluation": BASE_DIR / "evaluation",
    "scripts": BASE_DIR / "scripts",
    "reports": BASE_DIR / "reports",
}


def ensure_directories() -> None:
    """Create the required project directory tree."""
    for path in DIRS.values():
        path.mkdir(parents=True, exist_ok=True)


# ---------------------------------------------------------------------------
# 2. Structured Data Generators (Part 1)
# ---------------------------------------------------------------------------
def gen_suppliers() -> pd.DataFrame:
    """Generate 5 fictional Indian industrial electrical suppliers."""
    rows = [
        {
            "supplier_id": "SUP001",
            "supplier_name": "Kaveri ElectroMech Works Pvt Ltd",
            "contact_person": "Rajesh Kulkarni",
            "email": "rajesh.kulkarni@kaverielectromech.in",
            "phone": "+91-9823011401",
            "address": "Plot 14, Bhosari MIDC Industrial Area, Sector 7",
            "city": "Pune",
            "state": "Maharashtra",
            "gstin": "27AABCK4128M1Z5",
            "registration_date": "2021-06-15",
            "supplier_category": "Switchgear & Busbars",
            "risk_rating": "High",
            "status": "Active",
        },
        {
            "supplier_id": "SUP002",
            "supplier_name": "Vindhya Precision Controls Ltd",
            "contact_person": "Meera Deshmukh",
            "email": "meera.deshmukh@vindhyacontrols.in",
            "phone": "+91-9820455812",
            "address": "Unit 8, Mahape Electronic Zone, TTC Industrial Area",
            "city": "Navi Mumbai",
            "state": "Maharashtra",
            "gstin": "27AAACV8831N1Z2",
            "registration_date": "2019-11-04",
            "supplier_category": "Automation & Relays",
            "risk_rating": "Low",
            "status": "Active",
        },
        {
            "supplier_id": "SUP003",
            "supplier_name": "Sahyadri Power Components Pvt Ltd",
            "contact_person": "Anil Joshi",
            "email": "anil.joshi@sahyadripower.in",
            "phone": "+91-9881234907",
            "address": "Shed 22, Chakan MIDC Phase II",
            "city": "Pune",
            "state": "Maharashtra",
            "gstin": "27AADCS3290P1Z8",
            "registration_date": "2020-08-22",
            "supplier_category": "Cables & Transformers",
            "risk_rating": "Low",
            "status": "Active",
        },
        {
            "supplier_id": "SUP004",
            "supplier_name": "Deccan Industrial Insulators LLP",
            "contact_person": "Prakash Nair",
            "email": "prakash.nair@deccaninsulators.in",
            "phone": "+91-9765402198",
            "address": "Survey 45/2, Jeedimetla Industrial Estate",
            "city": "Hyderabad",
            "state": "Telangana",
            "gstin": "36AABCD6714R1Z9",
            "registration_date": "2022-02-10",
            "supplier_category": "Insulators & Enclosures",
            "risk_rating": "Medium",
            "status": "Active",
        },
        {
            "supplier_id": "SUP005",
            "supplier_name": "Suryodaya VoltEdge Enterprises",
            "contact_person": "Vikram Patwardhan",
            "email": "vikram.p@suryodayavoltedge.in",
            "phone": "+91-9890177364",
            "address": "Gala 5, Satpur MIDC Industrial Area",
            "city": "Nashik",
            "state": "Maharashtra",
            "gstin": "27AAECS9502K1Z1",
            "registration_date": "2023-01-19",
            "supplier_category": "Circuit Protection & Meters",
            "risk_rating": "High",
            "status": "Under Review",
        },
    ]
    return pd.DataFrame(rows)


def gen_products() -> pd.DataFrame:
    """Generate 10 industrial electrical component SKUs."""
    rows = [
        {
            "product_id": "PRD001",
            "product_name": "Copper Busbar 400A Electrolytic",
            "category": "Busbars & Conductors",
            "unit_of_measure": "Metres",
            "hsn_code": "74071020",
            "standard_unit_price": 2400.0,
            "gst_rate": 0.18,
            "description": "High-conductivity electrolytic copper busbar for LT panels.",
        },
        {
            "product_id": "PRD002",
            "product_name": "DIN Rail Terminal Block 35sqmm",
            "category": "Connectors & Terminals",
            "unit_of_measure": "Nos",
            "hsn_code": "85369090",
            "standard_unit_price": 120.0,
            "gst_rate": 0.18,
            "description": "Flame-retardant polyamide DIN rail screw terminal block.",
        },
        {
            "product_id": "PRD003",
            "product_name": "Molded Case Circuit Breaker 250A 3P",
            "category": "Circuit Protection",
            "unit_of_measure": "Nos",
            "hsn_code": "85362020",
            "standard_unit_price": 4500.0,
            "gst_rate": 0.18,
            "description": "3-pole thermal-magnetic MCCB rated 250A, 36kA breaking capacity.",
        },
        {
            "product_id": "PRD004",
            "product_name": "XLPE Armoured Control Cable 12C x 2.5sqmm",
            "category": "Cables & Wiring",
            "unit_of_measure": "Metres",
            "hsn_code": "85444999",
            "standard_unit_price": 1900.0,
            "gst_rate": 0.18,
            "description": "1.1kV grade XLPE insulated galvanized steel wire armoured cable.",
        },
        {
            "product_id": "PRD005",
            "product_name": "Epoxy Cast Resin Busbar Insulator 11kV",
            "category": "Insulators",
            "unit_of_measure": "Nos",
            "hsn_code": "85469010",
            "standard_unit_price": 850.0,
            "gst_rate": 0.12,
            "description": "Indoor cycloaliphatic epoxy post insulator for 11kV switchgear.",
        },
        {
            "product_id": "PRD006",
            "product_name": "Digital Multifunction Panel Meter",
            "category": "Metering & Monitoring",
            "unit_of_measure": "Nos",
            "hsn_code": "90303390",
            "standard_unit_price": 3200.0,
            "gst_rate": 0.18,
            "description": "True-RMS 3-phase digital panel meter with RS485 Modbus output.",
        },
        {
            "product_id": "PRD007",
            "product_name": "Industrial Surge Protection Device Type 2",
            "category": "Circuit Protection",
            "unit_of_measure": "Nos",
            "hsn_code": "85354020",
            "standard_unit_price": 2800.0,
            "gst_rate": 0.18,
            "description": "3-phase+N Type 2 SPD with replaceable varistor cartridges.",
        },
        {
            "product_id": "PRD008",
            "product_name": "Solid State Industrial Relay 40A",
            "category": "Automation & Relays",
            "unit_of_measure": "Nos",
            "hsn_code": "85364900",
            "standard_unit_price": 1650.0,
            "gst_rate": 0.18,
            "description": "Zero-cross switching SSR with integrated LED status indicator.",
        },
        {
            "product_id": "PRD009",
            "product_name": "Powder-Coated IP65 Control Enclosure",
            "category": "Enclosures",
            "unit_of_measure": "Nos",
            "hsn_code": "85381010",
            "standard_unit_price": 6200.0,
            "gst_rate": 0.18,
            "description": "CRCA sheet steel wall-mount enclosure 600x400x250mm, RAL 7035.",
        },
        {
            "product_id": "PRD010",
            "product_name": "Porcelain Pin Insulator 11kV Outdoor",
            "category": "Insulators",
            "unit_of_measure": "Nos",
            "hsn_code": "85462011",
            "standard_unit_price": 540.0,
            "gst_rate": 0.12,
            "description": "Wet-process glazed porcelain outdoor pin insulator.",
        },
    ]
    return pd.DataFrame(rows)


def gen_contracts() -> pd.DataFrame:
    """Generate 5 supplier master contracts (one per supplier)."""
    rows = [
        {
            "contract_id": "CON-2025-001",
            "supplier_id": "SUP001",
            "contract_start_date": "2025-04-01",
            "contract_end_date": "2026-03-31",
            "contract_value": 1800000.0,
            "payment_terms": "Net 30",
            "delivery_terms": "DAP Pune Plant; max 14-day lead time",
            "penalty_clause": "1.0% of PO value per week of delay, capped at 5%",
            "auto_renewal": False,
            "status": "Active",
        },
        {
            "contract_id": "CON-2025-002",
            "supplier_id": "SUP002",
            "contract_start_date": "2025-04-01",
            "contract_end_date": "2026-03-31",
            "contract_value": 2400000.0,
            "payment_terms": "Net 45",
            "delivery_terms": "DAP Pune Plant; max 10-day lead time",
            "penalty_clause": "0.5% per week of delay beyond grace period of 2 days",
            "auto_renewal": True,
            "status": "Active",
        },
        {
            "contract_id": "CON-2025-003",
            "supplier_id": "SUP003",
            "contract_start_date": "2025-04-01",
            "contract_end_date": "2026-03-31",
            "contract_value": 2000000.0,
            "payment_terms": "Net 30",
            "delivery_terms": "DAP Pune Plant; max 12-day lead time",
            "penalty_clause": "0.75% per week of delay; fixed unit prices for 12 months",
            "auto_renewal": True,
            "status": "Active",
        },
        {
            "contract_id": "CON-2025-004",
            "supplier_id": "SUP004",
            "contract_start_date": "2025-04-01",
            "contract_end_date": "2026-03-31",
            "contract_value": 1500000.0,
            "payment_terms": "Net 30",
            "delivery_terms": "DAP Pune Plant; valid ISO 9001 mandatory throughout term",
            "penalty_clause": "Immediate PO hold if compliance certificates lapse",
            "auto_renewal": False,
            "status": "Active",
        },
        {
            "contract_id": "CON-2025-005",
            "supplier_id": "SUP005",
            "contract_start_date": "2025-04-01",
            "contract_end_date": "2026-03-31",
            "contract_value": 1600000.0,
            "payment_terms": "Net 30",
            "delivery_terms": "DAP Pune Plant; max 14-day lead time",
            "penalty_clause": "Unit-price increases >5% prohibited without written buyer sign-off",
            "auto_renewal": False,
            "status": "Active",
        },
    ]
    return pd.DataFrame(rows)


def gen_purchase_orders(products_df: pd.DataFrame) -> pd.DataFrame:
    """
    Generate 25 purchase orders across SUP001-SUP005.

    Includes exact records required by the anomaly plan:
    - PO-2025-003 (SUP001): target for DEL-003 18-day delay (ANM-01)
    - PO-2025-004 (SUP003, PRD004): contracted unit_price 1850.0 (ANM-08)
    - PO-2025-008 (SUP001): target for DEL-008 12-day delay + partial (ANM-02)
    - PO-2025-015 (SUP003): grand_total = 236000.0 for partial payment (ANM-12)
    - PO-2025-018 (SUP005, PRD006): target for overbilled INV-2025-018 & QR-010 (ANM-03, ANM-05)
    - PO-2025-019 & PO-2025-022 (SUP005): target for duplicate invoice number (ANM-04)
    - PO-2025-021 (SUP001, PRD003): target for QR-012 quality failure (ANM-06)
    - PO-2025-023 (SUP005, PRD007): unapproved price jump to 3400.0 vs 2800.0 (ANM-11)
    """
    prod_lookup = products_df.set_index("product_id").to_dict("index")

    # Explicit deterministic PO schedule: (po_id, supplier_id, product_id, order_date, lead_days, qty, unit_price_override, status, approved_by)
    po_specs = [
        ("PO-2025-001", "SUP001", "PRD001", "2025-04-08", 12, 50, 2180.0, "Completed", "S. Gokhale (Procurement Mgr)"),
        ("PO-2025-002", "SUP002", "PRD008", "2025-04-15", 10, 80, 1820.0, "Completed", "S. Gokhale (Procurement Mgr)"),
        ("PO-2025-003", "SUP001", "PRD003", "2025-05-02", 14, 30, 4150.0, "Completed", "S. Gokhale (Procurement Mgr)"),
        ("PO-2025-004", "SUP003", "PRD004", "2025-05-12", 12, 100, 1850.0, "Completed", "A. Kulkarni (Sr. Buyer)"),
        ("PO-2025-005", "SUP004", "PRD005", "2025-05-20", 12, 120, 840.0, "Completed", "A. Kulkarni (Sr. Buyer)"),
        ("PO-2025-006", "SUP002", "PRD002", "2025-06-05", 10, 500, 135.0, "Completed", "S. Gokhale (Procurement Mgr)"),
        ("PO-2025-007", "SUP003", "PRD001", "2025-06-18", 12, 60, 2380.0, "Completed", "A. Kulkarni (Sr. Buyer)"),
        ("PO-2025-008", "SUP001", "PRD002", "2025-07-03", 12, 400, 108.0, "Completed", "S. Gokhale (Procurement Mgr)"),
        ("PO-2025-009", "SUP004", "PRD009", "2025-07-22", 14, 25, 6100.0, "Completed", "A. Kulkarni (Sr. Buyer)"),
        ("PO-2025-010", "SUP005", "PRD007", "2025-08-06", 14, 40, 2750.0, "Completed", "S. Gokhale (Procurement Mgr)"),
        ("PO-2025-011", "SUP002", "PRD006", "2025-08-19", 10, 45, 3450.0, "Completed", "S. Gokhale (Procurement Mgr)"),
        ("PO-2025-012", "SUP003", "PRD003", "2025-09-04", 12, 35, 4450.0, "Completed", "A. Kulkarni (Sr. Buyer)"),
        ("PO-2025-013", "SUP004", "PRD010", "2025-10-10", 12, 200, 535.0, "Completed", "A. Kulkarni (Sr. Buyer)"),
        ("PO-2025-014", "SUP001", "PRD004", "2025-10-21", 14, 80, 1740.0, "Completed", "S. Gokhale (Procurement Mgr)"),
        ("PO-2025-015", "SUP003", "PRD009", "2025-11-03", 12, 40, 5000.0, "Completed", "S. Gokhale (Procurement Mgr)"),
        ("PO-2025-016", "SUP002", "PRD008", "2025-11-18", 10, 100, 1800.0, "Completed", "S. Gokhale (Procurement Mgr)"),
        ("PO-2025-017", "SUP004", "PRD005", "2025-11-27", 12, 150, 845.0, "Completed", "A. Kulkarni (Sr. Buyer)"),
        ("PO-2025-018", "SUP005", "PRD006", "2025-12-08", 14, 50, 3150.0, "Completed", "S. Gokhale (Procurement Mgr)"),
        ("PO-2025-019", "SUP005", "PRD002", "2025-12-22", 12, 300, 118.0, "Completed", "A. Kulkarni (Sr. Buyer)"),
        ("PO-2025-020", "SUP002", "PRD003", "2026-01-09", 10, 30, 4850.0, "Completed", "S. Gokhale (Procurement Mgr)"),
        ("PO-2025-021", "SUP001", "PRD003", "2026-01-21", 14, 40, 4100.0, "Completed", "S. Gokhale (Procurement Mgr)"),
        ("PO-2025-022", "SUP005", "PRD001", "2026-02-04", 12, 45, 2350.0, "Completed", "A. Kulkarni (Sr. Buyer)"),
        ("PO-2025-023", "SUP005", "PRD007", "2026-02-18", 14, 35, 3400.0, "Completed", "Pending Buyer Sign-Off"),
        ("PO-2025-024", "SUP003", "PRD008", "2026-03-02", 12, 70, 1620.0, "Pending", "A. Kulkarni (Sr. Buyer)"),
        ("PO-2025-025", "SUP004", "PRD009", "2026-03-10", 12, 20, 6150.0, "Pending", "S. Gokhale (Procurement Mgr)"),
    ]

    rows = []
    for po_id, sup_id, prd_id, ord_str, lead_days, qty, unit_price, status, approved_by in po_specs:
        ord_dt = date.fromisoformat(ord_str)
        exp_dt = ord_dt + timedelta(days=lead_days)
        gst_rate = float(prod_lookup[prd_id]["gst_rate"])
        total_amount = round(qty * unit_price, 2)
        gst_amount = round(total_amount * gst_rate, 2)
        grand_total = round(total_amount + gst_amount, 2)
        rows.append(
            {
                "po_id": po_id,
                "supplier_id": sup_id,
                "product_id": prd_id,
                "order_date": ord_dt.isoformat(),
                "expected_delivery_date": exp_dt.isoformat(),
                "quantity": int(qty),
                "unit_price": float(unit_price),
                "total_amount": total_amount,
                "gst_amount": gst_amount,
                "grand_total": grand_total,
                "status": status,
                "approved_by": approved_by,
            }
        )
    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# 2B. Structured Data Generators (Part 2: Deliveries, Invoices, Payments)
# ---------------------------------------------------------------------------
def gen_deliveries(pos_df: pd.DataFrame) -> pd.DataFrame:
    """Generate 25 clean delivery records (1:1 with POs) before anomaly injection."""
    rows = []
    for idx, po in pos_df.iterrows():
        del_id = f"DEL-{idx + 1:03d}"
        exp_dt = date.fromisoformat(po["expected_delivery_date"])
        # Deterministic slight early/on-time variation (-2 to 0 days) for clean baseline
        offset = -(idx % 3)
        act_dt = max(date.fromisoformat(po["order_date"]) + timedelta(days=2), exp_dt + timedelta(days=offset))
        rows.append(
            {
                "delivery_id": del_id,
                "po_id": po["po_id"],
                "supplier_id": po["supplier_id"],
                "expected_delivery_date": po["expected_delivery_date"],
                "actual_delivery_date": act_dt.isoformat(),
                "quantity_ordered": int(po["quantity"]),
                "quantity_delivered": int(po["quantity"]),
                "delivery_status": "On-Time",
                "delay_days": 0,
                "remarks": "Delivered in full within agreed schedule.",
            }
        )
    return pd.DataFrame(rows)


def gen_invoices(
    pos_df: pd.DataFrame,
    products_df: pd.DataFrame,
    contracts_df: pd.DataFrame,
    deliveries_df: pd.DataFrame,
) -> pd.DataFrame:
    """Generate 25 clean invoice records (1:1 with POs) before anomaly injection."""
    prod_gst = products_df.set_index("product_id")["gst_rate"].to_dict()
    terms_map = contracts_df.set_index("supplier_id")["payment_terms"].to_dict()
    del_date_map = deliveries_df.set_index("po_id")["actual_delivery_date"].to_dict()

    prefix_map = {
        "SUP001": "KEW",
        "SUP002": "VPC",
        "SUP003": "SPC",
        "SUP004": "DII",
        "SUP005": "SVE",
    }

    rows = []
    for idx, po in pos_df.iterrows():
        seq = idx + 1
        inv_id = f"INV-2025-{seq:03d}"
        sup_id = po["supplier_id"]
        inv_num = f"{prefix_map[sup_id]}/25-26/{seq:04d}"

        inv_dt = date.fromisoformat(del_date_map[po["po_id"]])
        net_days = 45 if "45" in terms_map[sup_id] else 30
        due_dt = inv_dt + timedelta(days=net_days)

        taxable = round(float(po["total_amount"]), 2)
        gst_rate = float(prod_gst[po["product_id"]])
        gst_amt = round(taxable * gst_rate, 2)
        total_amt = round(taxable + gst_amt, 2)

        pay_status = "Pending" if po["status"] == "Pending" else "Paid"
        remarks = "Awaiting payment due date." if pay_status == "Pending" else "Verified against PO and GRN."

        rows.append(
            {
                "invoice_id": inv_id,
                "invoice_number": inv_num,
                "po_id": po["po_id"],
                "supplier_id": sup_id,
                "invoice_date": inv_dt.isoformat(),
                "taxable_amount": taxable,
                "gst_rate": gst_rate,
                "gst_amount": gst_amt,
                "total_amount": total_amt,
                "due_date": due_dt.isoformat(),
                "payment_status": pay_status,
                "remarks": remarks,
            }
        )
    return pd.DataFrame(rows)


def gen_payments(invoices_df: pd.DataFrame) -> pd.DataFrame:
    """Generate 25 clean payment records (1:1 with invoices) before anomaly injection."""
    modes = ["NEFT", "RTGS", "Cheque"]
    rows = []
    for idx, inv in invoices_df.iterrows():
        seq = idx + 1
        pay_id = f"PAY-{seq:03d}"
        inv_dt = date.fromisoformat(inv["invoice_date"])

        if inv["payment_status"] == "Pending":
            pay_dt = min(DATASET_END, inv_dt + timedelta(days=5))
            pay_amt = 0.0
            status = "Pending"
        else:
            pay_dt = min(DATASET_END, inv_dt + timedelta(days=12 + (idx % 10)))
            pay_amt = round(float(inv["total_amount"]), 2)
            status = "Completed"

        mode = modes[idx % len(modes)]
        ref = f"UTR2025{seq:06d}" if mode != "Cheque" else f"CHQ{710400 + seq}"

        rows.append(
            {
                "payment_id": pay_id,
                "invoice_id": inv["invoice_id"],
                "supplier_id": inv["supplier_id"],
                "payment_date": pay_dt.isoformat(),
                "payment_amount": pay_amt,
                "payment_mode": mode,
                "reference_number": ref,
                "status": status,
            }
        )
    return pd.DataFrame(rows)


def apply_part2_anomalies(
    deliveries_df: pd.DataFrame,
    invoices_df: pd.DataFrame,
    payments_df: pd.DataFrame,
) -> None:
    """
    Inject planned Part 2 anomalies in-place:
    - ANM-01: DEL-003 (SUP001, PO-2025-003) 18 days late
    - ANM-02: DEL-008 (SUP001, PO-2025-008) 12 days late + partial quantity
    - ANM-03: INV-2025-018 (SUP005, PO-2025-018) +24,000 taxable overbill, GST recalculated
    - ANM-04: INV-2025-019 & INV-2025-022 (SUP005) duplicate invoice_number SVE/25-26/0147
    - ANM-12: PAY-015 -> INV-2025-015 (SUP003) partial payment 198,000 vs 236,000
    """
    # ANM-01: DEL-003 18 days late
    m1 = deliveries_df["delivery_id"] == "DEL-003"
    exp_003 = date.fromisoformat(deliveries_df.loc[m1, "expected_delivery_date"].iloc[0])
    deliveries_df.loc[m1, "actual_delivery_date"] = (exp_003 + timedelta(days=18)).isoformat()
    deliveries_df.loc[m1, "delivery_status"] = "Late"
    deliveries_df.loc[m1, "delay_days"] = 18
    deliveries_df.loc[m1, "remarks"] = "Delayed by 18 days due to supplier production backlog."

    # ANM-02: DEL-008 12 days late + partial quantity
    m2 = deliveries_df["delivery_id"] == "DEL-008"
    exp_008 = date.fromisoformat(deliveries_df.loc[m2, "expected_delivery_date"].iloc[0])
    qty_ord_008 = int(deliveries_df.loc[m2, "quantity_ordered"].iloc[0])
    deliveries_df.loc[m2, "actual_delivery_date"] = (exp_008 + timedelta(days=12)).isoformat()
    deliveries_df.loc[m2, "quantity_delivered"] = qty_ord_008 - 60
    deliveries_df.loc[m2, "delivery_status"] = "Partial"
    deliveries_df.loc[m2, "delay_days"] = 12
    deliveries_df.loc[m2, "remarks"] = "Delivered 12 days late with 60-unit short shipment."

    # ANM-03: INV-2025-018 +24,000 overbill with valid GST arithmetic
    m3 = invoices_df["invoice_id"] == "INV-2025-018"
    orig_taxable = float(invoices_df.loc[m3, "taxable_amount"].iloc[0])
    gst_rate_18 = float(invoices_df.loc[m3, "gst_rate"].iloc[0])
    new_taxable = round(orig_taxable + 24000.0, 2)
    new_gst = round(new_taxable * gst_rate_18, 2)
    new_total = round(new_taxable + new_gst, 2)
    invoices_df.loc[m3, "taxable_amount"] = new_taxable
    invoices_df.loc[m3, "gst_amount"] = new_gst
    invoices_df.loc[m3, "total_amount"] = new_total
    invoices_df.loc[m3, "remarks"] = "Billed Rs 24,000 above PO-2025-018 taxable value (unreconciled surcharge)."
    # Keep PAY-018 aligned with invoice total or paid amount
    payments_df.loc[payments_df["invoice_id"] == "INV-2025-018", "payment_amount"] = new_total

    # ANM-04: Duplicate supplier invoice number SVE/25-26/0147 on INV-2025-019 and INV-2025-022
    invoices_df.loc[invoices_df["invoice_id"] == "INV-2025-019", "invoice_number"] = "SVE/25-26/0147"
    invoices_df.loc[invoices_df["invoice_id"] == "INV-2025-022", "invoice_number"] = "SVE/25-26/0147"
    invoices_df.loc[invoices_df["invoice_id"] == "INV-2025-022", "remarks"] = "Duplicate supplier invoice number SVE/25-26/0147 flagged."

    # ANM-12: Partial payment PAY-015 -> INV-2025-015 (198,000 vs 236,000)
    invoices_df.loc[invoices_df["invoice_id"] == "INV-2025-015", "payment_status"] = "Partial"
    invoices_df.loc[invoices_df["invoice_id"] == "INV-2025-015", "remarks"] = "Partial payment received; Rs 38,000 held pending reconciliation."
    m12 = payments_df["payment_id"] == "PAY-015"
    payments_df.loc[m12, "payment_amount"] = 198000.0
    payments_df.loc[m12, "status"] = "Completed"


# ---------------------------------------------------------------------------
# 2C. Quality Reports Generator (Part 3)
# ---------------------------------------------------------------------------
def gen_quality_reports(
    pos_df: pd.DataFrame,
    deliveries_df: pd.DataFrame,
) -> pd.DataFrame:
    """
    Generate 15 quality inspection reports (QR-001 to QR-015).

    Includes:
    - ANM-05: QR-010 (SUP005, PO-2025-018, PRD006) defect_rate=0.14, Fail
    - ANM-06: QR-012 (SUP001, PO-2025-021, PRD003) defect_rate=0.09, Conditional
      (preceded by QR-009 on SUP001 within two months)
    """
    po_map = pos_df.set_index("po_id").to_dict("index")
    del_map = deliveries_df.set_index("po_id")["actual_delivery_date"].to_dict()

    qr_specs = [
        ("QR-001", "PO-2025-001", 50, 1, "Pass", "R. Deshpande", "Minor surface scratch on 1 unit; accepted."),
        ("QR-002", "PO-2025-002", 80, 0, "Pass", "R. Deshpande", "Zero defects; full conformance to spec."),
        ("QR-003", "PO-2025-004", 100, 1, "Pass", "M. Nair", "Insulation thickness within tolerance."),
        ("QR-004", "PO-2025-005", 120, 1, "Pass", "M. Nair", "Dielectric test passed."),
        ("QR-005", "PO-2025-006", 500, 0, "Pass", "R. Deshpande", "Zero defects observed in sampling."),
        ("QR-006", "PO-2025-007", 60, 1, "Pass", "M. Nair", "Dimensions and conductivity verified."),
        ("QR-007", "PO-2025-011", 45, 0, "Pass", "R. Deshpande", "Calibration and RS485 output verified."),
        ("QR-008", "PO-2025-012", 35, 0, "Pass", "M. Nair", "Trip curve test passed on all units."),
        ("QR-009", "PO-2025-014", 80, 4, "Conditional", "R. Deshpande", "Armouring uniformity issue on 4m; sorted and accepted."),
        ("QR-010", "PO-2025-018", 50, 7, "Fail", "R. Deshpande", "ANM-05: 7/50 meters failed calibration & display check (14% defect rate)."),
        ("QR-011", "PO-2025-016", 100, 0, "Pass", "M. Nair", "Zero defects; switching endurance verified."),
        ("QR-012", "PO-2025-021", 100, 9, "Conditional", "R. Deshpande", "ANM-06: 9% defect rate on MCCB lot; second SUP001 quality issue in 2 months."),
        ("QR-013", "PO-2025-017", 150, 2, "Pass", "M. Nair", "Within 2% AQL threshold."),
        ("QR-014", "PO-2025-020", 30, 0, "Pass", "R. Deshpande", "Zero defects; 100% conformance."),
        ("QR-015", "PO-2025-022", 45, 1, "Pass", "M. Nair", "Minor burr on 1 bar; reworked at receiving."),
    ]

    rows = []
    for qr_id, po_id, batch_size, defects, result, inspector, remarks in qr_specs:
        po = po_map[po_id]
        insp_dt = min(DATASET_END, date.fromisoformat(del_map[po_id]) + timedelta(days=2))
        defect_rate = round(defects / batch_size, 4)
        rows.append(
            {
                "qr_id": qr_id,
                "supplier_id": po["supplier_id"],
                "po_id": po_id,
                "product_id": po["product_id"],
                "inspection_date": insp_dt.isoformat(),
                "batch_size": int(batch_size),
                "defects_found": int(defects),
                "defect_rate": defect_rate,
                "inspection_result": result,
                "inspector": inspector,
                "remarks": remarks,
            }
        )
    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# 2D. Unstructured Markdown Documents (12 files)
# ---------------------------------------------------------------------------
def write_markdown_docs(
    suppliers_df: pd.DataFrame,
    contracts_df: pd.DataFrame,
) -> list[Path]:
    """Generate exactly 12 Markdown documents (5 contracts, 3 emails, 2 quality, 2 management)."""
    sup_map = suppliers_df.set_index("supplier_id").to_dict("index")
    con_map = contracts_df.set_index("supplier_id").to_dict("index")
    written: list[Path] = []

    def save_md(folder_key: str, filename: str, content: str) -> None:
        path = DIRS[folder_key] / filename
        path.write_text(content.strip() + "\n", encoding="utf-8")
        written.append(path)

    contract_notes = {
        "SUP001": (
            "DOC-MD-001",
            "Standard pricing schedule applies for Switchgear & Busbars. "
            "Late delivery penalty of 1.0% of PO value per week applies after 14-day lead time.",
        ),
        "SUP002": (
            "DOC-MD-002",
            "Premium SLA agreement covering Automation & Relays with 10-day maximum lead time "
            "and mandatory ISO 9001:2015 compliance.",
        ),
        "SUP003": (
            "DOC-MD-003",
            "Fixed 12-month unit pricing schedule: PRD004 (XLPE Armoured Control Cable 12C x 2.5sqmm) "
            "is contracted at Rs 1,850.00 per Metre (ref PO-2025-004). Any revision requires a signed addendum.",
        ),
        "SUP004": (
            "DOC-MD-004",
            "Supplier Deccan Industrial Insulators LLP represents that its ISO 9001:2015 certification "
            "is valid and active for the full contract period (2025-04-01 to 2026-03-31). "
            "Immediate PO hold applies if compliance certificates lapse.",
        ),
        "SUP005": (
            "DOC-MD-005",
            "Unit-price increases exceeding 5% over baseline contract or standard rates are strictly "
            "prohibited without prior written buyer sign-off from Asteron Procurement.",
        ),
    }

    for sup_id in ["SUP001", "SUP002", "SUP003", "SUP004", "SUP005"]:
        sup = sup_map[sup_id]
        con = con_map[sup_id]
        doc_id, clause_note = contract_notes[sup_id]
        md = f"""# Supplier Master Agreement — {sup['supplier_name']}

- **Document ID:** {doc_id}
- **Document Type:** Supplier Contract
- **Date:** {con['contract_start_date']}
- **Contract ID:** {con['contract_id']}
- **Supplier ID:** {sup_id} ({sup['supplier_name']})
- **Buyer:** {COMPANY_NAME}, {COMPANY_LOCATION}

## Key Commercial Terms
- **Term:** {con['contract_start_date']} to {con['contract_end_date']}
- **Contract Value:** INR {con['contract_value']:,.2f}
- **Payment Terms:** {con['payment_terms']}
- **Delivery Terms:** {con['delivery_terms']}
- **Penalty & Governance Clause:** {con['penalty_clause']}
- **Auto-Renewal:** {con['auto_renewal']}

## Special Provisions
{clause_note}
"""
        save_md("contracts_md", f"contract_{sup_id}_{con['contract_id']}.md", md)

    save_md(
        "emails_md",
        "email_SUP003_pricing_revision.md",
        """# Email: Revised Pricing for PRD004 Control Cable

- **Document ID:** DOC-MD-006
- **Document Type:** Supplier Email
- **Date:** 2025-11-10
- **Supplier ID:** SUP003 (Sahyadri Power Components Pvt Ltd)
- **Contract ID:** CON-2025-003
- **From:** Anil Joshi <anil.joshi@sahyadripower.in>
- **To:** S. Gokhale <procurement@asteronindustrial.in>

Dear S. Gokhale,

Due to copper and XLPE compound input cost movements, please note our revised pricing for **PRD004 (XLPE Armoured Control Cable 12C x 2.5sqmm)** effective immediately is **INR 2,100.00 per Metre**, replacing the earlier INR 1,850.00 rate referenced in CON-2025-003 / PO-2025-004.

Regards,
Anil Joshi
Sahyadri Power Components Pvt Ltd
""",
    )

    save_md(
        "emails_md",
        "email_SUP001_performance_claim.md",
        """# Email: FY 2025-26 Delivery Performance Summary

- **Document ID:** DOC-MD-007
- **Document Type:** Supplier Email
- **Date:** 2026-02-15
- **Supplier ID:** SUP001 (Kaveri ElectroMech Works Pvt Ltd)
- **Contract ID:** CON-2025-001
- **From:** Rajesh Kulkarni <rajesh.kulkarni@kaverielectromech.in>
- **To:** S. Gokhale <procurement@asteronindustrial.in>

Dear Team Asteron,

As part of our annual review ahead of contract renewal, we confirm that Kaveri ElectroMech Works achieved a **95% on-time delivery (OTD)** rate across all Asteron Industrial Systems purchase orders during FY 2025-26, including PO-2025-001, PO-2025-003, PO-2025-008, PO-2025-014, and PO-2025-021.

Regards,
Rajesh Kulkarni
Kaveri ElectroMech Works Pvt Ltd
""",
    )

    save_md(
        "emails_md",
        "email_SUP005_invoice_clarification.md",
        """# Email: Clarification on PO-2025-018 and PO-2025-023 Billing

- **Document ID:** DOC-MD-008
- **Document Type:** Supplier Email
- **Date:** 2026-02-24
- **Supplier ID:** SUP005 (Suryodaya VoltEdge Enterprises)
- **Contract ID:** CON-2025-005
- **Related Records:** PO-2025-018, INV-2025-018, PO-2025-023
- **From:** Vikram Patwardhan <vikram.p@suryodayavoltedge.in>
- **To:** A. Kulkarni <procurement@asteronindustrial.in>

Dear A. Kulkarni,

Regarding **PO-2025-018 / INV-2025-018**, the additional INR 24,000.00 on the taxable value reflects an expedited calibration surcharge billed at dispatch. For **PO-2025-023 (PRD007)**, we billed INR 3,400.00/unit (versus standard INR 2,800.00/unit) prior to receiving formal written buyer sign-off.

Regards,
Vikram Patwardhan
Suryodaya VoltEdge Enterprises
""",
    )

    save_md(
        "quality_md",
        "quality_report_SUP005_QR-010.md",
        """# Receiving Quality Inspection Report — QR-010

- **Document ID:** DOC-MD-009
- **Document Type:** Quality Report
- **Date:** 2025-12-22
- **Supplier ID:** SUP005 (Suryodaya VoltEdge Enterprises)
- **PO ID:** PO-2025-018
- **Product ID:** PRD006 (Digital Multifunction Panel Meter)
- **Quality Record ID:** QR-010

## Inspection Findings
- **Batch Size Inspected:** 50 units
- **Defects Found:** 7 units
- **Defect Rate:** 14.0% (0.14)
- **Inspection Result:** **FAIL**
- **Inspector:** R. Deshpande

Seven meters failed RS485 Modbus communication and LCD segment diagnostics. Lot quarantined pending SCAR response from SUP005.
""",
    )

    save_md(
        "quality_md",
        "quality_report_SUP001_QR-012.md",
        """# Receiving Quality Inspection Report — QR-012

- **Document ID:** DOC-MD-010
- **Document Type:** Quality Report
- **Date:** 2026-02-04
- **Supplier ID:** SUP001 (Kaveri ElectroMech Works Pvt Ltd)
- **PO ID:** PO-2025-021
- **Product ID:** PRD003 (Molded Case Circuit Breaker 250A 3P)
- **Quality Record ID:** QR-012

## Inspection Findings
- **Sample Lot Size:** 100 units
- **Defects Found:** 9 units
- **Defect Rate:** 9.0% (0.09)
- **Inspection Result:** **CONDITIONAL**
- **Inspector:** R. Deshpande

This marks the **second quality issue from SUP001 within two months** (following QR-009 on PO-2025-014). Thermal trip calibration drifted outside IEC 60947-2 tolerance on 9% of sampled breakers.
""",
    )

    save_md(
        "management_md",
        "management_report_Q3_procurement_review.md",
        f"""# Q3 FY25-26 Procurement & Payables Exception Review

- **Document ID:** DOC-MD-011
- **Document Type:** Management Report
- **Date:** 2026-01-15
- **Scope:** {COMPANY_NAME} — Q3 Procurement Operations
- **Related Records:** INV-2025-015, PAY-015, INV-2025-018, INV-2025-019, INV-2025-022

## Key Exceptions Noted
1. **Partial Payment (SUP003):** Against `INV-2025-015` (INR 236,000.00), `PAY-015` settled INR 198,000.00, leaving INR 38,000.00 held pending reconciliation.
2. **Duplicate Invoice Number (SUP005):** Supplier invoice number `SVE/25-26/0147` appeared on `INV-2025-019` and was flagged again on `INV-2025-022`.
3. **Overbilling Alert (SUP005):** `INV-2025-018` exceeded `PO-2025-018` taxable value by INR 24,000.00.
""",
    )

    save_md(
        "management_md",
        "management_report_H1_supplier_reliability.md",
        f"""# H1 FY25-26 Supplier Delivery Reliability Briefing

- **Document ID:** DOC-MD-012
- **Document Type:** Management Report
- **Date:** 2025-10-08
- **Scope:** {COMPANY_NAME} — H1 (Apr–Sep 2025) Supplier OTD Review
- **Related Suppliers:** SUP001, SUP002, SUP003, SUP004, SUP005

## H1 Summary
- **SUP001 (Kaveri ElectroMech):** Experienced severe delays on `PO-2025-003` (`DEL-003`, 18 days late) and `PO-2025-008` (`DEL-008`, 12 days late, partial).
- **SUP002 (Vindhya Precision Controls):** Recorded **one late delivery** during H1 due to monsoon transit disruption in Navi Mumbai.
- **SUP003, SUP004, SUP005:** Maintained scheduled delivery windows during H1.
""",
    )

    return written


# ---------------------------------------------------------------------------
# 2E. Text-Based PDF Documents (5 files via ReportLab)
# ---------------------------------------------------------------------------
def _write_pdf(path: Path, title: str, lines: list[str]) -> None:
    """Render a clean, multi-line text-selectable PDF using ReportLab."""
    c = canvas.Canvas(str(path), pagesize=A4, pageCompression=0)
    width, height = A4
    y = height - 25 * mm

    c.setFont("Helvetica-Bold", 14)
    c.drawString(20 * mm, y, title)
    y -= 8 * mm

    c.setLineWidth(0.5)
    c.line(20 * mm, y, width - 20 * mm, y)
    y -= 8 * mm

    c.setFont("Helvetica", 10)
    for line in lines:
        if y < 20 * mm:
            c.showPage()
            c.setFont("Helvetica", 10)
            y = height - 20 * mm
        c.drawString(20 * mm, y, line)
        y -= 6 * mm

    c.save()


def write_pdf_docs() -> list[Path]:
    """Generate exactly 5 text-selectable PDF documents."""
    written: list[Path] = []

    pdf_specs = [
        (
            "cert_SUP004_iso9001.pdf",
            "ISO 9001:2015 QUALITY MANAGEMENT CERTIFICATE",
            [
                "Document ID: DOC-PDF-001",
                "Document Type: Supplier Compliance Certificate",
                "Issue Date: 2022-10-01",
                "Expiry Date: 2025-09-30",
                "Status: EXPIRED (as of 2025-10-01)",
                "Supplier ID: SUP004",
                "Supplier Name: Deccan Industrial Insulators LLP",
                "Address: Survey 45/2, Jeedimetla Industrial Estate, Hyderabad, Telangana",
                "Certificate No: ISO9001-IN-2022-88412",
                "Scope: Manufacture of Epoxy Cast Resin and Porcelain Insulators",
                "Note: Certificate expired on 2025-09-30 prior to PO-2025-013 (2025-10-10)",
                "and PO-2025-017 (2025-11-27). Renewal evidence not yet submitted.",
            ],
        ),
        (
            "cert_SUP002_iso9001.pdf",
            "ISO 9001:2015 QUALITY MANAGEMENT CERTIFICATE",
            [
                "Document ID: DOC-PDF-002",
                "Document Type: Supplier Compliance Certificate",
                "Issue Date: 2024-04-01",
                "Expiry Date: 2027-03-31",
                "Status: VALID / ACTIVE",
                "Supplier ID: SUP002",
                "Supplier Name: Vindhya Precision Controls Ltd",
                "Address: Unit 8, Mahape Electronic Zone, Navi Mumbai, Maharashtra",
                "Certificate No: ISO9001-IN-2024-31904",
                "Scope: Industrial Automation Relays, Panel Meters, and Control Gear",
                "Audit Status: Passed surveillance audit with zero major non-conformities.",
            ],
        ),
        (
            "scar_SUP005_corrective_action.pdf",
            "SUPPLIER CORRECTIVE ACTION REPORT (SCAR)",
            [
                "Document ID: DOC-PDF-003",
                "Document Type: Supplier Corrective Action Report",
                "Date: 2026-02-26",
                "Supplier ID: SUP005 (Suryodaya VoltEdge Enterprises)",
                "Contract ID: CON-2025-005",
                "Related Records: PO-2025-018, QR-010, INV-2025-018, INV-2025-019, INV-2025-022, PO-2025-023",
                "",
                "Non-Conformities Identified:",
                "1. Quality Failure (QR-010 / PO-2025-018): 14% defect rate (7/50 failed PRD006 units).",
                "2. Invoice Overbilling (INV-2025-018): Taxable amount billed Rs 24,000 above PO-2025-018.",
                "3. Duplicate Invoice Number: SVE/25-26/0147 reused on INV-2025-019 and INV-2025-022.",
                "4. Unapproved Price Increase (PO-2025-023): PRD007 billed at Rs 3,400 vs Rs 2,800 (+21.4%)",
                "   without written buyer sign-off required under CON-2025-005.",
                "Required Action: Credit note issuance, root-cause 8D report, and commercial freeze.",
            ],
        ),
        (
            "policy_asteron_procurement_2025.pdf",
            "ASTERON INDUSTRIAL SYSTEMS — PROCUREMENT GOVERNANCE POLICY",
            [
                "Document ID: DOC-PDF-004",
                "Document Type: Procurement Policy Document",
                "Effective Date: 2025-04-01",
                "Company Scope: Asteron Industrial Systems Pvt Ltd, Pune, Maharashtra, India",
                "",
                "1. Mandatory Compliance: All active suppliers must hold valid ISO 9001:2015",
                "   certification. Purchase orders must be placed on hold if a certificate expires.",
                "2. Three-Way Matching: Every invoice must match PO taxable value and GRN quantity.",
                "   Variances require written buyer and finance sign-off prior to settlement.",
                "3. Duplicate Invoice Control: ERP must block duplicate supplier invoice numbers.",
                "4. Price Escalation Control: Unit-price increases >5% require written buyer approval.",
                "5. Quality Threshold: Defect rates above 5% trigger Conditional hold or SCAR.",
            ],
        ),
        (
            "review_annual_supplier_performance_fy25_26.pdf",
            "ANNUAL SUPPLIER PERFORMANCE REVIEW — FY 2025-26",
            [
                "Document ID: DOC-PDF-005",
                "Document Type: Supplier Performance Review",
                "Date: 2026-03-25",
                "Company Scope: Asteron Industrial Systems Pvt Ltd (FY 2025-26)",
                "",
                "Supplier Scorecard Summary:",
                "- SUP001 (Kaveri ElectroMech): Low pricing, but 60% OTD (DEL-003 18d late,",
                "  DEL-008 12d late/partial) and repeat quality issues (QR-009, QR-012 9% defects).",
                "- SUP002 (Vindhya Precision): Highest unit pricing, 100% OTD (0 late deliveries),",
                "  0.0% defect rate across all inspected batches, valid ISO 9001 (DOC-PDF-002).",
                "- SUP003 (Sahyadri Power): Competitive pricing, 100% OTD, low defects; open items:",
                "  PRD004 email quote discrepancy (DOC-MD-006) and INV-2025-015 partial payment.",
                "- SUP004 (Deccan Insulators): Good delivery/quality, but ISO 9001 expired 2025-09-30.",
                "- SUP005 (Suryodaya VoltEdge): High risk — QR-010 (14% Fail), INV-2025-018 overbill,",
                "  duplicate invoice SVE/25-26/0147, and PO-2025-023 unapproved price jump.",
            ],
        ),
    ]

    for fname, title, lines in pdf_specs:
        out_path = DIRS["pdf"] / fname
        _write_pdf(out_path, title, lines)
        written.append(out_path)

    return written


# ---------------------------------------------------------------------------
# 3. Validation Assertions (Parts 1, 2, & 3)
# ---------------------------------------------------------------------------
def validate_part1(
    suppliers_df: pd.DataFrame,
    products_df: pd.DataFrame,
    contracts_df: pd.DataFrame,
    pos_df: pd.DataFrame,
) -> None:
    """Verify row counts, PK uniqueness, FK integrity, date bounds, and math."""
    assert len(suppliers_df) == 5, f"Expected 5 suppliers, got {len(suppliers_df)}"
    assert len(products_df) == 10, f"Expected 10 products, got {len(products_df)}"
    assert len(contracts_df) == 5, f"Expected 5 contracts, got {len(contracts_df)}"
    assert len(pos_df) == 25, f"Expected 25 POs, got {len(pos_df)}"

    assert suppliers_df["supplier_id"].is_unique
    assert products_df["product_id"].is_unique
    assert contracts_df["contract_id"].is_unique
    assert pos_df["po_id"].is_unique

    valid_sups = set(suppliers_df["supplier_id"])
    valid_prds = set(products_df["product_id"])
    assert set(contracts_df["supplier_id"]).issubset(valid_sups)
    assert set(pos_df["supplier_id"]).issubset(valid_sups)
    assert set(pos_df["product_id"]).issubset(valid_prds)

    for _, row in pos_df.iterrows():
        ord_dt = date.fromisoformat(row["order_date"])
        exp_dt = date.fromisoformat(row["expected_delivery_date"])
        assert DATASET_START <= ord_dt <= DATASET_END, f"PO {row['po_id']} out of range"
        assert exp_dt >= ord_dt, f"PO {row['po_id']} expected delivery before order"
        assert abs(row["total_amount"] - round(row["quantity"] * row["unit_price"], 2)) < 0.01
        assert abs(row["grand_total"] - round(row["total_amount"] + row["gst_amount"], 2)) < 0.01

    # Verify ANM-11 & ANM-12 setups
    po_23 = pos_df.loc[pos_df["po_id"] == "PO-2025-023"].iloc[0]
    std_p7 = products_df.loc[products_df["product_id"] == "PRD007", "standard_unit_price"].iloc[0]
    assert po_23["unit_price"] == 3400.0 and std_p7 == 2800.0

    po_15 = pos_df.loc[pos_df["po_id"] == "PO-2025-015"].iloc[0]
    assert po_15["grand_total"] == 236000.0


def validate_part2(
    pos_df: pd.DataFrame,
    deliveries_df: pd.DataFrame,
    invoices_df: pd.DataFrame,
    payments_df: pd.DataFrame,
) -> None:
    """Verify Part 2 counts, PK/FK integrity, GST arithmetic, and planted anomalies."""
    assert len(deliveries_df) == 25, f"Expected 25 deliveries, got {len(deliveries_df)}"
    assert len(invoices_df) == 25, f"Expected 25 invoices, got {len(invoices_df)}"
    assert len(payments_df) == 25, f"Expected 25 payments, got {len(payments_df)}"

    assert deliveries_df["delivery_id"].is_unique
    assert invoices_df["invoice_id"].is_unique
    assert payments_df["payment_id"].is_unique

    valid_pos = set(pos_df["po_id"])
    valid_sups = set(pos_df["supplier_id"])
    valid_invs = set(invoices_df["invoice_id"])

    assert set(deliveries_df["po_id"]).issubset(valid_pos)
    assert set(deliveries_df["supplier_id"]).issubset(valid_sups)
    assert set(invoices_df["po_id"]).issubset(valid_pos)
    assert set(invoices_df["supplier_id"]).issubset(valid_sups)
    assert set(payments_df["invoice_id"]).issubset(valid_invs)
    assert set(payments_df["supplier_id"]).issubset(valid_sups)

    # Check GST arithmetic on every invoice (including overbilled INV-2025-018)
    for _, inv in invoices_df.iterrows():
        exp_gst = round(inv["taxable_amount"] * inv["gst_rate"], 2)
        exp_tot = round(inv["taxable_amount"] + inv["gst_amount"], 2)
        assert abs(inv["gst_amount"] - exp_gst) < 0.01, f"GST mismatch on {inv['invoice_id']}"
        assert abs(inv["total_amount"] - exp_tot) < 0.01, f"Total mismatch on {inv['invoice_id']}"

    # Check payments never exceed invoice totals
    inv_totals = invoices_df.set_index("invoice_id")["total_amount"].to_dict()
    for _, pay in payments_df.iterrows():
        assert pay["payment_amount"] <= inv_totals[pay["invoice_id"]] + 0.01

    # Verify planted Part 2 anomalies
    d3 = deliveries_df.loc[deliveries_df["delivery_id"] == "DEL-003"].iloc[0]
    assert d3["delay_days"] == 18 and d3["delivery_status"] == "Late"
    d8 = deliveries_df.loc[deliveries_df["delivery_id"] == "DEL-008"].iloc[0]
    assert d8["delay_days"] == 12 and d8["delivery_status"] == "Partial"
    assert d8["quantity_delivered"] < d8["quantity_ordered"]

    # SUP001 OTD == 3/5 (60%) and SUP002 OTD == 5/5 (100%)
    sup1_dels = deliveries_df[deliveries_df["supplier_id"] == "SUP001"]
    assert (sup1_dels["delay_days"] == 0).sum() == 3
    sup2_dels = deliveries_df[deliveries_df["supplier_id"] == "SUP002"]
    assert (sup2_dels["delay_days"] == 0).sum() == 5

    # ANM-03: INV-2025-018 overbill by 24,000
    po18_amt = pos_df.loc[pos_df["po_id"] == "PO-2025-018", "total_amount"].iloc[0]
    inv18_amt = invoices_df.loc[invoices_df["invoice_id"] == "INV-2025-018", "taxable_amount"].iloc[0]
    assert abs((inv18_amt - po18_amt) - 24000.0) < 0.01

    # ANM-04: Duplicate invoice_number SVE/25-26/0147
    dup_cnt = (invoices_df["invoice_number"] == "SVE/25-26/0147").sum()
    assert dup_cnt == 2

    # ANM-12: Partial payment PAY-015
    p15 = payments_df.loc[payments_df["payment_id"] == "PAY-015"].iloc[0]
    assert p15["payment_amount"] == 198000.0 and inv_totals["INV-2025-015"] == 236000.0


def validate_part3(
    suppliers_df: pd.DataFrame,
    pos_df: pd.DataFrame,
    products_df: pd.DataFrame,
    quality_df: pd.DataFrame,
    md_files: list[Path],
    pdf_files: list[Path],
) -> None:
    """Verify quality_reports.csv, 12 Markdown docs, 5 text-based PDFs, and Part 3 anomalies."""
    assert len(quality_df) == 15, f"Expected 15 quality reports, got {len(quality_df)}"
    assert quality_df["qr_id"].is_unique

    assert set(quality_df["supplier_id"]).issubset(set(suppliers_df["supplier_id"]))
    assert set(quality_df["po_id"]).issubset(set(pos_df["po_id"]))
    assert set(quality_df["product_id"]).issubset(set(products_df["product_id"]))

    for _, row in quality_df.iterrows():
        exp_rate = round(row["defects_found"] / row["batch_size"], 4)
        assert abs(row["defect_rate"] - exp_rate) < 1e-4

    # ANM-05 & ANM-06 checks
    qr10 = quality_df.loc[quality_df["qr_id"] == "QR-010"].iloc[0]
    assert qr10["supplier_id"] == "SUP005" and qr10["defect_rate"] == 0.14 and qr10["inspection_result"] == "Fail"
    qr12 = quality_df.loc[quality_df["qr_id"] == "QR-012"].iloc[0]
    assert qr12["supplier_id"] == "SUP001" and qr12["defect_rate"] == 0.09 and qr12["inspection_result"] == "Conditional"

    # Document counts and PDF non-empty text stream check
    assert len(md_files) == 12, f"Expected 12 Markdown docs, got {len(md_files)}"
    assert len(pdf_files) == 5, f"Expected 5 PDF docs, got {len(pdf_files)}"
    for p in md_files:
        assert p.exists() and p.stat().st_size > 100
    for p in pdf_files:
        assert p.exists() and p.stat().st_size > 500
        raw_bytes = p.read_bytes()
        assert b"%PDF-" in raw_bytes and b"DOC-PDF-" in raw_bytes or b"BT" in raw_bytes


# ---------------------------------------------------------------------------
# 2F. Excel Workbook, Evaluation JSONs, Docs, Manifest & ZIP (Part 4)
# ---------------------------------------------------------------------------
def write_excel(datasets: dict[str, pd.DataFrame]) -> Path:
    """Write procurement_data.xlsx with 8 worksheets matching the 8 CSVs."""
    xlsx_path = DIRS["processed"] / "procurement_data.xlsx"
    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        for sheet_name, df in datasets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            ws = writer.sheets[sheet_name]
            for col_idx, col_name in enumerate(df.columns, start=1):
                ws.cell(row=1, column=col_idx).font = Font(bold=True)
                max_len = max(
                    [len(str(col_name))] + [len(str(v)) for v in df[col_name].tolist()]
                )
                ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 4, 48)
    return xlsx_path


def write_evaluation() -> tuple[Path, Path]:
    """Write evaluation/decision_scenarios.json (10 scenarios) and ground_truth_anomalies.json."""
    scenarios = [
        {
            "scenario_id": "SCN-01",
            "category": "Overall supplier selection",
            "question": "Which supplier should Asteron Industrial Systems select for our next procurement contract, and why?",
            "expected_recommendation": "SUP003 (Sahyadri Power Components Pvt Ltd) for balanced cost, 100% OTD, and low defects, or SUP002 (Vindhya Precision Controls Ltd) if zero-defect reliability outweighs higher pricing; exclude SUP001, SUP004 (until ISO 9001 renewed), and SUP005.",
            "key_evidence_ids": ["SUP002", "SUP003", "SUP001", "SUP004", "SUP005", "DOC-PDF-005"],
        },
        {
            "scenario_id": "SCN-02",
            "category": "Delivery performance",
            "question": "Which suppliers experienced delivery delays or short shipments during FY 2025-26?",
            "expected_recommendation": "Flag SUP001 for DEL-003 (18 days late on PO-2025-003) and DEL-008 (12 days late and 60-unit short shipment on PO-2025-008).",
            "key_evidence_ids": ["DEL-003", "PO-2025-003", "DEL-008", "PO-2025-008"],
        },
        {
            "scenario_id": "SCN-03",
            "category": "Invoice reconciliation",
            "question": "Are there any invoices whose taxable amount does not match the corresponding purchase order?",
            "expected_recommendation": "INV-2025-018 (SUP005) billed INR 181,500.00 taxable vs. PO-2025-018 total_amount INR 157,500.00 (+INR 24,000.00 overbilling), with GST recalculated accurately on the inflated base.",
            "key_evidence_ids": ["INV-2025-018", "PO-2025-018", "DOC-MD-008", "DOC-PDF-003"],
        },
        {
            "scenario_id": "SCN-04",
            "category": "Duplicate invoices",
            "question": "Has any supplier submitted a duplicate supplier invoice number across different records?",
            "expected_recommendation": "SUP005 reused supplier invoice number SVE/25-26/0147 on both INV-2025-019 (PO-2025-019) and INV-2025-022 (PO-2025-022).",
            "key_evidence_ids": ["INV-2025-019", "INV-2025-022", "DOC-MD-011", "DOC-PDF-003"],
        },
        {
            "scenario_id": "SCN-05",
            "category": "Quality risks",
            "question": "Which suppliers pose quality risks based on receiving inspection defect rates?",
            "expected_recommendation": "SUP005 failed QR-010 (14% defect rate on PO-2025-018 / PRD006) and SUP001 had repeat quality issues on QR-009 and QR-012 (9% defect rate, Conditional on PO-2025-021 / PRD003).",
            "key_evidence_ids": ["QR-010", "QR-012", "QR-009", "DOC-MD-009", "DOC-MD-010"],
        },
        {
            "scenario_id": "SCN-06",
            "category": "Compliance gaps",
            "question": "Do all active suppliers hold valid mandatory ISO 9001 certification for the POs placed in Q3-Q4?",
            "expected_recommendation": "No. SUP004's ISO 9001 certificate (DOC-PDF-001) expired on 2025-09-30, yet PO-2025-013 (2025-10-10) and PO-2025-017 (2025-11-27) were placed despite DOC-MD-004 requiring valid ISO 9001.",
            "key_evidence_ids": ["DOC-PDF-001", "DOC-MD-004", "CON-2025-004", "PO-2025-013", "PO-2025-017"],
        },
        {
            "scenario_id": "SCN-07",
            "category": "Contract/email contradictions",
            "question": "Are there contradictions between signed supplier contracts, emails, or internal management reports?",
            "expected_recommendation": "Yes: (1) DOC-MD-003 contracts PRD004 at INR 1,850 while email DOC-MD-006 quotes INR 2,100 without an addendum; (2) DOC-MD-012 claims SUP002 had 1 late H1 delivery while supplier_deliveries.csv shows 0.",
            "key_evidence_ids": ["DOC-MD-003", "DOC-MD-006", "DOC-MD-012", "SUP002", "SUP003"],
        },
        {
            "scenario_id": "SCN-08",
            "category": "Supplier performance claim verification",
            "question": "Verify SUP001's email claim of 95% on-time delivery against actual delivery records.",
            "expected_recommendation": "SUP001's claim in DOC-MD-007 (95% OTD) is contradicted by supplier_deliveries.csv, which shows 3 of 5 deliveries on time (60% OTD) due to DEL-003 and DEL-008.",
            "key_evidence_ids": ["DOC-MD-007", "DEL-001", "DEL-003", "DEL-008", "DEL-014", "DEL-021"],
        },
        {
            "scenario_id": "SCN-09",
            "category": "Unapproved price increases",
            "question": "Did any supplier charge a unit price significantly above the standard rate without documented buyer approval?",
            "expected_recommendation": "On PO-2025-023 (SUP005, PRD007), unit_price jumped to INR 3,400.00 vs. standard INR 2,800.00 (+21.4%) with approved_by='Pending Buyer Sign-Off', violating CON-2025-005's 5% threshold.",
            "key_evidence_ids": ["PO-2025-023", "PRD007", "CON-2025-005", "DOC-MD-005", "DOC-MD-008"],
        },
        {
            "scenario_id": "SCN-10",
            "category": "What-if weight changes",
            "question": "How does the top supplier recommendation change if decision weights shift from Cost-focused (50% price) to Risk/Quality-focused (60% quality + compliance + OTD)?",
            "expected_recommendation": "Under Cost-focused weights, SUP003 ranks #1 (competitive pricing, 100% OTD). Under Risk/Quality-focused weights, SUP002 ranks #1 (0% defect rate, 100% OTD, valid ISO 9001 in DOC-PDF-002) despite higher pricing.",
            "key_evidence_ids": ["SUP002", "SUP003", "DOC-PDF-002", "DOC-PDF-005"],
        },
    ]

    anomalies = [
        {
            "anomaly_id": "ANM-01",
            "anomaly_type": "delivery_delay",
            "supplier_id": "SUP001",
            "source_file": "data/raw/supplier_deliveries.csv",
            "source_record_id": "DEL-003",
            "related_record_ids": ["PO-2025-003", "CON-2025-001"],
            "explanation": "DEL-003 arrived 18 days after expected_delivery_date.",
            "expected_detection_method": "Compare actual_delivery_date vs expected_delivery_date / delay_days > 0.",
            "severity": "High",
        },
        {
            "anomaly_id": "ANM-02",
            "anomaly_type": "delivery_delay_and_shortage",
            "supplier_id": "SUP001",
            "source_file": "data/raw/supplier_deliveries.csv",
            "source_record_id": "DEL-008",
            "related_record_ids": ["PO-2025-008"],
            "explanation": "DEL-008 arrived 12 days late and delivered 340 of 400 ordered units (Partial).",
            "expected_detection_method": "Check delay_days > 0 and quantity_delivered < quantity_ordered.",
            "severity": "High",
        },
        {
            "anomaly_id": "ANM-03",
            "anomaly_type": "invoice_po_mismatch",
            "supplier_id": "SUP005",
            "source_file": "data/raw/invoices.csv",
            "source_record_id": "INV-2025-018",
            "related_record_ids": ["PO-2025-018", "DOC-MD-008", "DOC-PDF-003"],
            "explanation": "INV-2025-018 taxable_amount (181,500.00) is INR 24,000.00 higher than PO-2025-018 total_amount (157,500.00), with GST recalculated at 18%.",
            "expected_detection_method": "Three-way match comparing invoices.taxable_amount against purchase_orders.total_amount.",
            "severity": "Critical",
        },
        {
            "anomaly_id": "ANM-04",
            "anomaly_type": "duplicate_invoice_number",
            "supplier_id": "SUP005",
            "source_file": "data/raw/invoices.csv",
            "source_record_id": "INV-2025-022",
            "related_record_ids": ["INV-2025-019", "PO-2025-019", "PO-2025-022", "DOC-MD-011"],
            "explanation": "Supplier invoice number SVE/25-26/0147 is duplicated across INV-2025-019 and INV-2025-022.",
            "expected_detection_method": "Group invoices by (supplier_id, invoice_number) and flag count > 1.",
            "severity": "Critical",
        },
        {
            "anomaly_id": "ANM-05",
            "anomaly_type": "quality_failure",
            "supplier_id": "SUP005",
            "source_file": "data/raw/quality_reports.csv",
            "source_record_id": "QR-010",
            "related_record_ids": ["PO-2025-018", "PRD006", "DOC-MD-009", "DOC-PDF-003"],
            "explanation": "QR-010 failed inspection with 7/50 defective units (defect_rate = 0.14).",
            "expected_detection_method": "Filter quality_reports for inspection_result == 'Fail' or defect_rate > 0.05.",
            "severity": "Critical",
        },
        {
            "anomaly_id": "ANM-06",
            "anomaly_type": "repeat_quality_failure",
            "supplier_id": "SUP001",
            "source_file": "data/raw/quality_reports.csv",
            "source_record_id": "QR-012",
            "related_record_ids": ["QR-009", "PO-2025-021", "PRD003", "DOC-MD-010"],
            "explanation": "QR-012 recorded 9% defect rate (Conditional), marking SUP001's second quality issue within two months.",
            "expected_detection_method": "Track rolling quality_reports defect_rate > 0.05 per supplier.",
            "severity": "High",
        },
        {
            "anomaly_id": "ANM-07",
            "anomaly_type": "expired_compliance_certificate",
            "supplier_id": "SUP004",
            "source_file": "documents/pdf/cert_SUP004_iso9001.pdf",
            "source_record_id": "DOC-PDF-001",
            "related_record_ids": ["DOC-MD-004", "CON-2025-004", "PO-2025-013", "PO-2025-017"],
            "explanation": "SUP004 ISO 9001 certificate expired on 2025-09-30 while DOC-MD-004 claims active coverage and POs were placed in Oct-Nov 2025.",
            "expected_detection_method": "Extract certificate expiry_date from PDF and compare against PO order_dates and contract claims.",
            "severity": "High",
        },
        {
            "anomaly_id": "ANM-08",
            "anomaly_type": "contract_email_contradiction",
            "supplier_id": "SUP003",
            "source_file": "documents/emails/email_SUP003_pricing_revision.md",
            "source_record_id": "DOC-MD-006",
            "related_record_ids": ["DOC-MD-003", "CON-2025-003", "PO-2025-004", "PRD004"],
            "explanation": "Email DOC-MD-006 quotes PRD004 at INR 2,100/m, contradicting contracted INR 1,850/m in DOC-MD-003 without a signed addendum.",
            "expected_detection_method": "Cross-check quoted unit prices in emails against contract markdown and PO prices.",
            "severity": "Medium",
        },
        {
            "anomaly_id": "ANM-09",
            "anomaly_type": "report_data_contradiction",
            "supplier_id": "SUP002",
            "source_file": "documents/management/management_report_H1_supplier_reliability.md",
            "source_record_id": "DOC-MD-012",
            "related_record_ids": ["DEL-002", "DEL-006", "DEL-011"],
            "explanation": "Management report DOC-MD-012 claims SUP002 had 1 late delivery in H1, whereas supplier_deliveries.csv shows 0 late deliveries for SUP002.",
            "expected_detection_method": "Verify narrative OTD claims in management reports against supplier_deliveries.csv.",
            "severity": "Low",
        },
        {
            "anomaly_id": "ANM-10",
            "anomaly_type": "performance_claim_contradiction",
            "supplier_id": "SUP001",
            "source_file": "documents/emails/email_SUP001_performance_claim.md",
            "source_record_id": "DOC-MD-007",
            "related_record_ids": ["DEL-001", "DEL-003", "DEL-008", "DEL-014", "DEL-021"],
            "explanation": "SUP001 email DOC-MD-007 claims 95% OTD, but actual delivery records show 3/5 on-time (60% OTD).",
            "expected_detection_method": "Compute actual OTD% from supplier_deliveries.csv and compare to supplier email claims.",
            "severity": "High",
        },
        {
            "anomaly_id": "ANM-11",
            "anomaly_type": "unapproved_price_increase",
            "supplier_id": "SUP005",
            "source_file": "data/raw/purchase_orders.csv",
            "source_record_id": "PO-2025-023",
            "related_record_ids": ["PRD007", "CON-2025-005", "DOC-MD-005", "DOC-MD-008"],
            "explanation": "PO-2025-023 bills PRD007 at INR 3,400.00 vs. standard INR 2,800.00 (+21.4%) with approved_by='Pending Buyer Sign-Off'.",
            "expected_detection_method": "Compare PO unit_price against products.standard_unit_price (>5% increase) and check approved_by.",
            "severity": "High",
        },
        {
            "anomaly_id": "ANM-12",
            "anomaly_type": "partial_payment",
            "supplier_id": "SUP003",
            "source_file": "data/raw/payments.csv",
            "source_record_id": "PAY-015",
            "related_record_ids": ["INV-2025-015", "PO-2025-015", "DOC-MD-011"],
            "explanation": "PAY-015 paid INR 198,000.00 against INV-2025-015 total INR 236,000.00 (INR 38,000.00 shortfall).",
            "expected_detection_method": "Reconcile payments.payment_amount against invoices.total_amount and payment_status.",
            "severity": "Medium",
        },
    ]

    scn_path = DIRS["evaluation"] / "decision_scenarios.json"
    anm_path = DIRS["evaluation"] / "ground_truth_anomalies.json"
    scn_path.write_text(json.dumps(scenarios, indent=2), encoding="utf-8")
    anm_path.write_text(json.dumps(anomalies, indent=2), encoding="utf-8")
    return scn_path, anm_path


def write_manifest(
    datasets: dict[str, pd.DataFrame],
    md_files: list[Path],
    pdf_files: list[Path],
    xlsx_path: Path,
    scn_path: Path,
    anm_path: Path,
) -> Path:
    """Write SOURCE_MANIFEST.json indexing all generated artifacts and IDs."""
    manifest = {
        "project": "DecisionGraph",
        "company": COMPANY_NAME,
        "location": COMPANY_LOCATION,
        "period": {"start": DATASET_START.isoformat(), "end": DATASET_END.isoformat()},
        "currency": "INR",
        "seed": SEED,
        "csv_files": {
            name: {
                "path": f"data/raw/{name}.csv",
                "rows": len(df),
                "columns": list(df.columns),
            }
            for name, df in datasets.items()
        },
        "excel_workbook": {
            "path": str(xlsx_path.relative_to(BASE_DIR)).replace("\\", "/"),
            "sheets": list(datasets.keys()),
        },
        "markdown_documents": [
            str(p.relative_to(BASE_DIR)).replace("\\", "/") for p in md_files
        ],
        "pdf_documents": [
            str(p.relative_to(BASE_DIR)).replace("\\", "/") for p in pdf_files
        ],
        "evaluation_files": [
            str(scn_path.relative_to(BASE_DIR)).replace("\\", "/"),
            str(anm_path.relative_to(BASE_DIR)).replace("\\", "/"),
        ],
        "document_ids": [f"DOC-MD-{i:03d}" for i in range(1, 13)]
        + [f"DOC-PDF-{i:03d}" for i in range(1, 6)],
        "anomaly_ids": [f"ANM-{i:02d}" for i in range(1, 13)],
    }
    manifest_path = BASE_DIR / "SOURCE_MANIFEST.json"
    manifest_path.write_text(json.dumps(manifest, indent=2), encoding="utf-8")
    return manifest_path


def write_readme_and_dictionary() -> tuple[Path, Path]:
    """Write README.md and DATA_DICTIONARY.md at dataset root."""
    readme_text = f"""# DecisionGraph Synthetic Procurement Dataset

Synthetic, deterministic procurement dataset for **{COMPANY_NAME}** ({COMPANY_LOCATION}), covering **{DATASET_START.isoformat()} to {DATASET_END.isoformat()}** (INR, Indian GST).

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
"""
    dict_text = """# Data Dictionary — DecisionGraph Dataset

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
"""
    readme_path = BASE_DIR / "README.md"
    dict_path = BASE_DIR / "DATA_DICTIONARY.md"
    readme_path.write_text(readme_text, encoding="utf-8")
    dict_path.write_text(dict_text, encoding="utf-8")
    return readme_path, dict_path


def create_zip() -> Path:
    """Package the complete decisiongraph_dataset directory into decisiongraph_dataset.zip."""
    zip_path = BASE_DIR.parent / "decisiongraph_dataset.zip"
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for file_path in sorted(BASE_DIR.rglob("*")):
            if file_path.is_file() and "__pycache__" not in file_path.parts:
                arcname = Path(BASE_DIR.name) / file_path.relative_to(BASE_DIR)
                zf.write(file_path, arcname=str(arcname))
    return zip_path


def main() -> None:
    ensure_directories()

    suppliers_df = gen_suppliers()
    products_df = gen_products()
    contracts_df = gen_contracts()
    pos_df = gen_purchase_orders(products_df)

    deliveries_df = gen_deliveries(pos_df)
    invoices_df = gen_invoices(pos_df, products_df, contracts_df, deliveries_df)
    payments_df = gen_payments(invoices_df)

    apply_part2_anomalies(deliveries_df, invoices_df, payments_df)

    quality_df = gen_quality_reports(pos_df, deliveries_df)
    md_files = write_markdown_docs(suppliers_df, contracts_df)
    pdf_files = write_pdf_docs()

    validate_part1(suppliers_df, products_df, contracts_df, pos_df)
    validate_part2(pos_df, deliveries_df, invoices_df, payments_df)
    validate_part3(suppliers_df, pos_df, products_df, quality_df, md_files, pdf_files)

    datasets = {
        "suppliers": suppliers_df,
        "products": products_df,
        "purchase_orders": pos_df,
        "supplier_deliveries": deliveries_df,
        "invoices": invoices_df,
        "payments": payments_df,
        "contracts": contracts_df,
        "quality_reports": quality_df,
    }

    for name, df in datasets.items():
        df.to_csv(DIRS["raw"] / f"{name}.csv", index=False)

    xlsx_path = write_excel(datasets)
    scn_path, anm_path = write_evaluation()
    write_manifest(datasets, md_files, pdf_files, xlsx_path, scn_path, anm_path)
    write_readme_and_dictionary()

    # Run standalone validator if present so validation_report.json is fresh before zipping
    validator_script = DIRS["scripts"] / "validate_dataset.py"
    if validator_script.exists():
        subprocess.run([sys.executable, str(validator_script)], check=True)

    zip_path = create_zip()

    print("All parts generated successfully:")
    print(f"  - 8 CSV files written to {DIRS['raw']}")
    print(f"  - Excel workbook: {xlsx_path}")
    print(f"  - {len(md_files)} Markdown + {len(pdf_files)} PDF documents")
    print(f"  - Evaluation JSONs: {scn_path.name}, {anm_path.name}")
    print(f"  - ZIP package: {zip_path}")


if __name__ == "__main__":
    main()
