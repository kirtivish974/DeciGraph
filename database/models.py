from datetime import date
from decimal import Decimal

from sqlalchemy import (
    String,
    Integer,
    Boolean,
    Date,
    Numeric,
    ForeignKey,
    Text,
)
from sqlalchemy.orm import Mapped, mapped_column, relationship

from database.db import Base


class Supplier(Base):
    __tablename__ = "suppliers"

    supplier_id: Mapped[str] = mapped_column(String(20), primary_key=True)
    supplier_name: Mapped[str] = mapped_column(String(150), nullable=False)
    contact_person: Mapped[str | None] = mapped_column(String(100))
    email: Mapped[str | None] = mapped_column(String(150))
    phone: Mapped[str | None] = mapped_column(String(30))
    address: Mapped[str | None] = mapped_column(Text)
    city: Mapped[str | None] = mapped_column(String(100))
    state: Mapped[str | None] = mapped_column(String(100))
    gstin: Mapped[str | None] = mapped_column(String(20))
    registration_date: Mapped[date | None] = mapped_column(Date)
    supplier_category: Mapped[str | None] = mapped_column(String(100))
    risk_rating: Mapped[str | None] = mapped_column(String(30))
    status: Mapped[str | None] = mapped_column(String(30))

    contracts = relationship("Contract", back_populates="supplier")
    purchase_orders = relationship("PurchaseOrder", back_populates="supplier")
    invoices = relationship("Invoice", back_populates="supplier")
    payments = relationship("Payment", back_populates="supplier")
    deliveries = relationship("SupplierDelivery", back_populates="supplier")
    quality_reports = relationship("QualityReport", back_populates="supplier")


class Product(Base):
    __tablename__ = "products"

    product_id: Mapped[str] = mapped_column(String(20), primary_key=True)
    product_name: Mapped[str] = mapped_column(String(150), nullable=False)
    category: Mapped[str | None] = mapped_column(String(100))
    unit_of_measure: Mapped[str | None] = mapped_column(String(30))
    hsn_code: Mapped[str | None] = mapped_column(String(20))
    standard_unit_price: Mapped[Decimal | None] = mapped_column(Numeric(14, 2))
    gst_rate: Mapped[Decimal | None] = mapped_column(Numeric(6, 2))
    description: Mapped[str | None] = mapped_column(Text)

    purchase_orders = relationship("PurchaseOrder", back_populates="product")
    quality_reports = relationship("QualityReport", back_populates="product")


class Contract(Base):
    __tablename__ = "contracts"

    contract_id: Mapped[str] = mapped_column(String(20), primary_key=True)
    supplier_id: Mapped[str] = mapped_column(ForeignKey("suppliers.supplier_id"))
    contract_start_date: Mapped[date | None] = mapped_column(Date)
    contract_end_date: Mapped[date | None] = mapped_column(Date)
    contract_value: Mapped[Decimal | None] = mapped_column(Numeric(16, 2))
    payment_terms: Mapped[str | None] = mapped_column(Text)
    delivery_terms: Mapped[str | None] = mapped_column(Text)
    penalty_clause: Mapped[str | None] = mapped_column(Text)
    auto_renewal: Mapped[bool | None] = mapped_column(Boolean)
    status: Mapped[str | None] = mapped_column(String(30))

    supplier = relationship("Supplier", back_populates="contracts")


class PurchaseOrder(Base):
    __tablename__ = "purchase_orders"

    po_id: Mapped[str] = mapped_column(String(20), primary_key=True)
    supplier_id: Mapped[str] = mapped_column(ForeignKey("suppliers.supplier_id"))
    product_id: Mapped[str] = mapped_column(ForeignKey("products.product_id"))
    order_date: Mapped[date | None] = mapped_column(Date)
    expected_delivery_date: Mapped[date | None] = mapped_column(Date)
    quantity: Mapped[int | None] = mapped_column(Integer)
    unit_price: Mapped[Decimal | None] = mapped_column(Numeric(14, 2))
    total_amount: Mapped[Decimal | None] = mapped_column(Numeric(16, 2))
    gst_amount: Mapped[Decimal | None] = mapped_column(Numeric(16, 2))
    grand_total: Mapped[Decimal | None] = mapped_column(Numeric(16, 2))
    status: Mapped[str | None] = mapped_column(String(30))
    approved_by: Mapped[str | None] = mapped_column(String(100))

    supplier = relationship("Supplier", back_populates="purchase_orders")
    product = relationship("Product", back_populates="purchase_orders")
    invoices = relationship("Invoice", back_populates="purchase_order")
    deliveries = relationship("SupplierDelivery", back_populates="purchase_order")
    quality_reports = relationship("QualityReport", back_populates="purchase_order")


class Invoice(Base):
    __tablename__ = "invoices"

    invoice_id: Mapped[str] = mapped_column(String(20), primary_key=True)
    invoice_number: Mapped[str] = mapped_column(String(50), nullable=False)
    po_id: Mapped[str] = mapped_column(ForeignKey("purchase_orders.po_id"))
    supplier_id: Mapped[str] = mapped_column(ForeignKey("suppliers.supplier_id"))
    invoice_date: Mapped[date | None] = mapped_column(Date)
    taxable_amount: Mapped[Decimal | None] = mapped_column(Numeric(16, 2))
    gst_rate: Mapped[Decimal | None] = mapped_column(Numeric(6, 2))
    gst_amount: Mapped[Decimal | None] = mapped_column(Numeric(16, 2))
    total_amount: Mapped[Decimal | None] = mapped_column(Numeric(16, 2))
    due_date: Mapped[date | None] = mapped_column(Date)
    payment_status: Mapped[str | None] = mapped_column(String(30))
    remarks: Mapped[str | None] = mapped_column(Text)

    supplier = relationship("Supplier", back_populates="invoices")
    purchase_order = relationship("PurchaseOrder", back_populates="invoices")
    payments = relationship("Payment", back_populates="invoice")


class Payment(Base):
    __tablename__ = "payments"

    payment_id: Mapped[str] = mapped_column(String(20), primary_key=True)
    invoice_id: Mapped[str] = mapped_column(ForeignKey("invoices.invoice_id"))
    supplier_id: Mapped[str] = mapped_column(ForeignKey("suppliers.supplier_id"))
    payment_date: Mapped[date | None] = mapped_column(Date)
    payment_amount: Mapped[Decimal | None] = mapped_column(Numeric(16, 2))
    payment_mode: Mapped[str | None] = mapped_column(String(50))
    reference_number: Mapped[str | None] = mapped_column(String(100))
    status: Mapped[str | None] = mapped_column(String(30))

    supplier = relationship("Supplier", back_populates="payments")
    invoice = relationship("Invoice", back_populates="payments")


class SupplierDelivery(Base):
    __tablename__ = "supplier_deliveries"

    delivery_id: Mapped[str] = mapped_column(String(20), primary_key=True)
    po_id: Mapped[str] = mapped_column(ForeignKey("purchase_orders.po_id"))
    supplier_id: Mapped[str] = mapped_column(ForeignKey("suppliers.supplier_id"))
    expected_delivery_date: Mapped[date | None] = mapped_column(Date)
    actual_delivery_date: Mapped[date | None] = mapped_column(Date)
    quantity_ordered: Mapped[int | None] = mapped_column(Integer)
    quantity_delivered: Mapped[int | None] = mapped_column(Integer)
    delivery_status: Mapped[str | None] = mapped_column(String(30))
    delay_days: Mapped[int | None] = mapped_column(Integer)
    remarks: Mapped[str | None] = mapped_column(Text)

    supplier = relationship("Supplier", back_populates="deliveries")
    purchase_order = relationship("PurchaseOrder", back_populates="deliveries")


class QualityReport(Base):
    __tablename__ = "quality_reports"

    qr_id: Mapped[str] = mapped_column(String(20), primary_key=True)
    supplier_id: Mapped[str] = mapped_column(ForeignKey("suppliers.supplier_id"))
    po_id: Mapped[str] = mapped_column(ForeignKey("purchase_orders.po_id"))
    product_id: Mapped[str] = mapped_column(ForeignKey("products.product_id"))
    inspection_date: Mapped[date | None] = mapped_column(Date)
    batch_size: Mapped[int | None] = mapped_column(Integer)
    defects_found: Mapped[int | None] = mapped_column(Integer)
    defect_rate: Mapped[Decimal | None] = mapped_column(Numeric(6, 2))
    inspection_result: Mapped[str | None] = mapped_column(String(50))
    inspector: Mapped[str | None] = mapped_column(String(100))
    remarks: Mapped[str | None] = mapped_column(Text)

    supplier = relationship("Supplier", back_populates="quality_reports")
    purchase_order = relationship("PurchaseOrder", back_populates="quality_reports")
    product = relationship("Product", back_populates="quality_reports")