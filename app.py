# ============================================================
#  KANBAN PRODUCTION CALCULATOR (FINAL VERSION)
# ============================================================
import io
import os
import csv
import smtplib
from math import ceil
from datetime import datetime
from email.message import EmailMessage
from typing import Optional, List

from fastapi import (FastAPI, Depends, HTTPException, Query,
                     UploadFile, File, Response)
from fastapi.middleware.cors import CORSMiddleware
from sqlalchemy import (create_engine, Column, Integer, Float, String,
                        ForeignKey, DateTime, Boolean, Text,
                        UniqueConstraint, or_)
from sqlalchemy.orm import declarative_base, sessionmaker, Session
from pydantic import BaseModel, Field
from openpyxl import load_workbook, Workbook

from reportlab.pdfgen import canvas as rl_canvas
from reportlab.lib.units import mm
from reportlab.lib.colors import HexColor
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.graphics.barcode import code128

# ============================================================
#  НАСТРОЙКА БАЗЫ ДАННЫХ
# ============================================================
SQLALCHEMY_URL = "sqlite:///./kanban.db"
engine = create_engine(SQLALCHEMY_URL, connect_args={"check_same_thread": False})
SessionLocal = sessionmaker(bind=engine, autoflush=False, autocommit=False)
Base = declarative_base()

def _now() -> datetime:
    return datetime.utcnow()

# ============================================================
#  МОДЕЛИ (ТАБЛИЦЫ)
# ============================================================
class Site(Base):
    __tablename__ = "sites"
    id = Column(Integer, primary_key=True)
    name = Column(String, unique=True, nullable=False)
    code = Column(String, unique=True, nullable=False)
    site_type = Column(String, default="production") # production | warehouse

class Material(Base):
    __tablename__ = "materials"
    id = Column(Integer, primary_key=True)
    code = Column(String, unique=True, nullable=False)
    name = Column(String, nullable=False)
    unit = Column(String, nullable=False, default="кг")
    kind = Column(String, default="main")
    density_kg_m3 = Column(Float, nullable=True)

class Product(Base):
    __tablename__ = "products"
    id = Column(Integer, primary_key=True)
    code = Column(String, unique=True, nullable=False)
    name = Column(String, nullable=False)

class Packaging(Base):
    __tablename__ = "packagings"
    id = Column(Integer, primary_key=True)
    code = Column(String, unique=True, nullable=False)
    name = Column(String, nullable=False)
    packaging_type = Column(String, nullable=False) # box|pallet|barrel|ibc|tank|bag|roll|canister
    capacity = Column(Float, nullable=False)
    capacity_unit = Column(String, nullable=False)
    tare_weight_kg = Column(Float, nullable=True)
    net_weight_kg = Column(Float, nullable=True)
    is_returnable = Column(Boolean, default=False)
    notes = Column(Text, nullable=True)
    parent_id = Column(Integer, ForeignKey("packagings.id"), nullable=True)
    children_per_unit = Column(Float, nullable=True)

class MaterialPackaging(Base):
    __tablename__ = "material_packagings"
    __table_args__ = (UniqueConstraint("material_id", "packaging_id"),)
    id = Column(Integer, primary_key=True)
    material_id = Column(Integer, ForeignKey("materials.id"), nullable=False)
    packaging_id = Column(Integer, ForeignKey("packagings.id"), nullable=False)
    is_primary = Column(Boolean, default=False)
    supplier = Column(String, nullable=True)
    part_number = Column(String, nullable=True)
    conversion_notes = Column(Text, nullable=True)

class BOMLine(Base):
    __tablename__ = "bom_lines"
    __table_args__ = (UniqueConstraint("product_id", "material_id", "site_id", "version"),)
    id = Column(Integer, primary_key=True)
    product_id = Column(Integer, ForeignKey("products.id"), nullable=False)
    material_id = Column(Integer, ForeignKey("materials.id"), nullable=False)
    site_id = Column(Integer, ForeignKey("sites.id"), nullable=False)
    qty_per_unit = Column(Float, nullable=False)
    version = Column(Integer, nullable=False, default=1)
    valid_from = Column(DateTime, nullable=False, default=_now)
    valid_to = Column(DateTime, nullable=True)
    updated_at = Column(DateTime, default=_now, onupdate=_now)

class BOMHistory(Base):
    __tablename__ = "bom_history"
    id = Column(Integer, primary_key=True)
    product_id = Column(Integer, ForeignKey("products.id"), nullable=False)
    material_id = Column(Integer, ForeignKey("materials.id"), nullable=False)
    site_id = Column(Integer, ForeignKey("sites.id"), nullable=False)
    old_qty = Column(Float, nullable=False)
    new_qty = Column(Float, nullable=False)
    old_version = Column(Integer, nullable=True)
    new_version = Column(Integer, nullable=True)
    action = Column(String, default="correct")
    reason = Column(Text, nullable=True)
    changed_by = Column(String, nullable=True)
    changed_at = Column(DateTime, default=_now)

class KanbanProfile(Base):
    __tablename__ = "kanban_profiles"
    __table_args__ = (UniqueConstraint("product_id", "site_id"),)
    id = Column(Integer, primary_key=True)
    product_id = Column(Integer, ForeignKey("products.id"), nullable=False)
    site_id = Column(Integer, ForeignKey("sites.id"), nullable=False)
    lead_time_days = Column(Float, default=1.0)
    container_capacity = Column(Float, default=1.0)
    safety_factor_pct = Column(Float, default=20.0)

class KanbanLoop(Base):
    __tablename__ = "kanban_loops"
    id = Column(Integer, primary_key=True)
    name = Column(String, nullable=False)
    loop_type = Column(String, default="replenishment")
    from_site_id = Column(Integer, ForeignKey("sites.id"), nullable=True)
    to_site_id = Column(Integer, ForeignKey("sites.id"), nullable=True)
    supplier_name = Column(String, nullable=True)
    notes = Column(Text, nullable=True)

class KanbanCard(Base):
    __tablename__ = "kanban_cards"
    id = Column(Integer, primary_key=True)
    card_number = Column(String, unique=True, nullable=False)
    material_id = Column(Integer, ForeignKey("materials.id"), nullable=False)
    site_id = Column(Integer, ForeignKey("sites.id"), nullable=False)
    loop_id = Column(Integer, ForeignKey("kanban_loops.id"), nullable=True)
    container_capacity = Column(Float, nullable=False)
    packaging_id = Column(Integer, ForeignKey("packagings.id"), nullable=True)
    packaging_name = Column(String, nullable=True)
    packaging_capacity_native = Column(Float, nullable=True)
    packaging_native_unit = Column(String, nullable=True)
    reorder_point = Column(Float, nullable=False)
    supplier = Column(String, nullable=True)
    part_number = Column(String, nullable=True)
    is_active = Column(Boolean, default=True)
    created_at = Column(DateTime, default=_now)

class StockLevel(Base):
    __tablename__ = "stock_levels"
    __table_args__ = (UniqueConstraint("material_id", "site_id"),)
    id = Column(Integer, primary_key=True)
    material_id = Column(Integer, ForeignKey("materials.id"), nullable=False)
    site_id = Column(Integer, ForeignKey("sites.id"), nullable=False)
    quantity = Column(Float, default=0.0, nullable=False)
    updated_at = Column(DateTime, default=_now, onupdate=_now)

class StockMovement(Base):
    __tablename__ = "stock_movements"
    id = Column(Integer, primary_key=True)
    material_id = Column(Integer, ForeignKey("materials.id"), nullable=False)
    site_id = Column(Integer, ForeignKey("sites.id"), nullable=False)
    movement_type = Column(String, nullable=False) # receipt|issue|return|adjustment
    quantity = Column(Float, nullable=False)
    quantity_before = Column(Float, nullable=False)
    quantity_after = Column(Float, nullable=False)
    packaging_id = Column(Integer, ForeignKey("packagings.id"), nullable=True)
    packaging_qty = Column(Float, nullable=True)
    expected_qty = Column(Float, nullable=True)
    delta_qty = Column(Float, nullable=True)
    reason = Column(Text, nullable=True)
    performed_by = Column(String, nullable=True)
    created_at = Column(DateTime, default=_now)

Base.metadata.create_all(bind=engine)

# ============================================================
#  СХЕМЫ ДАННЫХ (PYDANTIC)
# ============================================================
class SiteIn(BaseModel):
    name: str
    code: str
    site_type: str = "production"

class MaterialIn(BaseModel):
    code: str
    name: str
    unit: str = "кг"
    kind: str = "main"
    density_kg_m3: Optional[float] = None

class ProductIn(BaseModel):
    code: str
    name: str

class PackagingIn(BaseModel):
    code: str
    name: str
    packaging_type: str
    capacity: float = Field(..., gt=0)
    capacity_unit: str
    tare_weight_kg: Optional[float] = None
    net_weight_kg: Optional[float] = None
    is_returnable: bool = False
    notes: Optional[str] = None
    parent_id: Optional[int] = None
    children_per_unit: Optional[float] = None

class MaterialPackagingIn(BaseModel):
    material_id: int
    packaging_id: int
    is_primary: bool = False
    supplier: Optional[str] = None
    part_number: Optional[str] = None
    conversion_notes: Optional[str] = None

class BOMLineIn(BaseModel):
    product_id: int
    material_id: int
    site_id: int
    qty_per_unit: float = Field(..., ge=0)

class BOMCorrect(BaseModel):
    qty_per_unit: float = Field(..., ge=0)
    reason: Optional[str] = None
    changed_by: Optional[str] = None

class BOMBulkCorrectItem(BaseModel):
    bom_line_id: int
    qty_per_unit: float = Field(..., ge=0)

class BOMBulkCorrect(BaseModel):
    items: List[BOMBulkCorrectItem]
    reason: Optional[str] = None
    changed_by: Optional[str] = None

class KanbanProfileIn(BaseModel):
    product_id: int
    site_id: int
    lead_time_days: float = 1.0
    container_capacity: float = 1.0
    safety_factor_pct: float = 20.0

class PlanItem(BaseModel):
    product_id: int
    site_id: int
    quantity: float

class KanbanLoopIn(BaseModel):
    name: str
    loop_type: str = "replenishment"
    from_site_id: Optional[int] = None
    to_site_id: Optional[int] = None
    supplier_name: Optional[str] = None
    notes: Optional[str] = None

class GenerateCardsIn(BaseModel):
    material_id: int
    site_id: int
    loop_id: Optional[int] = None
    packaging_id: Optional[int] = None
    D: float = Field(..., gt=0)
    L: float = Field(..., gt=0)
    S: float = Field(..., ge=0, le=1)
    C: Optional[float] = Field(None, gt=0)
    supplier: Optional[str] = None
    part_number: Optional[str] = None

class SendCardIn(BaseModel):
    to_email: str
    subject: Optional[str] = None
    message: Optional[str] = None

class ConvertIn(BaseModel):
    qty: float = Field(..., ge=0)
    from_unit: str
    to_unit: str

class DensityIn(BaseModel):
    density_kg_m3: float = Field(..., gt=0)

class ExpandIn(BaseModel):
    qty: float = Field(..., gt=0)

class ConvertPackagingIn(BaseModel):
    qty: float = Field(..., gt=0)
    to_packaging_id: int

class StockOperationIn(BaseModel):
    material_id: int
    site_id: int
    packaging_id: Optional[int] = None
    packaging_qty: Optional[float] = None
    actual_base_qty: Optional[float] = Field(None, ge=0)
    reason: Optional[str] = None
    performed_by: Optional[str] = None

# ============================================================
#  ПРИЛОЖЕНИЕ И ХЕЛПЕРЫ
# ============================================================
app = FastAPI(title="Kanban Production Calculator", version="3.0")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()

# --- Шрифты для PDF ---
def _register_card_font() -> str:
    candidates = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "C:/Windows/Fonts/Arial.ttf",
        "/System/Library/Fonts/Supplemental/Arial.ttf",
    ]
    for path in candidates:
        if os.path.exists(path):
            try:
                pdfmetrics.registerFont(TTFont("CardFont", path))
                return "CardFont"
            except Exception:
                continue
    return "Helvetica"
_CARD_FONT = _register_card_font()

# --- Конвертация единиц ---
_UNIT_ALIASES = {"m3": "м³", "м3": "м³", "куб": "м³", "l": "л", "л": "л", "kg": "кг", "кг": "кг", "шт": "шт", "pcs": "шт"}
def _norm_unit(u: str) -> str:
    return _UNIT_ALIASES.get((u or "").strip().lower().replace(" ", ""), (u or "").strip().lower())

def convert_qty(qty: float, from_unit: str, to_unit: str, material) -> float:
    fu, tu = _norm_unit(from_unit), _norm_unit(to_unit)
    if fu == tu: return qty
    if fu == "м³" and tu == "л": return qty * 1000
    if fu == "л" and tu == "м³": return qty / 1000
    density = getattr(material, "density_kg_m3", None)
    needs_density = (fu in {"м³", "л"} and tu == "кг") or (fu == "кг" and tu in {"м³", "л"})
    if needs_density and not density: raise HTTPException(400, "Задайте плотность материала (density_kg_m3)")
    if fu == "м³" and tu == "кг": return qty * density
    if fu == "л" and tu == "кг": return qty * density / 1000.0
    if fu == "кг" and tu == "м³": return qty / density
    if fu == "кг" and tu == "л": return qty / density * 1000.0
    raise HTTPException(400, f"Конвертация {from_unit}->{to_unit} невозможна")

def packaging_capacity_in_base(packaging, material) -> float:
    return convert_qty(packaging.capacity, packaging.capacity_unit, material.unit, material)

# --- Иерархия тары ---
def packaging_children(db: Session, parent_id: int) -> List[Packaging]:
    return db.query(Packaging).filter(Packaging.parent_id == parent_id).all()

def packaging_tree(db: Session, packaging: Packaging, qty: float) -> dict:
    levels = []
    def walk(pkg: Packaging, count: float):
        levels.append({"packaging_id": pkg.id, "code": pkg.code, "name": pkg.name, "type": pkg.packaging_type, "quantity": round(count, 4), "native_capacity": f"{pkg.capacity:g} {pkg.capacity_unit}", "children_per_unit": pkg.children_per_unit})
        for child in packaging_children(db, pkg.id):
            if child.children_per_unit: walk(child, count * child.children_per_unit)
    walk(packaging, qty)
    return {"root": packaging.code, "root_qty": qty, "levels": levels}

def expand_to_base(db: Session, packaging: Packaging, qty: float, material: Material) -> float:
    return qty * convert_qty(packaging.capacity, packaging.capacity_unit, material.unit, material)

def convert_packaging(db: Session, from_pkg: Packaging, qty: float, to_pkg: Packaging, material: Material) -> float:
    from_base = expand_to_base(db, from_pkg, qty, material)
    to_cap = convert_qty(to_pkg.capacity, to_pkg.capacity_unit, material.unit, material)
    return from_base / to_cap if to_cap else 0.0

# --- Склад ---
def _get_or_create_stock(db: Session, material_id: int, site_id: int) -> StockLevel:
    level = db.query(StockLevel).filter(StockLevel.material_id == material_id, StockLevel.site_id == site_id).first()
    if not level:
        level = StockLevel(material_id=material_id, site_id=site_id, quantity=0.0)
        db.add(level); db.flush()
    return level

def _record_movement(db: Session, *, material_id: int, site_id: int, movement_type: str, quantity_signed: float, packaging_id: Optional[int] = None, packaging_qty: Optional[float] = None, expected_qty: Optional[float] = None, delta_qty: Optional[float] = None, reason: Optional[str] = None, performed_by: Optional[str] = None) -> StockMovement:
    level = _get_or_create_stock(db, material_id, site_id)
    before = level.quantity
    after = before + quantity_signed
    if after < -1e-9: raise HTTPException(400, f"Недостаточно остатка: {before:.4f}, нужно {-quantity_signed:.4f}")
    level.quantity = after
    mv = StockMovement(material_id=material_id, site_id=site_id, movement_type=movement_type, quantity=quantity_signed, quantity_before=before, quantity_after=after, packaging_id=packaging_id, packaging_qty=packaging_qty, expected_qty=expected_qty, delta_qty=delta_qty, reason=reason, performed_by=performed_by)
    db.add(mv)
    return mv

def _resolve_stock_qty(db: Session, material: Material, p: StockOperationIn):
    packaging = db.query(Packaging).get(p.packaging_id) if p.packaging_id else None
    expected = expand_to_base(db, packaging, p.packaging_qty, material) if packaging and p.packaging_qty is not None else None
    actual = p.actual_base_qty if p.actual_base_qty is not None else expected
    if actual is None: raise HTTPException(400, "Укажите packaging_qty или actual_base_qty")
    delta = (actual - expected) if expected is not None else None
    return packaging, expected, actual, delta

# --- BOM ---
def _current_bom(db, product_id, site_id, at: Optional[datetime] = None):
    at = at or _now()
    return db.query(BOMLine).filter(BOMLine.product_id == product_id, BOMLine.site_id == site_id, BOMLine.valid_from <= at, or_(BOMLine.valid_to.is_(None), BOMLine.valid_to > at)).all()

def _apply_correction(db: Session, line: BOMLine, new_qty: float, reason: Optional[str], changed_by: Optional[str], action: str = "correct"):
    now = _now()
    old_qty, old_ver = line.qty_per_unit, line.version
    line.valid_to = now
    new_line = BOMLine(product_id=line.product_id, material_id=line.material_id, site_id=line.site_id, qty_per_unit=new_qty, version=old_ver + 1, valid_from=now)
    db.add(new_line); db.flush()
    db.add(BOMHistory(product_id=line.product_id, material_id=line.material_id, site_id=line.site_id, old_qty=old_qty, new_qty=new_qty, old_version=old_ver, new_version=new_line.version, reason=reason, changed_by=changed_by, action=action))
    return new_line

def _requirements(db: Session, plan: List[PlanItem], at: Optional[datetime] = None):
    at = at or _now()
    result = {}
    for item in plan:
        lines = _current_bom(db, item.product_id, item.site_id, at)
        for line in lines:
            result.setdefault(item.site_id, {})
            result[item.site_id][line.material_id] = result[item.site_id].get(line.material_id, 0.0) + item.quantity * line.qty_per_unit
    return result

# --- Канбан формула ---
def compute_kanban_levels(D: float, L: float, S: float, C: float) -> dict:
    if C <= 0 or D <= 0 or L <= 0: raise HTTPException(400, "D, L, C должны быть > 0")
    demand_lt = D * L
    safety_stock = demand_lt * S
    reorder_point = demand_lt * (1 + S)
    n_raw = reorder_point / C
    N = ceil(n_raw)
    max_stock = N * C
    return {"inputs": {"D_daily_demand": D, "L_lead_time_days": L, "S_safety_factor": S, "C_container_capacity": C}, "levels": {"demand_during_lead_time": round(demand_lt, 4), "min_safety_stock": round(safety_stock, 4), "reorder_point": round(reorder_point, 4), "max_stock": round(max_stock, 4)}, "cards": {"N_raw": round(n_raw, 4), "N_rounded_up": N}}

# --- Карточки (PDF/HTML) ---
def _packaging_icon(packaging_type: Optional[str]) -> str:
    return {"box": "📦", "pallet": "🎯", "barrel": "🛢", "ibc": "🚰", "tank": "🏭", "bag": "🧂", "roll": "🎞", "canister": "🪣"}.get((packaging_type or "").lower(), "📦")

def _packaging_label_ru(packaging_type: Optional[str]) -> str:
    return {"box": "КОРОБКА", "pallet": "ПОДДОН", "barrel": "БОЧКА", "ibc": "IBC", "tank": "ЁМКОСТЬ", "bag": "МЕШОК", "roll": "РУЛОН", "canister": "КАНИСТРА"}.get((packaging_type or "").lower(), "ТАРА")

def _card_data(db: Session, card: KanbanCard) -> dict:
    mat = db.query(Material).get(card.material_id)
    site = db.query(Site).get(card.site_id)
    loop = db.query(KanbanLoop).get(card.loop_id) if card.loop_id else None
    packaging = db.query(Packaging).get(card.packaging_id) if card.packaging_id else None

    packaging_label = None
    if card.packaging_name and card.packaging_capacity_native:
        packaging_label = f"{card.packaging_name} ({card.packaging_capacity_native:g} {card.packaging_native_unit}"
        if packaging and packaging.packaging_type == "bag" and packaging.net_weight_kg:
            packaging_label += f", нетто {packaging.net_weight_kg:g} кг"
        if mat.density_kg_m3 and card.packaging_native_unit in ("л", "м³"):
            packaging_label += f" × {mat.density_kg_m3} кг/м³ = {card.container_capacity:g} {mat.unit}"
        packaging_label += ")"

    return {
        "card_id": card.id, "card_number": card.card_number, "title": mat.name, "code": mat.code,
        "kind": "Основное" if mat.kind == "main" else "Вспомогательное",
        "quantity": card.container_capacity, "unit": mat.unit,
        "packaging": packaging_label, "packaging_type": packaging.packaging_type if packaging else None,
        "packaging_icon": _packaging_icon(packaging.packaging_type if packaging else None),
        "packaging_type_label": _packaging_label_ru(packaging.packaging_type if packaging else None),
        "reorder_point": card.reorder_point, "supplier": card.supplier or "—", "part_number": card.part_number or "—",
        "site_code": site.code, "site_name": site.name, "loop_name": loop.name if loop else "—",
        "is_active": card.is_active, "created_at": card.created_at,
        "links": {"pdf": f"/kanban/card/{card.id}/pdf/", "print": f"/kanban/card/{card.id}/print/", "send": f"/kanban/card/{card.id}/send/"}
    }

def _draw_card_on_canvas(c, data: dict, w_pt: float, h_pt: float):
    dark, accent, soft = HexColor("#2c3e50"), HexColor("#c0392b"), HexColor("#7f8c8d")
    c.setStrokeColor(dark); c.setLineWidth(2); c.rect(6, 6, w_pt - 12, h_pt - 12)
    c.setFillColor(dark); c.rect(6, h_pt - 32, w_pt - 12, 26, fill=1, stroke=0)
    c.setFillColor(HexColor("#ffffff")); c.setFont(_CARD_FONT, 9); c.drawString(12, h_pt - 18, "КАНБАН-КАРТОЧКА")
    c.setFont(_CARD_FONT, 8); c.setFillColor(HexColor("#e67e22")); c.drawString(12, h_pt - 28, f"[{data.get('packaging_type_label', 'ТАРА')}]")
    c.setFillColor(HexColor("#ffffff")); c.setFont(_CARD_FONT, 10); c.drawRightString(w_pt - 12, h_pt - 18, data["card_number"])
    c.setFillColor(dark); c.setFont(_CARD_FONT, 13); c.drawString(12, h_pt - 52, data["title"][:42])
    c.setFillColor(soft); c.setFont(_CARD_FONT, 8); c.drawString(12, h_pt - 64, f"{data['code']}  •  {data['kind']}")
    c.setFillColor(accent); c.setFont(_CARD_FONT, 24); c.drawRightString(w_pt - 12, h_pt - 56, f"{data['quantity']:g}")
    c.setFillColor(soft); c.setFont(_CARD_FONT, 8); c.drawRightString(w_pt - 12, h_pt - 68, data["unit"])
    c.setStrokeColor(soft); c.setLineWidth(0.5); c.line(12, h_pt - 76, w_pt - 12, h_pt - 76)
    y = h_pt - 86
    rows = [("Точка заказа (ROP)", f"{data['reorder_point']:g} {data['unit']}"), ("Тара", data["packaging"] or "—"), ("Поставщик", data["supplier"]), ("Артикул", data["part_number"]), ("Участок", f"{data['site_code']} — {data['site_name']}"), ("Петля", data["loop_name"])]
    for label, value in rows:
        c.setFillColor(soft); c.setFont(_CARD_FONT, 7.5); c.drawString(12, y, label)
        c.setFillColor(dark); c.setFont(_CARD_FONT, 8); c.drawString(90, y, str(value)[:50]); y -= 10.5
    try: code128.Code128(data["card_number"], barWidth=0.9, barHeight=28).drawOn(c, 12, 16)
    except Exception: pass
    c.setFillColor(soft); c.setFont(_CARD_FONT, 6)
    if data.get("created_at"): c.drawRightString(w_pt - 12, 14, f"создана: {data['created_at'].strftime('%d.%m.%Y')}")

def _build_card_pdf(data: dict, size_mm=(148, 105)) -> bytes:
    w_pt, h_pt = size_mm[0] * mm, size_mm[1] * mm
    buf = io.BytesIO(); c = rl_canvas.Canvas(buf, pagesize=(w_pt, h_pt))
    _draw_card_on_canvas(c, data, w_pt, h_pt); c.showPage(); c.save(); buf.seek(0)
    return buf.getvalue()

def _build_card_html(data: dict, auto_print: bool = True) -> str:
    auto_js = "setTimeout(() => window.print(), 400);" if auto_print else ""
    return f"""<!doctype html><html lang="ru"><head><meta charset="utf-8"><title>Канбан-карточка {data['card_number']}</title>
<style>@page {{ size: 148mm 105mm; margin: 0; }} * {{ box-sizing: border-box; }} body {{ margin: 0; font-family: Arial, sans-serif; background: #ecf0f1; display: flex; justify-content: center; padding: 20px; }} .card {{ width: 148mm; min-height: 105mm; background: #fff; border: 2px solid #2c3e50; border-radius: 6px; padding: 6mm; display: flex; flex-direction: column; box-shadow: 0 4px 16px rgba(0,0,0,.15); position: relative; }} .head {{ background: #2c3e50; color: #fff; padding: 6px 10px; margin: -6mm -6mm 4mm -6mm; border-radius: 4px 4px 0 0; display: flex; justify-content: space-between; font-size: 11px; }} .title {{ font-size: 18px; font-weight: bold; color: #2c3e50; padding-right: 80px; }} .sub {{ color: #7f8c8d; font-size: 11px; margin-bottom: 3mm; }} .qty {{ position: absolute; top: 12mm; right: 6mm; font-size: 30px; color: #c0392b; font-weight: bold; line-height: 1; text-align: right; }} .qty-unit {{ position: absolute; top: 22mm; right: 6mm; font-size: 10px; color: #7f8c8d; text-align: right; }} table {{ width: 100%; border-collapse: collapse; font-size: 11px; }} td {{ padding: 1.5px 0; vertical-align: top; }} td.l {{ color: #7f8c8d; width: 40%; }} td.v {{ color: #2c3e50; font-weight: 500; }} .footer {{ margin-top: auto; display: flex; justify-content: space-between; font-size: 8px; color: #7f8c8d; border-top: 1px solid #ddd; padding-top: 2mm; }} .toolbar {{ position: fixed; top: 10px; right: 10px; display: flex; gap: 8px; z-index: 10; }} .btn {{ padding: 8px 16px; background: #2980b9; color: #fff; border: none; border-radius: 4px; cursor: pointer; font-size: 13px; text-decoration: none; }} .btn:hover {{ background: #21618c; }} @media print {{ body {{ background: #fff; padding: 0; }} .toolbar {{ display: none; }} .card {{ box-shadow: none; border-radius: 0; }} }}</style></head><body>
<div class="toolbar"><button class="btn" onclick="window.print()">🖨 Печать</button><a class="btn" href="/kanban/card/{data['card_id']}/pdf/">⬇ PDF</a></div>
<div class="card"><div class="head"><span>КАНБАН-КАРТОЧКА [{data.get('packaging_type_label', 'ТАРА')}]</span><span>{data['card_number']}</span></div><div class="qty">{data['quantity']:g}</div><div class="qty-unit">{data['unit']}</div><div class="title">{data['title']}</div><div class="sub">{data['code']} • {data['kind']}</div>
<table><tr><td class="l">Точка заказа:</td><td class="v">{data['reorder_point']:g} {data['unit']}</td></tr><tr><td class="l">Тара:</td><td class="v"><span style="font-size:14px">{data.get('packaging_icon','📦')}</span> {data['packaging'] or '—'}</td></tr><tr><td class="l">Поставщик:</td><td class="v">{data['supplier']}</td></tr><tr><td class="l">Артикул:</td><td class="v">{data['part_number']}</td></tr><tr><td class="l">Участок:</td><td class="v">{data['site_code']} — {data['site_name']}</td></tr><tr><td class="l">Петля:</td><td class="v">{data['loop_name']}</td></tr></table>
<div class="footer"><span>Создана: {data['created_at'].strftime('%d.%m.%Y') if data.get('created_at') else '—'}</span><span>Вытягивающая система • Just-in-Time</span></div></div>
<script>{auto_js}</script></body></html>"""

def _send_email_with_attachment(*, to_email, subject, body, attachment_bytes, attachment_name):
    host, port, user, pwd = os.getenv("SMTP_HOST"), int(os.getenv("SMTP_PORT", "587")), os.getenv("SMTP_USER"), os.getenv("SMTP_PASSWORD")
    sender, use_tls = os.getenv("SMTP_FROM", user), os.getenv("SMTP_USE_TLS", "1") == "1"
    if not (host and user and pwd): raise HTTPException(503, "SMTP не настроен. Задайте SMTP_HOST, SMTP_USER, SMTP_PASSWORD.")
    msg = EmailMessage(); msg["From"] = sender; msg["To"] = to_email; msg["Subject"] = subject; msg.set_content(body)
    msg.add_attachment(attachment_bytes, maintype="application", subtype="pdf", filename=attachment_name)
    try:
        if use_tls:
            with smtplib.SMTP(host, port, timeout=20) as smtp: smtp.starttls(); smtp.login(user, pwd); smtp.send_message(msg)
        else:
            with smtplib.SMTP_SSL(host, port, timeout=20) as smtp: smtp.login(user, pwd); smtp.send_message(msg)
    except smtplib.SMTPException as e: raise HTTPException(502, f"SMTP-ошибка: {e}")

# ============================================================
#  ЭНДПОИНТЫ
# ============================================================

@app.post("/sites/")
def create_site(p: SiteIn, db: Session = Depends(get_db)):
    s = Site(**p.dict()); db.add(s); db.commit(); db.refresh(s); return s

@app.get("/sites/")
def list_sites(db: Session = Depends(get_db)): return db.query(Site).all()

@app.post("/materials/")
def create_material(p: MaterialIn, db: Session = Depends(get_db)):
    m = Material(**p.dict()); db.add(m); db.commit(); db.refresh(m); return m

@app.post("/products/")
def create_product(p: ProductIn, db: Session = Depends(get_db)):
    pr = Product(**p.dict()); db.add(pr); db.commit(); db.refresh(pr); return pr

@app.patch("/materials/{material_id}/density/")
def set_density(material_id: int, p: DensityIn, db: Session = Depends(get_db)):
    m = db.query(Material).get(material_id)
    if not m: raise HTTPException(404, "Материал не найден")
    m.density_kg_m3 = p.density_kg_m3; db.commit(); return {"status": "updated"}

@app.post("/materials/{material_id}/convert/")
def convert_units(material_id: int, p: ConvertIn, db: Session = Depends(get_db)):
    m = db.query(Material).get(material_id)
    if not m: raise HTTPException(404, "Материал не найден")
    return {"output": {"qty": round(convert_qty(p.qty, p.from_unit, p.to_unit, m), 6), "unit": p.to_unit}}

@app.post("/packagings/")
def create_packaging(p: PackagingIn, db: Session = Depends(get_db)):
    pkg = Packaging(**p.dict()); db.add(pkg); db.commit(); db.refresh(pkg); return pkg

@app.get("/packagings/")
def list_packagings(db: Session = Depends(get_db)): return db.query(Packaging).all()

@app.post("/material-packagings/")
def link_material_packaging(p: MaterialPackagingIn, db: Session = Depends(get_db)):
    link = MaterialPackaging(**p.dict()); db.add(link); db.commit(); db.refresh(link); return link

@app.get("/packagings/{pkg_id}/hierarchy/")
def get_packaging_hierarchy(pkg_id: int, qty: float = Query(1, gt=0), db: Session = Depends(get_db)):
    pkg = db.query(Packaging).get(pkg_id)
    if not pkg: raise HTTPException(404, "Тара не найдена")
    return packaging_tree(db, pkg, qty)

@app.post("/packagings/{pkg_id}/expand/")
def expand_packaging(pkg_id: int, p: ExpandIn, material_id: int = Query(...), db: Session = Depends(get_db)):
    pkg, mat = db.query(Packaging).get(pkg_id), db.query(Material).get(material_id)
    if not pkg or not mat: raise HTTPException(404, "Не найдено")
    tree = packaging_tree(db, pkg, p.qty); tree["base_quantity"] = round(expand_to_base(db, pkg, p.qty, mat), 4); tree["base_unit"] = mat.unit
    return tree

@app.post("/packagings/{pkg_id}/convert/")
def convert_packaging_endpoint(pkg_id: int, p: ConvertPackagingIn, material_id: int = Query(...), db: Session = Depends(get_db)):
    from_pkg, to_pkg, mat = db.query(Packaging).get(pkg_id), db.query(Packaging).get(p.to_packaging_id), db.query(Material).get(material_id)
    if not all([from_pkg, to_pkg, mat]): raise HTTPException(404, "Не найдено")
    return {"to_qty": round(convert_packaging(db, from_pkg, p.qty, to_pkg, mat), 4)}

@app.post("/bom/")
def create_bom_line(p: BOMLineIn, db: Session = Depends(get_db)):
    line = BOMLine(**p.dict(), version=1, valid_from=_now()); db.add(line); db.flush()
    db.add(BOMHistory(product_id=p.product_id, material_id=p.material_id, site_id=p.site_id, old_qty=0, new_qty=p.qty_per_unit, new_version=1, action="create"))
    db.commit(); db.refresh(line); return line

@app.get("/bom/")
def list_bom(site_id: Optional[int] = None, product_id: Optional[int] = None, db: Session = Depends(get_db)):
    q = db.query(BOMLine).filter(BOMLine.valid_to.is_(None))
    if site_id: q = q.filter(BOMLine.site_id == site_id)
    if product_id: q = q.filter(BOMLine.product_id == product_id)
    return q.all()

@app.get("/bom/product/{product_id}/")
def get_product_bom(product_id: int, site_id: Optional[int] = None, at: Optional[datetime] = None, db: Session = Depends(get_db)):
    return _current_bom(db, product_id, site_id, at)

@app.patch("/bom/{bom_id}/")
def correct_norm(bom_id: int, p: BOMCorrect, db: Session = Depends(get_db)):
    line = db.query(BOMLine).get(bom_id)
    if not line: raise HTTPException(404, "Не найдено")
    if line.valid_to is not None:
        line = db.query(BOMLine).filter(BOMLine.product_id==line.product_id, BOMLine.material_id==line.material_id, BOMLine.site_id==line.site_id, BOMLine.valid_to.is_(None)).first()
        if not line: raise HTTPException(400, "Устаревшая норма")
    if abs(line.qty_per_unit - p.qty_per_unit) < 1e-9: return {"status": "no_change"}
    new_line = _apply_correction(db, line, p.qty_per_unit, p.reason, p.changed_by)
    db.commit(); return {"status": "updated", "version": new_line.version}

@app.get("/bom/history/")
def get_bom_history(product_id: Optional[int] = None, material_id: Optional[int] = None, limit: int = 200, db: Session = Depends(get_db)):
    q = db.query(BOMHistory)
    if product_id: q = q.filter(BOMHistory.product_id == product_id)
    if material_id: q = q.filter(BOMHistory.material_id == material_id)
    return q.order_by(BOMHistory.changed_at.desc()).limit(limit).all()

@app.post("/stock/receive/")
def stock_receive(p: StockOperationIn, db: Session = Depends(get_db)):
    mat = db.query(Material).get(p.material_id)
    if not mat: raise HTTPException(404, "Не найдено")
    _, expected, actual, delta = _resolve_stock_qty(db, mat, p)
    mv = _record_movement(db, material_id=p.material_id, site_id=p.site_id, movement_type="receipt", quantity_signed=actual, packaging_id=p.packaging_id, packaging_qty=p.packaging_qty, expected_qty=expected, delta_qty=delta, reason=p.reason, performed_by=p.performed_by)
    db.commit(); return {"status": "ok", "quantity_after": mv.quantity_after}

@app.post("/stock/issue/")
def stock_issue(p: StockOperationIn, db: Session = Depends(get_db)):
    mat = db.query(Material).get(p.material_id)
    if not mat: raise HTTPException(404, "Не найдено")
    _, expected, actual, delta = _resolve_stock_qty(db, mat, p)
    mv = _record_movement(db, material_id=p.material_id, site_id=p.site_id, movement_type="issue", quantity_signed=-actual, packaging_id=p.packaging_id, packaging_qty=p.packaging_qty, expected_qty=expected, delta_qty=delta, reason=p.reason, performed_by=p.performed_by)
    db.commit(); return {"status": "ok", "quantity_after": mv.quantity_after}

@app.post("/stock/return/")
def stock_return(p: StockOperationIn, db: Session = Depends(get_db)):
    mat = db.query(Material).get(p.material_id)
    if not mat: raise HTTPException(404, "Не найдено")
    _, expected, actual, delta = _resolve_stock_qty(db, mat, p)
    mv = _record_movement(db, material_id=p.material_id, site_id=p.site_id, movement_type="return", quantity_signed=actual, packaging_id=p.packaging_id, packaging_qty=p.packaging_qty, expected_qty=expected, delta_qty=delta, reason=p.reason, performed_by=p.performed_by)
    db.commit(); return {"status": "returned", "quantity_after": mv.quantity_after, "delta": delta}

@app.post("/stock/adjust/")
def stock_adjust(p: StockOperationIn, db: Session = Depends(get_db)):
    if p.actual_base_qty is None: raise HTTPException(400, "Нужен actual_base_qty")
    level = _get_or_create_stock(db, p.material_id, p.site_id)
    delta = p.actual_base_qty - level.quantity
    mv = _record_movement(db, material_id=p.material_id, site_id=p.site_id, movement_type="adjustment", quantity_signed=delta, reason=p.reason)
    db.commit(); return {"status": "adjusted", "delta": delta, "quantity_after": mv.quantity_after}

@app.get("/stock/levels/")
def stock_levels(site_id: Optional[int] = None, material_id: Optional[int] = None, db: Session = Depends(get_db)):
    q = db.query(StockLevel, Material, Site).join(Material, StockLevel.material_id==Material.id).join(Site, StockLevel.site_id==Site.id)
    if site_id: q = q.filter(StockLevel.site_id==site_id)
    if material_id: q = q.filter(StockLevel.material_id==material_id)
    return [{"material": m.code, "site": s.code, "quantity": lv.quantity, "unit": m.unit} for lv, m, s in q.all()]

@app.post("/kanban/calc/")
def kanban_calc(D: float, L: float, S: float, C: float):
    return compute_kanban_levels(D, L, S, C)

@app.post("/kanban/generate-cards/")
def generate_cards(p: GenerateCardsIn, db: Session = Depends(get_db)):
    mat, site = db.query(Material).get(p.material_id), db.query(Site).get(p.site_id)
    if not mat or not site: raise HTTPException(404, "Не найдено")
    packaging = db.query(Packaging).get(p.packaging_id) if p.packaging_id else None
    if packaging: C = packaging_capacity_in_base(packaging, mat)
    elif p.C: C = p.C
    else: raise HTTPException(400, "Укажите C или packaging_id")
    
    calc = compute_kanban_levels(p.D, p.L, p.S, C); N = calc["cards"]["N_rounded_up"]; ROP = calc["levels"]["reorder_point"]
    for c in db.query(KanbanCard).filter(KanbanCard.material_id == p.material_id, KanbanCard.site_id == p.site_id).all(): c.is_active = False
    for i in range(1, N + 1):
        db.add(KanbanCard(card_number=f"KBN-{mat.code}-{site.code}-{i:03d}", material_id=p.material_id, site_id=p.site_id, loop_id=p.loop_id, container_capacity=C, reorder_point=ROP, packaging_id=packaging.id if packaging else None, packaging_name=packaging.name if packaging else None, packaging_capacity_native=packaging.capacity if packaging else None, packaging_native_unit=packaging.capacity_unit if packaging else None, supplier=p.supplier, part_number=p.part_number))
    db.commit(); return {"cards_created": N, "cards": calc}

@app.get("/kanban/card/{card_id}/")
def get_card(card_id: int, db: Session = Depends(get_db)):
    card = db.query(KanbanCard).get(card_id)
    if not card: raise HTTPException(404)
    return _card_data(db, card)

@app.get("/kanban/card/{card_id}/pdf/")
def download_card_pdf(card_id: int, db: Session = Depends(get_db)):
    card = db.query(KanbanCard).get(card_id)
    if not card: raise HTTPException(404)
    return Response(content=_build_card_pdf(_card_data(db, card)), media_type="application/pdf", headers={"Content-Disposition": f'attachment; filename="{card.card_number}.pdf"'})

@app.get("/kanban/card/{card_id}/print/")
def print_card(card_id: int, auto: bool = True, db: Session = Depends(get_db)):
    card = db.query(KanbanCard).get(card_id)
    if not card: raise HTTPException(404)
    return Response(content=_build_card_html(_card_data(db, card), auto_print=auto), media_type="text/html; charset=utf-8")

@app.post("/kanban/card/{card_id}/send/")
def send_card_email(card_id: int, p: SendCardIn, db: Session = Depends(get_db)):
    card = db.query(KanbanCard).get(card_id)
    if not card: raise HTTPException(404)
    data = _card_data(db, card); pdf_bytes = _build_card_pdf(data)
    _send_email_with_attachment(to_email=p.to_email, subject=p.subject or f"Канбан-карточка {card.card_number}", body=p.message or "Во вложении.", attachment_bytes=pdf_bytes, attachment_name=f"{card.card_number}.pdf")
    return {"status": "sent"}

# ============================================================
#  ДЕМО-ДАННЫЕ
# ============================================================
@app.post("/seed-demo/")
def seed_demo(db: Session = Depends(get_db)):
    if db.query(Site).count() > 0: return {"status": "already seeded"}
    
    s1 = Site(name="Участок смешивания", code="S1", site_type="production")
    wh = Site(name="Главный склад", code="WH", site_type="warehouse")
    db.add_all([s1, wh]); db.flush()

    m_sugar = Material(code="SUGAR", name="Сахар", unit="кг", kind="main")
    db.add(m_sugar); db.flush()

    pkg_pal = Packaging(code="PAL-SUG", name="Паллета сахара (40 мешков)", packaging_type="pallet", capacity=1000, capacity_unit="кг")
    db.add(pkg_pal); db.flush()
    pkg_bag = Packaging(code="BAG-SUG", name="Мешок сахара (25 кг)", packaging_type="bag", capacity=25, capacity_unit="кг", parent_id=pkg_pal.id, children_per_unit=40)
    db.add(pkg_bag); db.flush()

    db.add(MaterialPackaging(material_id=m_sugar.id, packaging_id=pkg_bag.id, is_primary=True, supplier="Сахар-Опт"))
    _record_movement(db, material_id=m_sugar.id, site_id=wh.id, movement_type="receipt", quantity_signed=2000, packaging_qty=2, expected_qty=2000, reason="Старт")
    db.commit()
    return {"status": "seeded", "note": "Перейдите в /docs для тестирования"}