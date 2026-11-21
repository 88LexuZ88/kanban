# ============================================================
#  KANBAN PRODUCTION CALCULATOR (WITH UI)
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
from fastapi.responses import HTMLResponse
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
    site_type = Column(String, default="production")

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
    packaging_type = Column(String, nullable=False)
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
    movement_type = Column(String, nullable=False)
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
#  СХЕМЫ ДАННЫХ
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

class BOMLineIn(BaseModel):
    product_id: int
    material_id: int
    site_id: int
    qty_per_unit: float = Field(..., ge=0)

class BOMCorrect(BaseModel):
    qty_per_unit: float = Field(..., ge=0)
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
app = FastAPI(title="Kanban Production Calculator", version="3.1")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()

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

def compute_kanban_levels(D: float, L: float, S: float, C: float) -> dict:
    if C <= 0 or D <= 0 or L <= 0: raise HTTPException(400, "D, L, C должны быть > 0")
    demand_lt = D * L
    safety_stock = demand_lt * S
    reorder_point = demand_lt * (1 + S)
    n_raw = reorder_point / C
    N = ceil(n_raw)
    max_stock = N * C
    return {"inputs": {"D": D, "L": L, "S": S, "C": C}, "levels": {"demand_during_lead_time": round(demand_lt, 4), "min_safety_stock": round(safety_stock, 4), "reorder_point": round(reorder_point, 4), "max_stock": round(max_stock, 4)}, "cards": {"N_raw": round(n_raw, 4), "N_rounded_up": N}}

def _packaging_icon(packaging_type: Optional[str]) -> str:
    return {"box": "📦", "pallet": "🎯", "barrel": "🛢", "ibc": "🚰", "tank": "🏭", "bag": "🧂", "roll": "🎞", "canister": "🪣"}.get((packaging_type or "").lower(), "📦")

def _packaging_label_ru(packaging_type: Optional[str]) -> str:
    return {"box": "КОРОБКА", "pallet": "ПОДДОН", "barrel": "БОЧКА", "ibc": "IBC", "tank": "ЁМКОСТЬ", "bag": "МЕШОК", "roll": "РУЛОН", "canister": "КАНИСТРА"}.get((packaging_type or "").lower(), "ТАРА")

def _card_data(db: Session, card: KanbanCard) -> dict:
    mat = db.query(Material).get(card.material_id)
    site = db.query(Site).get(card.site_id)
    loop = db.query(KanbanLoop).get(card.loop_id) if card.loop_id else None
    packaging = db.query(Packaging).get(card.packaging_id) if card.packaging_id else None
    packaging_label = f"{card.packaging_name} ({card.packaging_capacity_native:g} {card.packaging_native_unit})" if card.packaging_name else None
    return {
        "card_id": card.id, "card_number": card.card_number, "title": mat.name, "code": mat.code,
        "kind": "Основное" if mat.kind == "main" else "Вспомогательное",
        "quantity": card.container_capacity, "unit": mat.unit, "packaging": packaging_label,
        "packaging_type": packaging.packaging_type if packaging else None,
        "packaging_icon": _packaging_icon(packaging.packaging_type if packaging else None),
        "packaging_type_label": _packaging_label_ru(packaging.packaging_type if packaging else None),
        "reorder_point": card.reorder_point, "supplier": card.supplier or "—", "part_number": card.part_number or "—",
        "site_code": site.code, "site_name": site.name, "loop_name": loop.name if loop else "—",
        "created_at": card.created_at
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
    rows = [("Точка заказа (ROP)", f"{data['reorder_point']:g} {data['unit']}"), ("Тара", data["packaging"] or "—"), ("Поставщик", data["supplier"]), ("Участок", f"{data['site_code']} — {data['site_name']}")]
    for label, value in rows:
        c.setFillColor(soft); c.setFont(_CARD_FONT, 7.5); c.drawString(12, y, label)
        c.setFillColor(dark); c.setFont(_CARD_FONT, 8); c.drawString(90, y, str(value)[:50]); y -= 11
    try: code128.Code128(data["card_number"], barWidth=0.9, barHeight=28).drawOn(c, 12, 16)
    except Exception: pass

def _build_card_pdf(data: dict) -> bytes:
    w_pt, h_pt = 148 * mm, 105 * mm
    buf = io.BytesIO(); c = rl_canvas.Canvas(buf, pagesize=(w_pt, h_pt))
    _draw_card_on_canvas(c, data, w_pt, h_pt); c.showPage(); c.save(); buf.seek(0)
    return buf.getvalue()

# ============================================================
#  НОВЫЕ ЭНДПОИНТЫ ДЛЯ ИНТЕРФЕЙСА
# ============================================================
@app.get("/materials/")
def list_materials(db: Session = Depends(get_db)):
    return db.query(Material).all()

@app.get("/kanban/cards/")
def list_cards(db: Session = Depends(get_db)):
    return db.query(KanbanCard).filter(KanbanCard.is_active == True).all()

# ============================================================
#  ОСНОВНЫЕ ЭНДПОИНТЫ
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

@app.post("/stock/receive/")
def stock_receive(p: StockOperationIn, db: Session = Depends(get_db)):
    mat = db.query(Material).get(p.material_id)
    if not mat: raise HTTPException(404, "Не найдено")
    packaging = db.query(Packaging).get(p.packaging_id) if p.packaging_id else None
    expected = packaging_capacity_in_base(packaging, mat) * p.packaging_qty if packaging and p.packaging_qty else None
    actual = p.actual_base_qty if p.actual_base_qty is not None else expected
    if actual is None: raise HTTPException(400, "Ошибка ввода")
    level = db.query(StockLevel).filter_by(material_id=p.material_id, site_id=p.site_id).first()
    if not level: level = StockLevel(material_id=p.material_id, site_id=p.site_id, quantity=0); db.add(level)
    before = level.quantity; level.quantity += actual
    db.add(StockMovement(material_id=p.material_id, site_id=p.site_id, movement_type="receipt", quantity=actual, quantity_before=before, quantity_after=level.quantity, reason=p.reason))
    db.commit(); return {"status": "ok", "quantity_after": level.quantity}

@app.post("/kanban/calc/")
def kanban_calc(D: float, L: float, S: float, C: float):
    return compute_kanban_levels(D, L, S, C)

@app.post("/kanban/generate-cards/")
def generate_cards(p: GenerateCardsIn, db: Session = Depends(get_db)):
    mat, site = db.query(Material).get(p.material_id), db.query(Site).get(p.site_id)
    if not mat or not site: raise HTTPException(404, "Не найдено")
    packaging = db.query(Packaging).get(p.packaging_id) if p.packaging_id else None
    C = packaging_capacity_in_base(packaging, mat) if packaging else p.C
    if not C: raise HTTPException(400, "Укажите C или packaging_id")
    
    calc = compute_kanban_levels(p.D, p.L, p.S, C); N = calc["cards"]["N_rounded_up"]; ROP = calc["levels"]["reorder_point"]
    for c in db.query(KanbanCard).filter_by(material_id=p.material_id, site_id=p.site_id).all(): c.is_active = False
    for i in range(1, N + 1):
        db.add(KanbanCard(card_number=f"KBN-{mat.code}-{i:03d}", material_id=p.material_id, site_id=p.site_id, container_capacity=C, reorder_point=ROP, packaging_id=packaging.id if packaging else None, packaging_name=packaging.name if packaging else None, packaging_capacity_native=packaging.capacity if packaging else None, packaging_native_unit=packaging.capacity_unit if packaging else None, supplier=p.supplier))
    db.commit(); return {"cards_created": N, "calc": calc}

@app.get("/kanban/card/{card_id}/pdf/")
def download_card_pdf(card_id: int, db: Session = Depends(get_db)):
    card = db.query(KanbanCard).get(card_id)
    if not card: raise HTTPException(404)
    return Response(content=_build_card_pdf(_card_data(db, card)), media_type="application/pdf", headers={"Content-Disposition": f'attachment; filename="{card.card_number}.pdf"'})

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
    db.commit()
    return {"status": "seeded"}

# ============================================================
#  HTML ИНТЕРФЕЙС (ГЛАВНАЯ СТРАНИЦА)
# ============================================================
@app.get("/", response_class=HTMLResponse)
def index():
    return """
<!DOCTYPE html>
<html lang="ru">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Канбан Система</title>
    <style>
        :root { --primary: #2980b9; --bg: #f4f7f6; --card-bg: #ffffff; --text: #2c3e50; }
        * { box-sizing: border-box; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
        body { background-color: var(--bg); color: var(--text); margin: 0; padding: 15px; }
        .header { display: flex; justify-content: space-between; align-items: center; background: var(--primary); color: white; padding: 15px; border-radius: 8px; margin-bottom: 20px; }
        .header h1 { margin: 0; font-size: 18px; }
        .lang-switch { background: rgba(255,255,255,0.2); border: none; color: white; padding: 8px 12px; border-radius: 4px; cursor: pointer; font-weight: bold; }
        .card { background: var(--card-bg); border-radius: 8px; padding: 20px; box-shadow: 0 4px 6px rgba(0,0,0,0.05); margin-bottom: 20px; }
        h2 { font-size: 16px; margin-top: 0; border-bottom: 2px solid #eee; padding-bottom: 10px; }
        .form-group { margin-bottom: 15px; }
        label { display: block; font-size: 13px; color: #7f8c8d; margin-bottom: 5px; }
        input, select { width: 100%; padding: 12px; border: 1px solid #ccc; border-radius: 6px; font-size: 16px; }
        button.btn { width: 100%; padding: 15px; background: var(--primary); color: white; border: none; border-radius: 6px; font-size: 16px; font-weight: bold; cursor: pointer; margin-top: 10px; }
        button.btn:active { background: #21618c; }
        .result { background: #e8f4f8; padding: 15px; border-radius: 6px; margin-top: 15px; display: none; }
        .result-row { display: flex; justify-content: space-between; margin-bottom: 8px; font-size: 14px; }
        .result-row strong { font-size: 16px; color: var(--primary); }
        .card-list { display: flex; flex-direction: column; gap: 10px; }
        .card-item { border: 1px solid #eee; padding: 12px; border-radius: 6px; display: flex; justify-content: space-between; align-items: center; }
        .card-info { font-size: 14px; }
        .card-actions a { display: inline-block; padding: 8px 12px; background: #27ae60; color: white; text-decoration: none; border-radius: 4px; font-size: 12px; }
    </style>
</head>
<body>
    <div class="header">
        <h1 id="app-title">Канбан Система</h1>
        <button class="lang-switch" onclick="toggleLang()" id="lang-btn">EN</button>
    </div>

    <div class="card">
        <h2 id="calc-title">Расчет карточек</h2>
        <div class="form-group">
            <label id="lbl-mat">Материал</label>
            <select id="material_id"></select>
        </div>
        <div class="form-group">
            <label id="lbl-d">Дневная потребность (шт/кг)</label>
            <input type="number" id="D" value="400">
        </div>
        <div class="form-group">
            <label id="lbl-l">Время пополнения (дней)</label>
            <input type="number" id="L" step="0.1" value="1.5">
        </div>
        <div class="form-group">
            <label id="lbl-s">Страховой запас (%)</label>
            <input type="number" id="S" value="20">
        </div>
        <div class="form-group">
            <label id="lbl-c">Вместимость контейнера</label>
            <input type="number" id="C" value="50">
        </div>
        <button class="btn" onclick="calculate()" id="btn-calc">Рассчитать</button>
        
        <div class="result" id="result">
            <div class="result-row"><span id="res-reorder">Точка заказа:</span> <strong id="val-reorder">0</strong></div>
            <div class="result-row"><span id="res-max">Макс. запас:</span> <strong id="val-max">0</strong></div>
            <div class="result-row"><span id="res-cards">Кол-во карточек:</span> <strong id="val-cards" style="font-size: 20px; color: #c0392b;">0</strong></div>
            <button class="btn" onclick="generateCards()" style="background: #c0392b;" id="btn-gen">Создать карточки</button>
        </div>
    </div>

    <div class="card">
        <h2 id="list-title">Активные карточки</h2>
        <div class="card-list" id="cards-container">
            <!-- Заполняется через JS -->
        </div>
    </div>

<script>
    const translations = {
        ru: {
            title: "Канбан Система", calcTitle: "Расчет карточек", lblMat: "Материал",
            lblD: "Дневная потребность (шт/кг)", lblL: "Время пополнения (дней)",
            lblS: "Страховой запас (%)", lblC: "Вместимость контейнера",
            btnCalc: "Рассчитать", btnGen: "Создать карточки",
            resReorder: "Точка заказа:", resMax: "Макс. запас:", resCards: "Кол-во карточек:",
            listTitle: "Активные карточки", download: "Скачать PDF", noCards: "Нет активных карточек",
            lang: "EN"
        },
        en: {
            title: "Kanban System", calcTitle: "Card Calculation", lblMat: "Material",
            lblD: "Daily Demand (pcs/kg)", lblL: "Lead Time (days)",
            lblS: "Safety Stock (%)", lblC: "Container Capacity",
            btnCalc: "Calculate", btnGen: "Generate Cards",
            resReorder: "Reorder Point:", resMax: "Max Stock:", resCards: "Cards Qty:",
            listTitle: "Active Cards", download: "Download PDF", noCards: "No active cards",
            lang: "RU"
        }
    };

    let currentLang = 'ru';
    let currentCalcResult = null;

    function applyLang() {
        const t = translations[currentLang];
        document.getElementById('app-title').innerText = t.title;
        document.getElementById('calc-title').innerText = t.calcTitle;
        document.getElementById('lbl-mat').innerText = t.lblMat;
        document.getElementById('lbl-d').innerText = t.lblD;
        document.getElementById('lbl-l').innerText = t.lblL;
        document.getElementById('lbl-s').innerText = t.lblS;
        document.getElementById('lbl-c').innerText = t.lblC;
        document.getElementById('btn-calc').innerText = t.btnCalc;
        document.getElementById('btn-gen').innerText = t.btnGen;
        document.getElementById('res-reorder').innerText = t.resReorder;
        document.getElementById('res-max').innerText = t.resMax;
        document.getElementById('res-cards').innerText = t.resCards;
        document.getElementById('list-title').innerText = t.listTitle;
        document.getElementById('lang-btn').innerText = t.lang;
    }

    function toggleLang() {
        currentLang = currentLang === 'ru' ? 'en' : 'ru';
        applyLang();
    }

    async function loadMaterials() {
        const res = await fetch('/materials/');
        const mats = await res.json();
        const select = document.getElementById('material_id');
        select.innerHTML = mats.map(m => `<option value="${m.id}">${m.name} (${m.code})</option>`).join('');
    }

    async function calculate() {
        const D = parseFloat(document.getElementById('D').value);
        const L = parseFloat(document.getElementById('L').value);
        const S = parseFloat(document.getElementById('S').value) / 100;
        const C = parseFloat(document.getElementById('C').value);

        const res = await fetch(`/kanban/calc/?D=${D}&L=${L}&S=${S}&C=${C}`, { method: 'POST' });
        const data = await res.json();
        
        currentCalcResult = { D, L, S, C };
        
        document.getElementById('result').style.display = 'block';
        document.getElementById('val-reorder').innerText = data.levels.reorder_point;
        document.getElementById('val-max').innerText = data.levels.max_stock;
        document.getElementById('val-cards').innerText = data.cards.N_rounded_up;
    }

    async function generateCards() {
        if (!currentCalcResult) return;
        const material_id = parseInt(document.getElementById('material_id').value);
        const site_id = 1; // По умолчанию первый участок

        const payload = {
            material_id: material_id,
            site_id: site_id,
            ...currentCalcResult
        };

        await fetch('/kanban/generate-cards/', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        });
        
        alert(currentLang === 'ru' ? 'Карточки успешно созданы!' : 'Cards generated successfully!');
        loadCards();
    }

    async function loadCards() {
        const res = await fetch('/kanban/cards/');
        const cards = await res.json();
        const container = document.getElementById('cards-container');
        const t = translations[currentLang];

        if (cards.length === 0) {
            container.innerHTML = `<p style="text-align:center; color:#7f8c8d;">${t.noCards}</p>`;
            return;
        }

        container.innerHTML = cards.map(c => `
            <div class="card-item">
                <div class="card-info">
                    <strong>${c.card_number}</strong><br>
                    <span style="color: #7f8c8d;">${c.container_capacity} ед.</span>
                </div>
                <div class="card-actions">
                    <a href="/kanban/card/${c.id}/pdf/">${t.download}</a>
                </div>
            </div>
        `).join('');
    }

    // Инициализация при загрузке
    window.onload = () => {
        loadMaterials();
        loadCards();
    };
</script>
</body>
</html>
    """
