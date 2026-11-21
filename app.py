# ============================================================
#  KANBAN PRODUCTION CALCULATOR (UI + DELETE + BUG FIXES)
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

class KanbanCard(Base):
    __tablename__ = "kanban_cards"
    id = Column(Integer, primary_key=True)
    card_number = Column(String, unique=True, nullable=False)
    material_id = Column(Integer, ForeignKey("materials.id"), nullable=False)
    site_id = Column(Integer, ForeignKey("sites.id"), nullable=False)
    container_capacity = Column(Float, nullable=False)
    reorder_point = Column(Float, nullable=False)
    supplier = Column(String, nullable=True)
    is_active = Column(Boolean, default=True)
    created_at = Column(DateTime, default=_now)

Base.metadata.create_all(bind=engine)

# ============================================================
#  СХЕМЫ ДАННЫХ
# ============================================================
class MaterialNameUpdate(BaseModel):
    name: str

class GenerateCardsIn(BaseModel):
    material_id: int
    site_id: int
    D: float = Field(..., gt=0)
    L: float = Field(..., gt=0)
    S: float = Field(..., ge=0, le=1)
    C: float = Field(..., gt=0)
    supplier: Optional[str] = None

# ============================================================
#  ПРИЛОЖЕНИЕ
# ============================================================
app = FastAPI(title="Kanban Production Calculator", version="3.4")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()

def _register_card_font() -> str:
    candidates = ["/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", "C:/Windows/Fonts/Arial.ttf"]
    for path in candidates:
        if os.path.exists(path):
            try:
                pdfmetrics.registerFont(TTFont("CardFont", path))
                return "CardFont"
            except Exception:
                continue
    return "Helvetica"
_CARD_FONT = _register_card_font()

def compute_kanban_levels(D: float, L: float, S: float, C: float) -> dict:
    if C <= 0 or D <= 0 or L <= 0: raise HTTPException(400, "D, L, C должны быть > 0")
    demand_lt = D * L
    safety_stock = demand_lt * S
    reorder_point = demand_lt * (1 + S)
    n_raw = reorder_point / C
    N = ceil(n_raw)
    max_stock = N * C
    return {"levels": {"reorder_point": round(reorder_point, 4), "max_stock": round(max_stock, 4)}, "cards": {"N_rounded_up": N}}

def _card_data(db: Session, card: KanbanCard) -> dict:
    mat = db.query(Material).get(card.material_id)
    site = db.query(Site).get(card.site_id)
    return {
        "card_id": card.id, "card_number": card.card_number, "title": mat.name, "code": mat.code,
        "quantity": card.container_capacity, "unit": mat.unit, "reorder_point": card.reorder_point,
        "supplier": card.supplier or "—", "site_name": site.name, "created_at": card.created_at
    }

def _build_card_pdf(data: dict) -> bytes:
    w_pt, h_pt = 148 * mm, 105 * mm
    buf = io.BytesIO(); c = rl_canvas.Canvas(buf, pagesize=(w_pt, h_pt))
    dark, accent = HexColor("#2c3e50"), HexColor("#c0392b")
    c.setStrokeColor(dark); c.setLineWidth(2); c.rect(6, 6, w_pt - 12, h_pt - 12)
    c.setFillColor(dark); c.rect(6, h_pt - 32, w_pt - 12, 26, fill=1, stroke=0)
    c.setFillColor(HexColor("#ffffff")); c.setFont(_CARD_FONT, 10); c.drawString(12, h_pt - 18, "КАНБАН-КАРТОЧКА")
    c.drawRightString(w_pt - 12, h_pt - 18, data["card_number"])
    c.setFillColor(dark); c.setFont(_CARD_FONT, 14); c.drawString(12, h_pt - 52, data["title"][:40])
    c.setFillColor(accent); c.setFont(_CARD_FONT, 24); c.drawRightString(w_pt - 12, h_pt - 56, f"{data['quantity']:g}")
    c.setFillColor(dark); c.setFont(_CARD_FONT, 10); c.drawString(12, h_pt - 80, f"Точка заказа: {data['reorder_point']:g} {data['unit']}")
    c.drawString(12, h_pt - 95, f"Поставщик: {data['supplier']}")
    try: code128.Code128(data["card_number"], barWidth=0.9, barHeight=28).drawOn(c, 12, 16)
    except Exception: pass
    c.showPage(); c.save(); buf.seek(0)
    return buf.getvalue()

# ============================================================
#  ЭНДПОИНТЫ
# ============================================================
@app.get("/materials/")
def list_materials(db: Session = Depends(get_db)): return db.query(Material).all()

@app.patch("/materials/{material_id}/name/")
def update_material_name(material_id: int, p: MaterialNameUpdate, db: Session = Depends(get_db)):
    mat = db.query(Material).get(material_id)
    if not mat: raise HTTPException(404, "Материал не найден")
    mat.name = p.name
    db.commit(); return {"status": "ok", "new_name": mat.name}

@app.get("/kanban/cards/")
def list_cards(db: Session = Depends(get_db)): return db.query(KanbanCard).filter(KanbanCard.is_active == True).all()

# --- НОВЫЙ ЭНДПОИНТ ДЛЯ УДАЛЕНИЯ ---
@app.delete("/kanban/card/{card_id}/")
def delete_card(card_id: int, db: Session = Depends(get_db)):
    card = db.query(KanbanCard).get(card_id)
    if not card: raise HTTPException(404, "Карточка не найдена")
    card.is_active = False # Мягкое удаление
    db.commit()
    return {"status": "deleted"}

@app.post("/kanban/calc/")
def kanban_calc(D: float, L: float, S: float, C: float):
    return compute_kanban_levels(D, L, S, C)

@app.post("/kanban/generate-cards/")
def generate_cards(p: GenerateCardsIn, db: Session = Depends(get_db)):
    mat = db.query(Material).get(p.material_id)
    if not mat: raise HTTPException(404, "Материал не найден")
    
    calc = compute_kanban_levels(p.D, p.L, p.S, p.C)
    N = calc["cards"]["N_rounded_up"]
    ROP = calc["levels"]["reorder_point"]
    
    # При генерации новых карточек, старые автоматически скрываются
    for c in db.query(KanbanCard).filter_by(material_id=p.material_id, site_id=p.site_id).all():
        c.is_active = False
        
    for i in range(1, N + 1):
        db.add(KanbanCard(card_number=f"KBN-{mat.code}-{i:03d}", material_id=p.material_id, site_id=p.site_id, container_capacity=p.C, reorder_point=ROP, supplier=p.supplier))
    db.commit()
    return {"cards_created": N, "calc": calc}

@app.get("/kanban/card/{card_id}/pdf/")
def download_card_pdf(card_id: int, db: Session = Depends(get_db)):
    card = db.query(KanbanCard).get(card_id)
    if not card: raise HTTPException(404)
    return Response(content=_build_card_pdf(_card_data(db, card)), media_type="application/pdf", headers={"Content-Disposition": f'attachment; filename="{card.card_number}.pdf"'})

@app.post("/seed-demo/")
def seed_demo(db: Session = Depends(get_db)):
    if db.query(Site).count() > 0: return {"status": "already seeded"}
    
    wh = Site(name="Главный склад", code="WH", site_type="warehouse")
    db.add(wh); db.flush()
    
    m_sugar = Material(code="SUGAR", name="Сахар", unit="кг", kind="main")
    m_film = Material(code="FILM", name="Плёнка", unit="м", kind="aux")
    db.add_all([m_sugar, m_film]); db.flush()
    
    db.commit()
    return {"status": "seeded"}

# ============================================================
#  HTML ИНТЕРФЕЙС
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
        
        .actions-container { display: flex; flex-direction: column; gap: 8px; }
        .btn-download { display: inline-block; padding: 8px 12px; background: #27ae60; color: white; text-decoration: none; border-radius: 4px; font-size: 12px; text-align: center; border: none; cursor: pointer; }
        .btn-delete { display: inline-block; padding: 8px 12px; background: #e74c3c; color: white; text-decoration: none; border-radius: 4px; font-size: 12px; text-align: center; border: none; cursor: pointer; }
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
            <label id="lbl-mat">Выберите материал</label>
            <select id="material_id" onchange="updateNameField()"></select>
        </div>
        <div class="form-group">
            <label id="lbl-mat-name">Название материала (можно изменить)</label>
            <input type="text" id="material_name_input" placeholder="Введите название...">
        </div>
        <div class="form-group">
            <label id="lbl-d">Дневная потребность (шт/кг)</label>
            <input type="number" id="D" value="400" step="any">
        </div>
        <div class="form-group">
            <label id="lbl-l">Время пополнения (дней)</label>
            <input type="number" id="L" value="1.5" step="any">
        </div>
        <div class="form-group">
            <label id="lbl-s">Страховой запас (%)</label>
            <input type="number" id="S" value="20" step="any">
        </div>
        <div class="form-group">
            <label id="lbl-c">Вместимость контейнера</label>
            <input type="number" id="C" value="50" step="any">
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
            title: "Канбан Система", calcTitle: "Расчет карточек", lblMat: "Выберите материал",
            lblMatName: "Название материала (можно изменить)",
            lblD: "Дневная потребность (шт/кг)", lblL: "Время пополнения (дней)",
            lblS: "Страховой запас (%)", lblC: "Вместимость контейнера",
            btnCalc: "Рассчитать", btnGen: "Создать карточки",
            resReorder: "Точка заказа:", resMax: "Макс. запас:", resCards: "Кол-во карточек:",
            listTitle: "Активные карточки", download: "PDF", delete: "Удалить", noCards: "Нет активных карточек",
            lang: "EN"
        },
        en: {
            title: "Kanban System", calcTitle: "Card Calculation", lblMat: "Select Material",
            lblMatName: "Material Name (editable)",
            lblD: "Daily Demand (pcs/kg)", lblL: "Lead Time (days)",
            lblS: "Safety Stock (%)", lblC: "Container Capacity",
            btnCalc: "Calculate", btnGen: "Generate Cards",
            resReorder: "Reorder Point:", resMax: "Max Stock:", resCards: "Cards Qty:",
            listTitle: "Active Cards", download: "PDF", delete: "Delete", noCards: "No active cards",
            lang: "RU"
        }
    };

    let currentLang = 'ru';
    let currentCalcResult = null;
    let materialsData = [];

    function applyLang() {
        const t = translations[currentLang];
        document.getElementById('app-title').innerText = t.title;
        document.getElementById('calc-title').innerText = t.calcTitle;
        document.getElementById('lbl-mat').innerText = t.lblMat;
        document.getElementById('lbl-mat-name').innerText = t.lblMatName;
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
        loadCards();
    }

    function toggleLang() {
        currentLang = currentLang === 'ru' ? 'en' : 'ru';
        applyLang();
    }

    async function loadMaterials() {
        try {
            let res = await fetch('/materials/');
            let mats = await res.json();
            
            if (mats.length === 0) {
                document.getElementById('material_id').innerHTML = '<option>Загрузка...</option>';
                await fetch('/seed-demo/', { method: 'POST' });
                res = await fetch('/materials/');
                mats = await res.json();
            }

            materialsData = mats;
            const select = document.getElementById('material_id');
            select.innerHTML = mats.map(m => `<option value="${m.id}">${m.name} (${m.code})</option>`).join('');
            updateNameField();
        } catch (e) {
            console.error(e);
        }
    }

    function updateNameField() {
        const select = document.getElementById('material_id');
        const selectedId = parseInt(select.value);
        const mat = materialsData.find(m => m.id === selectedId);
        if (mat) {
            document.getElementById('material_name_input').value = mat.name;
        }
    }

    async function calculate() {
        const D = parseFloat(document.getElementById('D').value.replace(',', '.'));
        const L = parseFloat(document.getElementById('L').value.replace(',', '.'));
        const S = parseFloat(document.getElementById('S').value.replace(',', '.')) / 100;
        const C = parseFloat(document.getElementById('C').value.replace(',', '.'));

        if (isNaN(D) || isNaN(L) || isNaN(S) || isNaN(C)) {
            alert(currentLang === 'ru' ? "Пожалуйста, используйте точку или запятую для дробей." : "Please use dots for decimals.");
            return;
        }

        try {
            const res = await fetch(`/kanban/calc/?D=${D}&L=${L}&S=${S}&C=${C}`, { method: 'POST' });
            if (!res.ok) throw new Error('Ошибка расчета на сервере');
            
            const data = await res.json();
            currentCalcResult = { D, L, S, C };
            
            document.getElementById('result').style.display = 'block';
            document.getElementById('val-reorder').innerText = data.levels.reorder_point;
            document.getElementById('val-max').innerText = data.levels.max_stock;
            document.getElementById('val-cards').innerText = data.cards.N_rounded_up;
        } catch (e) {
            alert("Ошибка: " + e.message);
        }
    }

    async function generateCards() {
        if (!currentCalcResult) return;
        
        const material_id = parseInt(document.getElementById('material_id').value);
        if (isNaN(material_id)) {
            alert("Ошибка: Материал не выбран.");
            return;
        }

        const newName = document.getElementById('material_name_input').value;

        try {
            if (newName) {
                await fetch(`/materials/${material_id}/name/`, {
                    method: 'PATCH',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ name: newName })
                });
            }

            const payload = {
                material_id: material_id,
                site_id: 1,
                ...currentCalcResult
            };

            const res = await fetch('/kanban/generate-cards/', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            });
            
            if (res.ok) {
                alert(currentLang === 'ru' ? 'Карточки успешно созданы!' : 'Cards generated successfully!');
                await loadCards();
                await loadMaterials();
            } else {
                const err = await res.json();
                alert("Ошибка сервера: " + JSON.stringify(err));
            }
        } catch (e) {
            alert("Ошибка сети: " + e.message);
        }
    }

    async function deleteCard(cardId) {
        const msg = currentLang === 'ru' ? 'Вы уверены, что хотите удалить эту карточку? (Она будет скрыта из активных)' : 'Are you sure you want to delete this card?';
        if (!confirm(msg)) return;

        try {
            const res = await fetch(`/kanban/card/${cardId}/`, { method: 'DELETE' });
            if (res.ok) {
                await loadCards(); // Обновляем список
            } else {
                alert("Ошибка при удалении");
            }
        } catch (e) {
            alert("Ошибка сети: " + e.message);
        }
    }

    async function loadCards() {
        try {
            const res = await fetch('/kanban/cards/');
            if (!res.ok) throw new Error('Не удалось загрузить карточки');
            
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
                    <div class="actions-container">
                        <a href="/kanban/card/${c.id}/pdf/" class="btn-download">${t.download}</a>
                        <button onclick="deleteCard(${c.id})" class="btn-delete">${t.delete}</button>
                    </div>
                </div>
            `).join('');
        } catch (e) {
            console.error(e);
        }
    }

    window.onload = () => {
        loadMaterials();
        loadCards();
    };
</script>
</body>
</html>
    """
