import os
import io
import enum
import secrets
from datetime import date
from typing import List
from fastapi import FastAPI, Depends, Request, HTTPException, status, Cookie, Form
from fastapi.responses import StreamingResponse, RedirectResponse
from fastapi.templating import Jinja2Templates
from sqlalchemy import create_engine, Column, Integer, String, Date, Numeric, Enum as SQLEnum
from sqlalchemy.orm import declarative_base, sessionmaker, Session
from pydantic import BaseModel
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet

# Configuración de Seguridad y Base de Usuarios
USER_DB = {
    "scortes": "acatenestusistemita",
    "afellenz": "Va2005fe"
}

def verificar_usuario(session_user: str = Cookie(None)):
    if not session_user or session_user not in USER_DB:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Inicie sesión para acceder a este recurso",
        )
    return session_user

# --- NUEVA Configuración de Base de Datos ---
# Busca la base de datos en la nube (Render), si no la encuentra, usa SQLite en tu PC
DATABASE_URL = os.getenv("DATABASE_URL", "sqlite:///./enlaces.db")

# SQLAlchemy requiere que las URLs de Postgres empiecen con 'postgresql://'
if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://", 1)

# Solo usamos check_same_thread si estamos en SQLite (local)
if "sqlite" in DATABASE_URL:
    engine = create_engine(DATABASE_URL, connect_args={"check_same_thread": False})
else:
    engine = create_engine(DATABASE_URL)

SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()
# ---------------------------------------------

class EstadoEnlace(str, enum.Enum):
    ACTIVO = "ACTIVO"
    INACTIVO = "INACTIVO"

class TipoEnlace(str, enum.Enum):
    FIBRA_DEDICADO = "Fibra Optica dedicado"
    FIBRA_ASIMETRICO = "Fibra Optica asimetrico"
    RF = "Radio Frecuencia"
    LEO = "LEO"
    LEO_ITINERANTE = "LEO Itinerante"
    GEO = "GEO"
    L2L = "L2L (transporte de datos)"

class Enlace(Base):
    __tablename__ = "enlaces_telecom"
    id = Column(Integer, primary_key=True, index=True)
    referencia = Column(String(50), unique=True, index=True, nullable=False)
    organismo = Column(String(150), nullable=False)
    localidad = Column(String(100), nullable=False)
    tipo_enlace = Column(SQLEnum(TipoEnlace), nullable=False)
    estado = Column(SQLEnum(EstadoEnlace), default=EstadoEnlace.ACTIVO)
    ancho_banda = Column(String(50), nullable=False)
    moneda = Column(String(10), nullable=False)
    precio_costo = Column(Numeric(12, 2), default=0.00)
    costo_mantenimiento = Column(Numeric(12, 2), default=0.00)
    precio_venta_sin_iva = Column(Numeric(12, 2), default=0.00)
    precio_venta_con_iva = Column(Numeric(12, 2), default=0.00)
    costo_instalacion = Column(Numeric(12, 2), default=0.00)
    fecha_alta = Column(Date, nullable=False)
    proveedor = Column(String(100), nullable=False)

Base.metadata.create_all(bind=engine)

class EnlaceCreate(BaseModel):
    referencia: str
    organismo: str
    localidad: str
    tipo_enlace: TipoEnlace
    ancho_banda: str
    moneda: str
    precio_costo: float
    costo_mantenimiento: float
    precio_venta_sin_iva: float
    precio_venta_con_iva: float
    costo_instalacion: float
    fecha_alta: date
    proveedor: str

app = FastAPI()
templates = Jinja2Templates(directory="templates")

def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()

# --- RUTAS DE NAVEGACIÓN Y AUTENTICACIÓN ---

@app.get("/")
def pagina_bienvenida(request: Request, session_user: str = Cookie(None)):
    if session_user and session_user in USER_DB:
        return RedirectResponse(url="/sistema", status_code=status.HTTP_303_SEE_OTHER)
    return templates.TemplateResponse(request=request, name="login.html")

@app.post("/login")
def procesar_login(username: str = Form(...), password: str = Form(...)):
    password_correcta = secrets.compare_digest(password, USER_DB.get(username, ""))
    if username in USER_DB and password_correcta:
        response = RedirectResponse(url="/sistema", status_code=status.HTTP_303_SEE_OTHER)
        response.set_cookie(key="session_user", value=username, httponly=True)
        return response
    return RedirectResponse(url="/?error=true", status_code=status.HTTP_303_SEE_OTHER)

@app.get("/sistema")
def leer_interfaz(request: Request, session_user: str = Cookie(None)):
    if not session_user or session_user not in USER_DB:
        return RedirectResponse(url="/", status_code=status.HTTP_303_SEE_OTHER)
    return templates.TemplateResponse(request=request, name="index.html")

@app.get("/logout")
def logout():
    response = RedirectResponse(url="/", status_code=status.HTTP_303_SEE_OTHER)
    response.delete_cookie("session_user")
    return response

# --- ENDPOINTS DE API REST ---

@app.post("/enlaces/")
def crear_enlace(enlace: EnlaceCreate, db: Session = Depends(get_db), user: str = Depends(verificar_usuario)):
    nuevo_enlace = Enlace(**enlace.model_dump())
    db.add(nuevo_enlace)
    db.commit()
    return {"ok": True}

@app.get("/enlaces/")
def listar_enlaces(db: Session = Depends(get_db), user: str = Depends(verificar_usuario)):
    return db.query(Enlace).all()

@app.get("/enlaces/{enlace_id}")
def obtener_enlace(enlace_id: int, db: Session = Depends(get_db), user: str = Depends(verificar_usuario)):
    enlace = db.query(Enlace).filter(Enlace.id == enlace_id).first()
    if not enlace:
        raise HTTPException(status_code=404, detail="Enlace no encontrado")
    return enlace

@app.put("/enlaces/{enlace_id}")
def actualizar_enlace(enlace_id: int, enlace_data: EnlaceCreate, db: Session = Depends(get_db), user: str = Depends(verificar_usuario)):
    enlace = db.query(Enlace).filter(Enlace.id == enlace_id).first()
    if not enlace:
        raise HTTPException(status_code=404, detail="Enlace no encontrado")
    
    for key, value in enlace_data.model_dump().items():
        setattr(enlace, key, value)
        
    db.commit()
    return {"ok": True}

@app.patch("/enlaces/{enlace_id}/estado")
def cambiar_estado(enlace_id: int, estado: EstadoEnlace, db: Session = Depends(get_db), user: str = Depends(verificar_usuario)):
    enlace = db.query(Enlace).filter(Enlace.id == enlace_id).first()
    enlace.estado = estado
    db.commit()
    return {"ok": True}

@app.get("/reportes/excel")
def exportar_excel(db: Session = Depends(get_db), user: str = Depends(verificar_usuario)):
    data = [e.__dict__ for e in db.query(Enlace).all()]
    df = pd.DataFrame(data)
    
    wb = Workbook()
    ws = wb.active
    
    header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    border = Border(
        left=Side(style='thin'), 
        right=Side(style='thin'), 
        top=Side(style='thin'), 
        bottom=Side(style='thin')
    )
    
    headers = [
        'Referencia', 'Organismo', 'Localidad', 'Tipo', 'Ancho Banda', 
        'Moneda', 'Vta s/IVA', 'Vta c/IVA', 'Instalacion', 'Mantenimiento', 'Alta'
    ]
    ws.append(headers)
    
    for cell in ws[1]: 
        cell.fill = header_fill
        cell.font = header_font
        cell.border = border
        
    for e in data:
        row = [
            e['referencia'], 
            e['organismo'], 
            e['localidad'], 
            e['tipo_enlace'].value, 
            e['ancho_banda'], 
            e['moneda'], 
            float(e['precio_venta_sin_iva']), 
            float(e['precio_venta_con_iva']), 
            float(e['costo_instalacion']), 
            float(e['costo_mantenimiento']), 
            str(e['fecha_alta'])
        ]
        ws.append(row)
        for cell in ws[ws.max_row]: 
            cell.border = border
            
    totales = [
        'TOTALES', '', '', '', '', '', 
        df['precio_venta_sin_iva'].sum() if not df.empty else 0, 
        df['precio_venta_con_iva'].sum() if not df.empty else 0, 
        df['costo_instalacion'].sum() if not df.empty else 0, 
        df['costo_mantenimiento'].sum() if not df.empty else 0, 
        ''
    ]
    ws.append(totales)
    
    stream = io.BytesIO()
    wb.save(stream)
    stream.seek(0)
    
    return StreamingResponse(
        stream, 
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
        headers={"Content-Disposition": "attachment; filename=Reporte.xlsx"}
    )

@app.get("/reportes/pdf")
def exportar_pdf(db: Session = Depends(get_db), user: str = Depends(verificar_usuario)):
    data = [e.__dict__ for e in db.query(Enlace).all()]
    df = pd.DataFrame(data)
    
    pdf_buffer = io.BytesIO()
    doc = SimpleDocTemplate(pdf_buffer, pagesize=landscape(A4))
    
    table_data = [[
        'Ref', 'Organismo', 'Localidad', 'Tipo', 'Ancho B.', 
        'Moneda', 'Vta s/IVA', 'Vta c/IVA', 'Instal.', 'Mant.', 'Alta'
    ]]
    
    for e in data:
        ref = e['referencia']
        org = e['organismo']
        loc = e['localidad']
        tipo = e['tipo_enlace'].value
        ancho = e['ancho_banda']
        moneda = e['moneda']
        vta_s_iva = f"{float(e['precio_venta_sin_iva']):.2f}"
        vta_c_iva = f"{float(e['precio_venta_con_iva']):.2f}"
        instalacion = f"{float(e['costo_instalacion']):.2f}"
        mantenimiento = f"{float(e['costo_mantenimiento']):.2f}"
        fecha_alta = str(e['fecha_alta'])
        
        row = [ref, org, loc, tipo, ancho, moneda, vta_s_iva, vta_c_iva, instalacion, mantenimiento, fecha_alta]
        table_data.append(row)
        
    totales_row = [
        'TOTALES', '', '', '', '', '', 
        f"{df['precio_venta_sin_iva'].sum():.2f}" if not df.empty else "0.00", 
        f"{df['precio_venta_con_iva'].sum():.2f}" if not df.empty else "0.00", 
        f"{df['costo_instalacion'].sum():.2f}" if not df.empty else "0.00", 
        f"{df['costo_mantenimiento'].sum():.2f}" if not df.empty else "0.00", 
        ''
    ]
    table_data.append(totales_row)
    
    t = Table(table_data)
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1F497D')), 
        ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke), 
        ('GRID', (0,0), (-1,-1), 0.5, colors.black)
    ]))
    
    doc.build([Paragraph("Reporte Institucional de Enlaces", getSampleStyleSheet()['Heading2']), Spacer(1, 10), t])
    pdf_buffer.seek(0)
    
    return StreamingResponse(
        pdf_buffer, 
        media_type="application/pdf", 
        headers={"Content-Disposition": "attachment; filename=Reporte.pdf"}
    )
