import io
import os
import shutil
import enum
import secrets
from datetime import date, datetime
from typing import List, Optional
from fastapi import FastAPI, Depends, Request, HTTPException, status, Cookie, Form, UploadFile, File
from fastapi.responses import StreamingResponse, RedirectResponse
from fastapi.templating import Jinja2Templates
from fastapi.staticfiles import StaticFiles
from sqlalchemy import create_engine, Column, Integer, String, Date, Numeric, Enum as SQLEnum, DateTime, desc, Boolean, text
from sqlalchemy.orm import declarative_base, sessionmaker, Session
from pydantic import BaseModel
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet

# Crear carpeta para adjuntos si no existe
os.makedirs("adjuntos", exist_ok=True)

# Configuración de Base de Datos
DATABASE_URL = "sqlite:///./enlaces.db"
engine = create_engine(DATABASE_URL, connect_args={"check_same_thread": False})
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()

# --- MODELOS DE BASE DE DATOS ---

class Usuario(Base):
    __tablename__ = "usuarios"
    id = Column(Integer, primary_key=True, index=True)
    username = Column(String(50), unique=True, index=True, nullable=False)
    password = Column(String(100), nullable=False)

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
    ubicacion = Column(String(150), nullable=True) 
    observaciones = Column(String(1000), nullable=True)
    archivo_adjunto = Column(String(255), nullable=True)
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
    cupo_transferencia = Column(String(100), nullable=True)
    coordenadas_gps = Column(String(100), nullable=True)
    sn_antena = Column(String(100), nullable=True)
    sn_modem = Column(String(100), nullable=True)
    nro_te = Column(String(50), nullable=True)
    nro_item = Column(String(50), nullable=True)
    eliminado = Column(Boolean, default=False)

class LogActividad(Base):
    __tablename__ = "logs_actividad"
    id = Column(Integer, primary_key=True, index=True)
    fecha_hora = Column(DateTime, default=datetime.now)
    usuario = Column(String(50), nullable=False)
    accion = Column(String(50), nullable=False)
    entidad_id = Column(Integer, nullable=False)
    detalle = Column(String(500), nullable=False)

Base.metadata.create_all(bind=engine)

# Inicializar Base de datos de Usuarios si está vacía
db_init = SessionLocal()
if db_init.query(Usuario).count() == 0:
    db_init.add(Usuario(username="scortes", password="acatenestusistemita"))
    db_init.add(Usuario(username="afellenz", password="Va2005fe"))
    db_init.commit()
db_init.close()

# Parche seguro para columna adjuntos (por si acaso)
with engine.connect() as conn:
    try:
        conn.execute(text("ALTER TABLE enlaces_telecom ADD COLUMN archivo_adjunto VARCHAR(255)"))
        conn.commit()
    except Exception:
        pass 

# --- SCHEMAS Pydantic ---
class UsuarioCreate(BaseModel):
    username: str
    password: str

class EnlaceCreate(BaseModel):
    referencia: str
    organismo: str
    ubicacion: Optional[str] = None 
    observaciones: Optional[str] = None
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
    cupo_transferencia: Optional[str] = None
    coordenadas_gps: Optional[str] = None
    sn_antena: Optional[str] = None
    sn_modem: Optional[str] = None
    nro_te: Optional[str] = None
    nro_item: Optional[str] = None


app = FastAPI()
app.mount("/adjuntos", StaticFiles(directory="adjuntos"), name="adjuntos")
templates = Jinja2Templates(directory="templates")

def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()

def verificar_usuario(session_user: str = Cookie(None), db: Session = Depends(get_db)):
    if not session_user:
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="Inicie sesión")
    user = db.query(Usuario).filter(Usuario.username == session_user).first()
    if not user:
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="Inicie sesión")
    return session_user

# --- RUTAS DE NAVEGACIÓN Y AUTENTICACIÓN ---

@app.get("/")
def pagina_bienvenida(request: Request, session_user: str = Cookie(None), db: Session = Depends(get_db)):
    if session_user:
        user = db.query(Usuario).filter(Usuario.username == session_user).first()
        if user:
            return RedirectResponse(url="/sistema", status_code=status.HTTP_303_SEE_OTHER)
    return templates.TemplateResponse(request=request, name="login.html")

@app.post("/login")
def procesar_login(username: str = Form(...), password: str = Form(...), db: Session = Depends(get_db)):
    user = db.query(Usuario).filter(Usuario.username == username).first()
    if user and secrets.compare_digest(password, user.password):
        response = RedirectResponse(url="/sistema", status_code=status.HTTP_303_SEE_OTHER)
        response.set_cookie(key="session_user", value=username, httponly=True)
        return response
    return RedirectResponse(url="/?error=true", status_code=status.HTTP_303_SEE_OTHER)

@app.get("/sistema")
def leer_interfaz(request: Request, session_user: str = Cookie(None), db: Session = Depends(get_db)):
    if not session_user:
        return RedirectResponse(url="/", status_code=status.HTTP_303_SEE_OTHER)
    user = db.query(Usuario).filter(Usuario.username == session_user).first()
    if not user:
        return RedirectResponse(url="/", status_code=status.HTTP_303_SEE_OTHER)
    return templates.TemplateResponse(request=request, name="index.html", context={"session_user": session_user})

@app.get("/logout")
def logout():
    response = RedirectResponse(url="/", status_code=status.HTTP_303_SEE_OTHER)
    response.delete_cookie("session_user")
    return response

# --- ENDPOINTS DE GESTIÓN DE USUARIOS (Solo afellenz) ---
@app.get("/usuarios/")
def listar_usuarios(db: Session = Depends(get_db), user: str = Depends(verificar_usuario)):
    if user != "afellenz":
        raise HTTPException(status_code=403, detail="No autorizado")
    return [{"id": u.id, "username": u.username} for u in db.query(Usuario).all()]

@app.post("/usuarios/")
def crear_usuario(u: UsuarioCreate, db: Session = Depends(get_db), user: str = Depends(verificar_usuario)):
    if user != "afellenz":
        raise HTTPException(status_code=403, detail="No autorizado")
    db.add(Usuario(username=u.username, password=u.password))
    db.commit()
    return {"ok": True}

@app.put("/usuarios/{id}")
def editar_usuario(id: int, u: UsuarioCreate, db: Session = Depends(get_db), user: str = Depends(verificar_usuario)):
    if user != "afellenz":
        raise HTTPException(status_code=403, detail="No autorizado")
    db_u = db.query(Usuario).filter(Usuario.id == id).first()
    if db_u:
        db_u.username = u.username
        db_u.password = u.password
        db.commit()
    return {"ok": True}

@app.delete("/usuarios/{id}")
def eliminar_usuario(id: int, db: Session = Depends(get_db), user: str = Depends(verificar_usuario)):
    if user != "afellenz":
        raise HTTPException(status_code=403, detail="No autorizado")
    db_u = db.query(Usuario).filter(Usuario.id == id).first()
    if db_u:
        db.delete(db_u)
        db.commit()
    return {"ok": True}

# --- ENDPOINTS DE API REST (ENLACES) ---

@app.post("/enlaces/")
def crear_enlace(enlace: EnlaceCreate, db: Session = Depends(get_db), user: str = Depends(verificar_usuario)):
    nuevo_enlace = Enlace(**enlace.model_dump())
    db.add(nuevo_enlace)
    db.commit()
    db.refresh(nuevo_enlace)
    db.add(LogActividad(usuario=user, accion="CREAR", entidad_id=nuevo_enlace.id, detalle=f"Creación: {nuevo_enlace.referencia}"))
    db.commit()
    return {"ok": True, "id": nuevo_enlace.id}

@app.get("/enlaces/{id}")
def obtener_enlace(id: int, db: Session = Depends(get_db), user: str = Depends(verificar_usuario)):
    return db.query(Enlace).filter(Enlace.id == id).first()

@app.put("/enlaces/{id}")
def actualizar_enlace(id: int, e: EnlaceCreate, db: Session = Depends(get_db), user: str = Depends(verificar_usuario)):
    db_e = db.query(Enlace).filter(Enlace.id == id).first()
    cambios = []
    for key, value in e.model_dump().items():
        if str(getattr(db_e, key)) != str(value):
            cambios.append(f"{key}: {getattr(db_e, key)} -> {value}")
        setattr(db_e, key, value)
    if cambios:
        db.add(LogActividad(usuario=user, accion="EDITAR", entidad_id=id, detalle=" | ".join(cambios)[:500]))
    db.commit()
    return {"ok": True, "id": id}

@app.post("/enlaces/{id}/adjunto")
def subir_adjunto(id: int, file: UploadFile = File(...), db: Session = Depends(get_db), user: str = Depends(verificar_usuario)):
    enlace = db.query(Enlace).filter(Enlace.id == id).first()
    if not enlace:
        raise HTTPException(status_code=404)
    
    ext = file.filename.split(".")[-1]
    filename = f"alta_{id}_{int(datetime.now().timestamp())}.{ext}"
    filepath = os.path.join("adjuntos", filename)
    
    with open(filepath, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)
    
    enlace.archivo_adjunto = filename
    db.add(LogActividad(usuario=user, accion="ADJUNTAR", entidad_id=id, detalle=f"Archivo de alta subido: {filename}"))
    db.commit()
    return {"ok": True, "filename": filename}

@app.get("/enlaces/")
def listar_enlaces(db: Session = Depends(get_db), user: str = Depends(verificar_usuario)):
    return db.query(Enlace).filter(Enlace.eliminado == False).all()

@app.patch("/enlaces/{enlace_id}/estado")
def cambiar_estado(enlace_id: int, estado: EstadoEnlace, db: Session = Depends(get_db), user: str = Depends(verificar_usuario)):
    enlace = db.query(Enlace).filter(Enlace.id == enlace_id).first()
    if enlace.estado != estado:
        db.add(LogActividad(usuario=user, accion="CAMBIAR_ESTADO", entidad_id=enlace.id, detalle=f"Estado: {enlace.estado} -> {estado.value}"))
        enlace.estado = estado
        db.commit()
    return {"ok": True}

@app.patch("/enlaces/{id}/eliminar")
def eliminar_enlace(id: int, db: Session = Depends(get_db), user: str = Depends(verificar_usuario)):
    enlace = db.query(Enlace).filter(Enlace.id == id).first()
    enlace.eliminado = True
    db.add(LogActividad(usuario=user, accion="ELIMINAR", entidad_id=enlace.id, detalle=f"Eliminación lógica: {enlace.referencia}"))
    db.commit()
    return {"ok": True}

@app.get("/logs/")
def listar_logs(db: Session = Depends(get_db), user: str = Depends(verificar_usuario)):
    return db.query(LogActividad).order_by(desc(LogActividad.fecha_hora)).limit(100).all()

# --- REPORTES (Filtrados por eliminado == False) ---
@app.get("/reportes/excel")
def exportar_excel(db: Session = Depends(get_db), user: str = Depends(verificar_usuario)):
    data = [e.__dict__ for e in db.query(Enlace).filter(Enlace.eliminado == False).all()]
    df = pd.DataFrame(data)
    wb = Workbook()
    ws = wb.active
    header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    headers = ['ITEM', 'Referencia', 'Organismo', 'Ubicación', 'Localidad', 'Tipo', 'Ancho Banda', 'Vta s/IVA', 'Vta c/IVA', 'Instalacion', 'Mantenimiento', 'Alta', 'Cupo', 'GPS', 'S/N Antena', 'S/N Modem', 'Nro TE', 'Nro ITEM']
    ws.append(headers)
    for cell in ws[1]: cell.fill = header_fill; cell.font = header_font; cell.border = border
    
    for idx, e in enumerate(data, start=1):
        row = [
            idx, e['referencia'], e['organismo'], e.get('ubicacion', ''), e['localidad'], e['tipo_enlace'].value, 
            e['ancho_banda'], 
            f"USD {float(e['precio_venta_sin_iva']):.2f}", 
            f"USD {float(e['precio_venta_con_iva']):.2f}", 
            f"USD {float(e['costo_instalacion']):.2f}", 
            f"$ {float(e['costo_mantenimiento']):.2f}", 
            str(e['fecha_alta']), e.get('cupo_transferencia', ''), e.get('coordenadas_gps', ''), 
            e.get('sn_antena', ''), e.get('sn_modem', ''), e.get('nro_te', ''), e.get('nro_item', '')
        ]
        ws.append(row)
        for cell in ws[ws.max_row]: cell.border = border
        
    if not df.empty:
        totales = ['TOTALES', '', '', '', '', '', '', f"USD {df['precio_venta_sin_iva'].sum():.2f}", f"USD {df['precio_venta_con_iva'].sum():.2f}", f"USD {df['costo_instalacion'].sum():.2f}", f"$ {df['costo_mantenimiento'].sum():.2f}", '', '', '', '', '', '', '']
        ws.append(totales)
        
    stream = io.BytesIO(); wb.save(stream); stream.seek(0)
    return StreamingResponse(stream, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-Disposition": "attachment; filename=Reporte.xlsx"})

@app.get("/reportes/pdf")
def exportar_pdf(db: Session = Depends(get_db), user: str = Depends(verificar_usuario)):
    data = [e.__dict__ for e in db.query(Enlace).filter(Enlace.eliminado == False).all()]
    df = pd.DataFrame(data)
    pdf_buffer = io.BytesIO()
    
    doc = SimpleDocTemplate(pdf_buffer, pagesize=landscape(A4), leftMargin=15, rightMargin=15, topMargin=30, bottomMargin=30)
    
    table_data = [['ITEM', 'Ref', 'Organismo', 'Ubicación', 'Localidad', 'Tipo', 'Ancho B.', 'Vta s/IVA', 'Vta c/IVA', 'Instal.', 'Mant.', 'Alta']]
    
    for idx, e in enumerate(data, start=1):
        table_data.append([
            idx, e['referencia'], e['organismo'], e.get('ubicacion', ''), e['localidad'], e['tipo_enlace'].value, e['ancho_banda'], 
            f"USD {float(e['precio_venta_sin_iva']):.2f}", 
            f"USD {float(e['precio_venta_con_iva']):.2f}", 
            f"USD {float(e['costo_instalacion']):.2f}", 
            f"$ {float(e['costo_mantenimiento']):.2f}", 
            str(e['fecha_alta'])
        ])
        
    if not df.empty:
        table_data.append(['TOTALES', '', '', '', '', '', '', f"USD {df['precio_venta_sin_iva'].sum():.2f}", f"USD {df['precio_venta_con_iva'].sum():.2f}", f"USD {df['costo_instalacion'].sum():.2f}", f"$ {df['costo_mantenimiento'].sum():.2f}", ''])
        
    t = Table(table_data)
    
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1F497D')), 
        ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke), 
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('FONTSIZE', (0,0), (-1,-1), 7), 
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('ALIGN', (0,0), (-1,-1), 'CENTER')
    ]))
    
    doc.build([Paragraph("Reporte Institucional de Enlaces", getSampleStyleSheet()['Heading2']), Spacer(1, 10), t])
    pdf_buffer.seek(0)
    return StreamingResponse(pdf_buffer, media_type="application/pdf", headers={"Content-Disposition": "attachment; filename=Reporte.pdf"})
