"""
Gestió Aigues Les Lloses — v5.0
- Solo modo light (fuerza fondo blanco)
- Login con inputs visibles sobre fondo blanco
- Carga de vecinos desde CSV
- Email simple: solo dirección remitente (usa sendmail local o Gmail app password)
- Facturas guardadas en disco /data/facturas/
- Consumo medio mensual visible para admin y vecino
- Diseño para VPS (no Streamlit Cloud)
"""
import streamlit as st
import pandas as pd
import sqlite3
import openpyxl
import hashlib
import secrets
import os
import shutil
import threading
import json
from datetime import date, datetime
from contextlib import contextmanager
from io import BytesIO
from pathlib import Path
import smtplib, ssl as ssl_lib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle
import plotly.graph_objects as go

# ── MQTT LISTENER ─────────────────────────────────────────────────────────────
def start_mqtt_listener(broker, port, topic):
    """Arranca el listener MQTT en un hilo de fondo."""
    try:
        import paho.mqtt.client as mqtt
    except ImportError:
        return

    def on_message(client, userdata, msg):
        try:
            payload = msg.payload.decode("utf-8")
            topic_parts = msg.topic.split("/")
            # topic esperado: leslloses/contador/SENSOR_ID
            if len(topic_parts) >= 3:
                sensor_id = topic_parts[-1]
                # Intentar parsear JSON o valor directo
                try:
                    data = json.loads(payload)
                    pulsos = int(data.get("pulse", data.get("count", data.get("value", 0))))
                except Exception:
                    pulsos = int(float(payload))
                # Guardar en BD
                with sqlite3.connect(DB) as con:
                    con.row_factory = sqlite3.Row
                    sensor = con.execute(
                        "SELECT * FROM sensores WHERE sensor_id=? AND activo=1",
                        (sensor_id,)).fetchone()
                    if sensor:
                        pulsos_m3 = sensor["pulsos_m3"] or 100
                        m3 = round(pulsos / pulsos_m3, 3)
                        hoy = date.today()
                        vecino_id = sensor["vecino_id"]
                        # Actualizar o insertar consumo del mes actual
                        ex = con.execute(
                            "SELECT id FROM consumos WHERE vecino_id=? AND anyo=? AND mes=?",
                            (vecino_id, hoy.year, hoy.month)).fetchone()
                        if ex:
                            con.execute(
                                "UPDATE consumos SET m3=?,fuente=? WHERE id=?",
                                (m3, "mqtt", ex[0]))
                        else:
                            con.execute(
                                "INSERT INTO consumos(vecino_id,anyo,mes,m3,fuente) VALUES(?,?,?,?,?)",
                                (vecino_id, hoy.year, hoy.month, m3, "mqtt"))
                        # Actualizar última lectura del sensor
                        con.execute(
                            "UPDATE sensores SET ultima_lectura=?,ultimo_pulso=? WHERE sensor_id=?",
                            (datetime.now().strftime("%Y-%m-%d %H:%M"), pulsos, sensor_id))
        except Exception:
            pass

    def on_connect(client, userdata, flags, rc):
        if rc == 0:
            client.subscribe(topic)

    client = mqtt.Client()
    client.on_connect = on_connect
    client.on_message = on_message
    try:
        client.connect(broker, port, 60)
        client.loop_forever()
    except Exception:
        pass

def ensure_mqtt_running():
    """Arranca el listener MQTT si está configurado y no está corriendo."""
    if st.session_state.get("mqtt_thread_started"):
        return
    tar = get_tar()
    if tar.get("mqtt_activo"):
        broker = tar.get("mqtt_broker", "localhost")
        port = int(tar.get("mqtt_port", 1883))
        topic = tar.get("mqtt_topic", "leslloses/#")
        t = threading.Thread(
            target=start_mqtt_listener,
            args=(broker, port, topic),
            daemon=True)
        t.start()
        st.session_state["mqtt_thread_started"] = True


DB           = "lloses.db"
FACTURAS_DIR = Path("data/facturas")
FACTURAS_DIR.mkdir(parents=True, exist_ok=True)

st.set_page_config(
    page_title="Aigues Les Lloses",
    page_icon="💧",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── FORZAR MODO LIGHT — anula dark mode completamente ─────────────────────────
st.markdown("""
<style>
  /* Forzar light mode independientemente de preferencias del sistema */
  :root {
    color-scheme: light only !important;
  }
  html, body, .stApp, [data-testid="stAppViewContainer"],
  [data-testid="stHeader"], [data-testid="stToolbar"] {
    background-color: #f1f5f9 !important;
    color: #0f172a !important;
  }
  /* Sidebar */
  section[data-testid="stSidebar"] {
    background-color: #ffffff !important;
    border-right: 1.5px solid #e2e8f0 !important;
  }
  section[data-testid="stSidebar"] * { color: #0f172a !important; }

  /* Inputs — fondo BLANCO siempre, texto oscuro */
  input, textarea, [data-baseweb="input"] input,
  [data-baseweb="textarea"] textarea,
  [data-testid="stTextInput"] input,
  [data-testid="stNumberInput"] input {
    background-color: #ffffff !important;
    color: #0f172a !important;
    border: 1.5px solid #cbd5e1 !important;
    border-radius: 8px !important;
  }
  input::placeholder { color: #94a3b8 !important; }
  input:focus, textarea:focus {
    border-color: #1a6fc4 !important;
    box-shadow: 0 0 0 3px rgba(26,111,196,.12) !important;
  }
  /* Labels inputs */
  [data-testid="stTextInput"] label,
  [data-testid="stNumberInput"] label,
  [data-testid="stSelectbox"] label,
  [data-testid="stRadio"] label,
  p, span, div { color: #0f172a !important; }

  /* Selectbox */
  [data-baseweb="select"] > div {
    background-color: #ffffff !important;
    border: 1.5px solid #cbd5e1 !important;
    border-radius: 8px !important;
    color: #0f172a !important;
  }

  /* Ocultar elementos innecesarios */
  footer, .stDeployButton, #MainMenu { display:none!important; visibility:hidden!important; }

  /* Botones */
  .stButton > button {
    background: #1a6fc4 !important;
    color: #ffffff !important;
    border: none !important;
    border-radius: 8px !important;
    font-weight: 600 !important;
    padding: .45rem 1.2rem !important;
    transition: all .15s !important;
  }
  .stButton > button:hover { background: #3b9ede !important; }
  .stButton > button[kind="secondary"] {
    background: transparent !important;
    color: #1a6fc4 !important;
    border: 1.5px solid #1a6fc4 !important;
  }

  /* Métricas */
  div[data-testid="stMetric"] {
    background: #ffffff !important;
    border: 1px solid #e2e8f0 !important;
    border-radius: 12px !important;
    padding: .9rem 1.1rem !important;
    box-shadow: 0 1px 3px rgba(0,0,0,.06) !important;
  }
  div[data-testid="stMetric"] label { color: #64748b !important; font-size:.78rem !important; font-weight:600 !important; }
  div[data-testid="stMetric"] [data-testid="stMetricValue"] { color: #0f172a !important; }

  /* Tabs */
  .stTabs [data-baseweb="tab-list"] {
    background: #ffffff !important;
    border-radius: 10px !important;
    border: 1px solid #e2e8f0 !important;
    padding: 2px !important;
  }
  .stTabs [data-baseweb="tab"] { color: #64748b !important; border-radius:8px !important; font-weight:500 !important; }
  .stTabs [aria-selected="true"] { color: #1a6fc4 !important; background: #e8f4fd !important; }

  /* Expander */
  div[data-testid="stExpander"] {
    background: #ffffff !important;
    border: 1px solid #e2e8f0 !important;
    border-radius: 10px !important;
    margin-bottom: .5rem !important;
  }

  /* Dataframe */
  .stDataFrame { border-radius:10px !important; border:1px solid #e2e8f0 !important; }
  [data-testid="stDataFrame"] { background:#ffffff !important; }

  /* Form */
  [data-testid="stForm"] {
    background: #ffffff !important;
    border: 1px solid #e2e8f0 !important;
    border-radius: 12px !important;
    padding: 1rem !important;
  }

  /* Cards custom */
  .card {
    background: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 12px;
    padding: 1.1rem 1.3rem;
    margin-bottom: .9rem;
    box-shadow: 0 1px 3px rgba(0,0,0,.05);
  }
  .badge-admin {
    background: #1a6fc4; color:#fff; border-radius:20px;
    padding:3px 12px; font-size:.72rem; font-weight:600;
  }
  .badge-vecino {
    background: #1a9e5c; color:#fff; border-radius:20px;
    padding:3px 12px; font-size:.72rem; font-weight:600;
  }
  .login-box {
    background: #ffffff !important;
    border: 1px solid #e2e8f0 !important;
    border-radius: 16px !important;
    padding: 2rem !important;
    box-shadow: 0 4px 24px rgba(0,0,0,.08) !important;
  }
  h1 { color:#0f172a !important; font-size:1.6rem !important; font-weight:700 !important; }
  h2 { color:#0f172a !important; font-size:1.2rem !important; }
  h3 { color:#0f172a !important; font-size:1rem !important; }
  .block-container { padding-top:1.5rem !important; }
</style>
""", unsafe_allow_html=True)

# ── COLORES ───────────────────────────────────────────────────────────────────
P = {
    "blue":"#1a6fc4","lblue":"#3b9ede","sky":"#e8f4fd",
    "green":"#1a9e5c","lgreen":"#e8f9f1",
    "red":"#c0392b","lred":"#fdecea",
    "amber":"#d97706","lamber":"#fef3c7",
    "gray":"#64748b","dark":"#0f172a",
    "bg":"#f1f5f9","card":"#ffffff","border":"#e2e8f0",
}

# ── IDIOMAS ───────────────────────────────────────────────────────────────────
LANG = {
  "es":{
    "months":["Enero","Febrero","Marzo","Abril","Mayo","Junio",
              "Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"],
    "menu_consums":"📊 Consumos","menu_veins":"👥 Vecinos",
    "menu_fact":"🧾 Facturación","menu_cfg":"⚙️ Configuración",
    "menu_myfact":"🧾 Mis Facturas","menu_mycons":"📊 Mis Consumos",
    "login_title":"Iniciar sesión","login_user":"Usuario","login_pass":"Contraseña",
    "login_btn":"Entrar","login_err":"Usuario o contraseña incorrectos","logout":"Cerrar sesión",
    "year":"Año","month":"Mes","neighbor":"Vecino","total":"Total",
    "download_pdf":"⬇️ Descargar PDF","send_all":"📨 Enviar todas por email",
    "save":"Guardar","add":"Añadir","name":"Nombre","address":"Dirección",
    "email":"Email","no_email":"✉️ Sin email","list":"Listado",
    "add_neighbor":"Añadir vecino","del_neighbor":"Eliminar vecino",
    "register_reading":"Registrar lectura","tariff":"Tarifa €/m³",
    "fixed_fee":"Cuota fija €","vat":"IVA %","saved_ok":"✅ Guardado correctamente",
    "sent_ok":"facturas enviadas","no_data":"Sin datos para este período.",
    "view_per_neighbor":"Por vecino","view_accumulated":"Acumulado total",
    "bar_chart":"Comparativa mensual","invoice_period":"Facturas de",
    "total_billed":"Total facturado","concept":"Concepto",
    "quantity":"Cantidad","unit_price":"Precio unitario","amount":"Importe",
    "fixed_service":"Cuota fija de servicio","tax_base":"Base imponible",
    "total_pay":"TOTAL A PAGAR",
    "thanks":"Gracias por su colaboración en el uso responsable del agua.",
    "payment_info":"Pago: domiciliación bancaria o transferencia",
    "search_neighbor":"🔍 Buscar vecino...","role_admin":"Administrador","role_vecino":"Vecino",
    "client":"Cliente","issue_date":"Fecha emisión","invoice_number":"Nº Factura",
    "send_single":"Enviar factura individual","send_btn":"📨 Enviar",
    "users":"Usuarios","edit_user":"Editar usuario","del_user":"Eliminar usuario",
    "add_user":"Añadir usuario","new_pw":"Nueva contraseña",
    "export_list":"⬇️ Exportar Excel","phone":"Teléfono","edit_neighbor":"Editar vecino",
    "export_consumption":"⬇️ CSV consumos","export_chart":"⬇️ PNG gráfica",
    "annual_evolution":"Evolución anual","avg_monthly":"Consumo medio mensual",
    "my_consumption":"Mi consumo","cost_per_m3":"Tarifa €/m³",
    "confirm_delete":"⚠️ Confirmo que quiero eliminar este registro",
    "prev_reading":"Lectura anterior","curr_reading":"Lectura actual",
    "consumption":"Consumo calculado",
    "reading_error":"⚠️ La lectura actual no puede ser menor que la anterior.",
    "backup_db":"⬇️ Backup base de datos","iban":"IBAN",
    "sensors_tab":"📡 Sensores","sensor_id":"ID Sensor","pulsos_m3":"Pulsos/m³",
    "last_read":"Última lectura","add_sensor":"Añadir sensor","sensor_ok":"Sensor guardado",
    "mqtt_cfg":"🔌 Configuración MQTT","mqtt_broker":"Broker MQTT","mqtt_port":"Puerto",
    "mqtt_topic":"Topic base","mqtt_active":"MQTT activo","no_signal":"🔴 Sin señal",
    "online":"🟢 Online","never":"Nunca",
    "iban_hint":"ES00 0000 0000 0000 0000 0000",
    "new_year":"Crear nuevo año","import_csv":"📂 Importar vecinos desde Excel",
    "csv_format":"Columnas: nombre (obligatorio), direccion, email, telefono, iban",
    "import_ok":"vecinos importados correctamente","email_from":"Email remitente",
    "email_pass":"Contraseña del email","email_cfg":"Configuración de email",
    "email_hint":"Gmail: usa contraseña de aplicación (no tu contraseña normal)",
    "facturas_saved":"Facturas guardadas en servidor",
    "sent_saved":"Factura enviada y guardada en servidor",
    "view_saved":"📁 Ver facturas guardadas","no_facturas":"No hay facturas guardadas aún.",
  },
  "ca":{
    "months":["Gener","Febrer","Març","Abril","Maig","Juny",
              "Juliol","Agost","Setembre","Octubre","Novembre","Desembre"],
    "menu_consums":"📊 Consums","menu_veins":"👥 Veïns",
    "menu_fact":"🧾 Facturació","menu_cfg":"⚙️ Configuració",
    "menu_myfact":"🧾 Les meves factures","menu_mycons":"📊 Els meus consums",
    "login_title":"Iniciar sessió","login_user":"Usuari","login_pass":"Contrasenya",
    "login_btn":"Entrar","login_err":"Usuari o contrasenya incorrectes","logout":"Tancar sessió",
    "year":"Any","month":"Mes","neighbor":"Veí","total":"Total",
    "download_pdf":"⬇️ Descarregar PDF","send_all":"📨 Enviar totes per email",
    "save":"Guardar","add":"Afegir","name":"Nom","address":"Adreça",
    "email":"Email","no_email":"✉️ Sense email","list":"Llistat",
    "add_neighbor":"Afegir veí","del_neighbor":"Eliminar veí",
    "register_reading":"Registrar lectura","tariff":"Tarifa €/m³",
    "fixed_fee":"Quota fixa €","vat":"IVA %","saved_ok":"✅ Guardat correctament",
    "sent_ok":"factures enviades","no_data":"Sense dades per aquest període.",
    "view_per_neighbor":"Per veí","view_accumulated":"Acumulat total",
    "bar_chart":"Comparativa mensual","invoice_period":"Factures de",
    "total_billed":"Total facturat","concept":"Concepte",
    "quantity":"Quantitat","unit_price":"Preu unitari","amount":"Import",
    "fixed_service":"Quota fixa de servei","tax_base":"Base imposable",
    "total_pay":"TOTAL A PAGAR",
    "thanks":"Gràcies per la seva col·laboració en l'ús responsable de l'aigua.",
    "payment_info":"Pagament: domiciliació bancària o transferència",
    "search_neighbor":"🔍 Cercar veí...","role_admin":"Administrador","role_vecino":"Veí",
    "client":"Client","issue_date":"Data emissió","invoice_number":"Nº Factura",
    "send_single":"Enviar factura individual","send_btn":"📨 Enviar",
    "users":"Usuaris","edit_user":"Editar usuari","del_user":"Eliminar usuari",
    "add_user":"Afegir usuari","new_pw":"Nova contrasenya",
    "export_list":"⬇️ Exportar Excel","phone":"Telèfon","edit_neighbor":"Editar veí",
    "export_consumption":"⬇️ CSV consums","export_chart":"⬇️ PNG gràfica",
    "annual_evolution":"Evolució anual","avg_monthly":"Consum mig mensual",
    "my_consumption":"El meu consum","cost_per_m3":"Tarifa €/m³",
    "confirm_delete":"⚠️ Confirmo que vull eliminar aquest registre",
    "prev_reading":"Lectura anterior","curr_reading":"Lectura actual",
    "consumption":"Consum calculat",
    "reading_error":"⚠️ La lectura actual no pot ser menor que l'anterior.",
    "backup_db":"⬇️ Backup base de dades","iban":"IBAN",
    "sensors_tab":"📡 Sensors","sensor_id":"ID Sensor","pulsos_m3":"Polsos/m³",
    "last_read":"Última lectura","add_sensor":"Afegir sensor","sensor_ok":"Sensor desat",
    "mqtt_cfg":"🔌 Configuració MQTT","mqtt_broker":"Broker MQTT","mqtt_port":"Port",
    "mqtt_topic":"Topic base","mqtt_active":"MQTT actiu","no_signal":"🔴 Sense senyal",
    "online":"🟢 Online","never":"Mai",
    "iban_hint":"ES00 0000 0000 0000 0000 0000",
    "new_year":"Crear nou any","import_csv":"📂 Importar veïns des d'Excel",
    "csv_format":"Columnes: nombre (obligatori), direccion, email, telefono, iban",
    "import_ok":"veïns importats correctament","email_from":"Email remitent",
    "email_pass":"Contrasenya de l'email","email_cfg":"Configuració d'email",
    "email_hint":"Gmail: utilitza contrasenya d'aplicació (no la teva contrasenya normal)",
    "facturas_saved":"Factures guardades al servidor",
    "sent_saved":"Factura enviada i guardada al servidor",
    "view_saved":"📁 Veure factures guardades","no_facturas":"No hi ha factures guardades encara.",
  }
}

def T(k): return LANG[st.session_state.get("lang","es")].get(k,k)
def months(): return T("months")

# ── BASE DE DATOS ─────────────────────────────────────────────────────────────
@contextmanager
def db():
    con = sqlite3.connect(DB)
    con.row_factory = sqlite3.Row
    try:
        yield con
        con.commit()
    except Exception:
        con.rollback()
        raise
    finally:
        con.close()

def hash_pw(pw):
    salt = secrets.token_hex(16)
    h = hashlib.sha256((salt+pw).encode()).hexdigest()
    return f"sha256${salt}${h}"

def verify_pw(pw, stored):
    try:
        parts = stored.split("$")
        if len(parts)==3 and parts[0]=="sha256":
            return hashlib.sha256((parts[1]+pw).encode()).hexdigest()==parts[2]
        return hashlib.sha256(pw.encode()).hexdigest()==stored
    except:
        return False

def init_db():
    with db() as con:
        con.execute("""CREATE TABLE IF NOT EXISTS users(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE, password TEXT, role TEXT,
            vecino_id INTEGER DEFAULT NULL)""")
        con.execute("""CREATE TABLE IF NOT EXISTS vecinos(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT, direccion TEXT, email TEXT,
            tiene_email INTEGER DEFAULT 1,
            telefono TEXT DEFAULT '', iban TEXT DEFAULT '')""")
        con.execute("""CREATE TABLE IF NOT EXISTS consumos(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            vecino_id INTEGER, anyo INTEGER, mes INTEGER,
            m3 REAL, lectura_actual REAL DEFAULT NULL,
            fuente TEXT DEFAULT 'manual')""")
        con.execute("""CREATE TABLE IF NOT EXISTS sensores(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            sensor_id TEXT UNIQUE NOT NULL,
            vecino_id INTEGER,
            pulsos_m3 INTEGER DEFAULT 100,
            activo INTEGER DEFAULT 1,
            ultima_lectura TEXT DEFAULT '',
            ultimo_pulso INTEGER DEFAULT 0)""")
        con.execute("""CREATE TABLE IF NOT EXISTS tarifas(
            id INTEGER PRIMARY KEY,
            precio_m3 REAL DEFAULT 0.85, cuota_fija REAL DEFAULT 5.0,
            iva REAL DEFAULT 0.10,
            entidad TEXT DEFAULT 'Ajuntament de Les Lloses',
            direccion TEXT DEFAULT 'C/ Major s/n · 08512 Les Lloses',
            contacto TEXT DEFAULT 'aigues@leslloses.cat  |  938 000 000',
            iban TEXT DEFAULT 'ES00 0000 0000 0000 0000 0000',
            email_from TEXT DEFAULT '',
            email_pass TEXT DEFAULT '',
            smtp_host TEXT DEFAULT '',
            smtp_port INTEGER DEFAULT 587,
            smtp_ssl INTEGER DEFAULT 0,
            smtp_no_verify INTEGER DEFAULT 0)""")
        con.execute("""INSERT OR IGNORE INTO tarifas
            (id,precio_m3,cuota_fija,iva,entidad,direccion,contacto,iban,
             email_from,email_pass,smtp_host,smtp_port,smtp_ssl,smtp_no_verify)
            VALUES (1,0.85,5.0,0.10,'Ajuntament de Les Lloses',
            'C/ Major s/n · 08512 Les Lloses',
            'aigues@leslloses.cat  |  938 000 000',
            'ES00 0000 0000 0000 0000 0000','','','',587,0,0)""")
        con.execute("INSERT OR IGNORE INTO users(username,password,role) VALUES(?,?,?)",
                    ("admin",hash_pw("admin123"),"admin"))
        # Retrocompatibilidad columnas
        for tbl,col,typ in [
            ("vecinos","telefono","TEXT DEFAULT ''"),
            ("vecinos","iban","TEXT DEFAULT ''"),
            ("consumos","lectura_actual","REAL DEFAULT NULL"),
            ("tarifas","email_from","TEXT DEFAULT ''"),
            ("tarifas","email_pass","TEXT DEFAULT ''"),
            ("tarifas","smtp_host","TEXT DEFAULT ''"),
            ("tarifas","smtp_port","INTEGER DEFAULT 587"),
            ("tarifas","smtp_ssl","INTEGER DEFAULT 0"),
            ("tarifas","smtp_no_verify","INTEGER DEFAULT 0"),
            ("tarifas","mqtt_broker","TEXT DEFAULT 'localhost'"),
            ("tarifas","mqtt_port","INTEGER DEFAULT 1883"),
            ("tarifas","mqtt_topic","TEXT DEFAULT 'leslloses/#'"),
            ("tarifas","mqtt_activo","INTEGER DEFAULT 0"),
        ]:
            try: con.execute(f"ALTER TABLE {tbl} ADD COLUMN {col} {typ}")
            except: pass
        # Demo si no hay vecinos
        if con.execute("SELECT COUNT(*) FROM vecinos").fetchone()[0]==0:
            _insert_demo(con)

def _insert_demo(con):
    import random; random.seed(42)
    demo=[
        ("Joan Puig","Carrer Major 1","joan@exemple.com",1,"600111001",""),
        ("Maria Sala","Carrer Major 3","maria@exemple.com",1,"600111002",""),
        ("Pere Font","Carrer del Pi 2","",0,"",""),
        ("Anna Roca","Plaça Església 4","anna@exemple.com",1,"600111004",""),
        ("Jordi Mas","Carrer Nou 7","jordi@exemple.com",1,"600111005",""),
        ("Rosa Vidal","Carrer del Molí 5","",0,"",""),
        ("Lluís Serra","Carrer Major 12","lluis@exemple.com",1,"600111007",""),
        ("Núria Bosch","Carrer Bosc 3","nuria@exemple.com",1,"600111008",""),
        ("Miquel Tort","Carrer Alt 8","",0,"",""),
        ("Teresa Vila","Plaça Nova 2","teresa@exemple.com",1,"600111010",""),
    ]
    con.executemany(
        "INSERT INTO vecinos(nombre,direccion,email,tiene_email,telefono,iban) VALUES(?,?,?,?,?,?)",
        demo)
    ids=[r[0] for r in con.execute("SELECT id FROM vecinos").fetchall()]
    for vid in ids:
        for mes in range(1,13):
            m3=round(random.uniform(4,18),2)
            con.execute("INSERT INTO consumos(vecino_id,anyo,mes,m3,fuente) VALUES(?,?,?,?,?)",
                        (vid,2025,mes,m3,"demo"))
        row=con.execute("SELECT nombre FROM vecinos WHERE id=?",(vid,)).fetchone()
        uname=row[0].lower().replace(" ",".").replace("à","a").replace("è","e")\
            .replace("ú","u").replace("ï","i").replace("ü","u").replace("·","")[:14]
        con.execute("INSERT OR IGNORE INTO users(username,password,role,vecino_id) VALUES(?,?,?,?)",
                    (uname,hash_pw("vecino123"),"vecino",vid))

def check_login(u,p):
    with db() as con:
        r=con.execute("SELECT id,role,vecino_id,password FROM users WHERE username=?",(u,)).fetchone()
    if r and verify_pw(p,r["password"]):
        return r["id"],r["role"],r["vecino_id"]
    return None

def get_tar():
    with db() as con:
        r=con.execute("SELECT * FROM tarifas WHERE id=1").fetchone()
    return dict(r) if r else {}

# ── IMPORTAR CSV ──────────────────────────────────────────────────────────────
def import_excel(uploaded_file):
    """
    Importa vecinos desde Excel (.xlsx / .xls).
    Columnas: nombre, direccion, email, telefono, iban
    Solo 'nombre' es obligatorio.
    """
    try:
        fname = getattr(uploaded_file, "name", "")
        if fname.lower().endswith(".xls") and not fname.lower().endswith(".xlsx"):
            df = pd.read_excel(uploaded_file, dtype=str, engine="xlrd").fillna("")
        else:
            df = pd.read_excel(uploaded_file, dtype=str, engine="openpyxl").fillna("")
        # Normalizar nombres de columnas
        df.columns = [c.strip().lower() for c in df.columns]
        if "nombre" not in df.columns:
            return 0, "Falta columna obligatoria: 'nombre'"
        count = 0
        with db() as con:
            for _, row in df.iterrows():
                nombre = str(row.get("nombre", "")).strip()
                if not nombre:
                    continue
                direccion = str(row.get("direccion", "")).strip()
                email = str(row.get("email", "")).strip()
                telefono = str(row.get("telefono", "")).strip()
                iban = str(row.get("iban", "")).strip()
                tiene_email = 1 if email else 0
                ex = con.execute("SELECT id FROM vecinos WHERE nombre=?", (nombre,)).fetchone()
                if ex:
                    con.execute(
                        "UPDATE vecinos SET direccion=?,email=?,tiene_email=?,telefono=?,iban=? WHERE id=?",
                        (direccion, email, tiene_email, telefono, iban, ex[0]))
                else:
                    con.execute(
                        "INSERT INTO vecinos(nombre,direccion,email,tiene_email,telefono,iban) VALUES(?,?,?,?,?,?)",
                        (nombre, direccion, email, tiene_email, telefono, iban))
                    vid = con.execute("SELECT id FROM vecinos WHERE nombre=?", (nombre,)).fetchone()[0]
                    uname = nombre.lower().replace(" ", ".").replace("à", "a") \
                        .replace("è", "e").replace("ú", "u").replace("ï", "i") \
                        .replace("ü", "u").replace("·", "")[:14]
                    con.execute("INSERT OR IGNORE INTO users(username,password,role,vecino_id) VALUES(?,?,?,?)",
                                (uname, hash_pw("vecino123"), "vecino", vid))
                count += 1
        return count, ""
    except Exception as e:
        return 0, str(e)

# ── EMAIL SIMPLE ──────────────────────────────────────────────────────────────
def send_email(to_addr, nombre, nf, mes, anyo, pdf_bytes, from_addr, from_pass, lang="es",
               smtp_host="", smtp_port=587, smtp_ssl=False, smtp_no_verify=False):
    """
    Envia email amb adjunt PDF.
    - Si smtp_host buit, detecta servidor automaticament pel domini.
    - smtp_no_verify=True desactiva verificacio SSL (hostings amb certificat erroni).
    """
    L=LANG[lang]
    try:
        msg=MIMEMultipart()
        msg["From"]=from_addr
        msg["To"]=to_addr
        msg["Subject"]="Factura Aigues - {} {}".format(L["months"][mes-1],anyo)
        body=("Benvolgut/da {},\n\nAdjuntem la factura {} del mes de {} {}.\n\n"
              "Gracies per la seva collaboracio.\n\nAjuntament de Les Lloses").format(
                  nombre,nf,L["months"][mes-1],anyo)
        msg.attach(MIMEText(body,"plain","utf-8"))
        att=MIMEBase("application","octet-stream")
        att.set_payload(pdf_bytes)
        encoders.encode_base64(att)
        att.add_header("Content-Disposition","attachment; filename={}.pdf".format(nf))
        msg.attach(att)
        # Determinar servidor SMTP
        if smtp_host:
            host=smtp_host
            port=int(smtp_port) if smtp_port else 587
            use_ssl=bool(smtp_ssl)
        else:
            domain=from_addr.split("@")[-1].lower() if "@" in from_addr else ""
            smtp_map={
                "gmail.com":("smtp.gmail.com",587,False),
                "googlemail.com":("smtp.gmail.com",587,False),
                "yahoo.com":("smtp.mail.yahoo.com",587,False),
                "yahoo.es":("smtp.mail.yahoo.com",587,False),
                "hotmail.com":("smtp.office365.com",587,False),
                "outlook.com":("smtp.office365.com",587,False),
                "live.com":("smtp.office365.com",587,False),
                "icloud.com":("smtp.mail.me.com",587,False),
                "ovh.net":("ssl0.ovh.net",465,True),
                "ovh.es":("ssl0.ovh.net",465,True),
            }
            host,port,use_ssl=smtp_map.get(domain,("smtp."+domain,587,False))
        # Contexto SSL
        if smtp_no_verify:
            ctx=ssl_lib.create_default_context()
            ctx.check_hostname=False
            ctx.verify_mode=ssl_lib.CERT_NONE
        else:
            ctx=ssl_lib.create_default_context()
        if use_ssl:
            with smtplib.SMTP_SSL(host,port,context=ctx,timeout=30) as s:
                s.login(from_addr,from_pass)
                s.sendmail(from_addr,to_addr,msg.as_string())
        else:
            with smtplib.SMTP(host,port,timeout=30) as s:
                s.ehlo(); s.starttls(context=ctx); s.ehlo()
                s.login(from_addr,from_pass)
                s.sendmail(from_addr,to_addr,msg.as_string())
        return True,""
    except Exception as e:
        return False,str(e)

# ── GUARDAR FACTURA EN DISCO ──────────────────────────────────────────────────
def save_factura_disk(pdf_bytes, fname, anyo, mes):
    """Guarda PDF en data/facturas/YYYY/MM/"""
    folder=FACTURAS_DIR/str(anyo)/f"{mes:02d}"
    folder.mkdir(parents=True,exist_ok=True)
    dest=folder/fname
    with open(dest,"wb") as f:
        f.write(pdf_bytes)
    return str(dest)

def list_facturas(anyo=None, mes=None):
    """Lista facturas guardadas en disco."""
    results=[]
    base=FACTURAS_DIR
    if anyo:
        base=base/str(anyo)
        if mes: base=base/f"{mes:02d}"
    if base.exists():
        for p in sorted(base.rglob("*.pdf")):
            parts=p.relative_to(FACTURAS_DIR).parts
            results.append({
                "archivo":p.name,
                "anyo":parts[0] if len(parts)>0 else "",
                "mes":parts[1] if len(parts)>1 else "",
                "ruta":str(p),
                "size":f"{p.stat().st_size/1024:.1f} KB",
            })
    return results

# ── PDF ────────────────────────────────────────────────────────────────────────
def make_pdf(row,mes,anyo,m3,pm3,cf,iva,lang="es",entidad="",dir_ent="",contacto="",iban_ent=""):
    L=LANG[lang]
    buf=BytesIO()
    c=rl_canvas.Canvas(buf,pagesize=A4)
    W,H=A4
    nf="FAC-{}-{:02d}-{:03d}".format(anyo,mes,int(row.get("id",0)))
    nombre_vec=str(row.get("nombre","vecino")).replace(" ","_")
    fname="{}-{}_{}.pdf".format(L["months"][mes-1],anyo,nombre_vec)

    # Cabecera oscura
    c.setFillColor(colors.HexColor("#0d3b6e"))
    c.rect(0,H-80,W,80,fill=1,stroke=0)
    c.setFillColor(colors.HexColor("#1a9e5c"))
    c.rect(0,H-80,8,80,fill=1,stroke=0)
    # Logo texto
    c.setFillColor(colors.white)
    c.setFont("Helvetica-Bold",15)
    c.drawString(22,H-32,entidad or "Ajuntament de Les Lloses")
    c.setFont("Helvetica",9)
    c.setFillColor(colors.HexColor("#a8c8e8"))
    c.drawString(22,H-48,"Servei Municipal d'Aigues")
    c.drawString(22,H-60,dir_ent or "C/ Major s/n · 08512 Les Lloses")
    c.drawString(22,H-72,contacto or "aigues@leslloses.cat  |  938 000 000")
    # Nº factura
    c.setFillColor(colors.white)
    c.setFont("Helvetica-Bold",10)
    c.drawRightString(W-20,H-32,"{}: {}".format(L["invoice_number"],nf))
    c.setFont("Helvetica",9)
    c.setFillColor(colors.HexColor("#a8c8e8"))
    c.drawRightString(W-20,H-48,"{}: {}".format(L["issue_date"],date.today().strftime("%d/%m/%Y")))

    # Banda período
    c.setFillColor(colors.HexColor("#1a6fc4"))
    c.rect(0,H-108,W,26,fill=1,stroke=0)
    c.setFillColor(colors.white)
    c.setFont("Helvetica-Bold",11)
    c.drawCentredString(W/2,H-99,"{} {} {}".format(
        L["invoice_period"].upper(),L["months"][mes-1].upper(),anyo))

    # Caja cliente
    c.setFillColor(colors.HexColor("#f8fafc"))
    c.roundRect(20,H-200,W-40,84,6,fill=1,stroke=0)
    c.setStrokeColor(colors.HexColor("#e2e8f0"))
    c.setLineWidth(0.5)
    c.roundRect(20,H-200,W-40,84,6,fill=0,stroke=1)
    c.setFillColor(colors.HexColor("#1a6fc4"))
    c.setFont("Helvetica-Bold",8)
    c.drawString(30,H-124,L["client"].upper())
    c.setStrokeColor(colors.HexColor("#1a6fc4"))
    c.setLineWidth(1)
    c.line(30,H-128,W/2,H-128)
    c.setFillColor(colors.HexColor("#0f172a"))
    c.setFont("Helvetica-Bold",13)
    c.drawString(30,H-145,str(row.get("nombre","")))
    c.setFont("Helvetica",10)
    c.setFillColor(colors.HexColor("#64748b"))
    c.drawString(30,H-160,str(row.get("direccion","") or ""))
    if row.get("email"):
        c.drawString(30,H-174,str(row["email"]))
    if row.get("telefono"):
        c.drawString(30,H-188,str(row["telefono"]))

    # Tabla de conceptos
    sub=round(m3*pm3+cf,2); iva_e=round(sub*iva,2); total=round(sub+iva_e,2)
    col_w=[220,75,110,90]; x_table=(W-sum(col_w))/2
    data=[
        [L["concept"],L["quantity"],L["unit_price"],L["amount"]],
        ["Consum d'aigua  {} {}".format(L["months"][mes-1],anyo),
         "{:.2f} m³".format(m3),"{:.4f} €/m³".format(pm3),"{:.2f} €".format(m3*pm3)],
        [L["fixed_service"],"1 ud","{:.2f} €".format(cf),"{:.2f} €".format(cf)],
        ["","",L["tax_base"],"{:.2f} €".format(sub)],
        ["","","IVA ({}%)".format(int(iva*100)),"{:.2f} €".format(iva_e)],
        ["","",L["total_pay"],"{:.2f} €".format(total)],
    ]
    ts=TableStyle([
        ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#0d3b6e")),
        ("TEXTCOLOR",(0,0),(-1,0),colors.white),
        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
        ("FONTSIZE",(0,0),(-1,-1),9),
        ("TOPPADDING",(0,0),(-1,-1),6),("BOTTOMPADDING",(0,0),(-1,-1),6),
        ("ALIGN",(1,0),(-1,-1),"RIGHT"),("ALIGN",(0,0),(0,-1),"LEFT"),
        ("FONTNAME",(0,1),(-1,-1),"Helvetica"),
        ("GRID",(0,0),(-1,2),0.3,colors.HexColor("#e2e8f0")),
        ("ROWBACKGROUNDS",(0,1),(-1,2),[colors.white,colors.HexColor("#f8fafc")]),
        ("LINEABOVE",(0,3),(-1,3),1,colors.HexColor("#e2e8f0")),
        ("FONTNAME",(2,3),(-1,4),"Helvetica"),
        ("FONTNAME",(2,5),(-1,5),"Helvetica-Bold"),
        ("FONTSIZE",(2,5),(-1,5),10),
        ("BACKGROUND",(0,5),(-1,5),colors.HexColor("#0d3b6e")),
        ("TEXTCOLOR",(0,5),(-1,5),colors.white),
        ("LINEABOVE",(0,5),(-1,5),1.5,colors.HexColor("#1a9e5c")),
    ])
    t=Table(data,colWidths=col_w)
    t.setStyle(ts)
    t.wrapOn(c,W,H)
    t.drawOn(c,x_table,H-420)

    # Caja IBAN
    c.setFillColor(colors.HexColor("#f0f7ff"))
    c.roundRect(20,H-460,W-40,38,5,fill=1,stroke=0)
    c.setStrokeColor(colors.HexColor("#1a6fc4"))
    c.setLineWidth(0.5)
    c.roundRect(20,H-460,W-40,38,5,fill=0,stroke=1)
    c.setFillColor(colors.HexColor("#1a6fc4"))
    c.setFont("Helvetica-Bold",8)
    c.drawString(30,H-435,"IBAN: {}".format(iban_ent or "ES00 0000 0000 0000 0000 0000"))
    c.setFillColor(colors.HexColor("#64748b"))
    c.setFont("Helvetica",8)
    c.drawString(30,H-449,L["payment_info"])

    # Pie
    c.setFillColor(colors.HexColor("#0d3b6e"))
    c.rect(0,0,W,44,fill=1,stroke=0)
    c.setFillColor(colors.HexColor("#1a9e5c"))
    c.rect(0,42,W,3,fill=1,stroke=0)
    c.setFillColor(colors.white)
    c.setFont("Helvetica",8)
    c.drawString(22,28,L["thanks"])
    c.setFillColor(colors.HexColor("#a8c8e8"))
    c.setFont("Helvetica",7)
    c.drawString(22,14,"{} · {}".format(entidad or "Les Lloses",date.today().strftime("%d/%m/%Y")))
    c.drawRightString(W-22,14,"Ref: {}".format(nf))
    c.save()
    buf.seek(0)
    return buf,nf,total,fname

# ── CHARTS ─────────────────────────────────────────────────────────────────────
def chart_line(df_v,vname,lang):
    MES=LANG[lang]["months"]
    df_v=df_v.sort_values("mes")
    df_v["Mes"]=df_v["mes"].apply(lambda x:MES[x-1])
    fig=go.Figure()
    fig.add_trace(go.Scatter(
        x=df_v["Mes"],y=df_v["m3"],mode="lines+markers",
        line=dict(color=P["blue"],width=3),
        marker=dict(size=9,color=P["blue"],line=dict(color="white",width=2)),
        fill="tozeroy",fillcolor="rgba(26,111,196,.07)",name=vname))
    # Línea media
    media=df_v["m3"].mean()
    fig.add_hline(y=media,line_dash="dash",line_color=P["amber"],
                  annotation_text=f"Media: {media:.1f} m³",
                  annotation_position="bottom right")
    fig.update_layout(
        paper_bgcolor="white",plot_bgcolor="#f8fafc",
        font=dict(color=P["dark"]),height=290,
        xaxis=dict(gridcolor=P["border"]),
        yaxis=dict(gridcolor=P["border"],title="m³"),
        margin=dict(l=50,r=20,t=20,b=50),showlegend=False)
    return fig

def chart_bar(df_mes,lang):
    fig=go.Figure(go.Bar(
        x=df_mes["nombre"],y=df_mes["m3"],
        marker_color=P["blue"],opacity=.85,
        text=df_mes["m3"].round(1),textposition="outside",
        marker_line_color=P["lblue"],marker_line_width=1))
    fig.update_layout(
        paper_bgcolor="white",plot_bgcolor="#f8fafc",
        font=dict(color=P["dark"]),height=290,
        xaxis=dict(gridcolor=P["border"]),
        yaxis=dict(gridcolor=P["border"],title="m³"),
        margin=dict(l=50,r=20,t=20,b=80))
    return fig

def chart_acumulat(df_all,lang):
    MES=LANG[lang]["months"]
    tot=df_all.groupby("mes")["m3"].sum().reset_index().sort_values("mes")
    tot["Mes"]=tot["mes"].apply(lambda x:MES[x-1])
    fig=go.Figure()
    fig.add_trace(go.Bar(x=tot["Mes"],y=tot["m3"],marker_color=P["blue"],opacity=.7,name="m³ mensual"))
    fig.add_trace(go.Scatter(x=tot["Mes"],y=tot["m3"].cumsum(),mode="lines+markers",
        line=dict(color=P["green"],width=3),marker=dict(size=7,color=P["green"]),name="Acumulat"))
    fig.update_layout(
        paper_bgcolor="white",plot_bgcolor="#f8fafc",
        font=dict(color=P["dark"]),height=300,
        xaxis=dict(gridcolor=P["border"]),yaxis=dict(gridcolor=P["border"],title="m³"),
        legend=dict(bgcolor="white",bordercolor=P["border"],borderwidth=1),
        margin=dict(l=50,r=20,t=20,b=50))
    return fig

def buscar_vecino(vdf,key_suffix=""):
    q=st.text_input(T("search_neighbor"),key="search_"+key_suffix)
    if q:
        mask=vdf["nombre"].str.lower().str.contains(q.lower(),na=False)
        return vdf[mask]
    return vdf

# ══════════════════════════════════════════════════════════════════════════════
# INICIO
# ══════════════════════════════════════════════════════════════════════════════
init_db()
ensure_mqtt_running()
if "lang" not in st.session_state: st.session_state["lang"]="es"

# ── LOGIN ─────────────────────────────────────────────────────────────────────
if "role" not in st.session_state:
    col1,col2,col3=st.columns([1,2,1])
    with col2:
        # Logo
        st.markdown("""
        <div style="text-align:center;padding:2rem 0 1.5rem">
          <div style="font-size:3.5rem;margin-bottom:.5rem">💧</div>
          <div style="color:#1a6fc4;font-size:1.6rem;font-weight:800;letter-spacing:.5px">LES LLOSES</div>
          <div style="color:#64748b;font-size:.85rem;margin-top:.2rem">Gestió Municipal d'Aigues · Booster MT</div>
        </div>""",unsafe_allow_html=True)

        # Selector idioma
        lang_opt=st.radio("🌐",["Castellano","Català"],horizontal=True,key="lang_login",label_visibility="collapsed")
        st.session_state["lang"]="ca" if lang_opt=="Català" else "es"

        # Caja de login con fondo blanco explícito
        st.markdown("""
        <div style="background:#fff;border:1.5px solid #cbd5e1;border-radius:16px;
             padding:2rem;box-shadow:0 4px 20px rgba(0,0,0,.08);margin-top:.5rem">
        """,unsafe_allow_html=True)

        st.markdown(f"<h2 style='color:{P['blue']};text-align:center;margin-bottom:1rem'>"
                    f"🔐 {T('login_title')}</h2>",unsafe_allow_html=True)

        with st.form("lf",clear_on_submit=False):
            u=st.text_input(f"👤 {T('login_user')}",placeholder="admin")
            p_in=st.text_input(f"🔑 {T('login_pass')}",type="password",placeholder="••••••••")
            sub=st.form_submit_button(T("login_btn"),use_container_width=True)

        if sub:
            r=check_login(u,p_in)
            if r:
                st.session_state["user_id"]=r[0]
                st.session_state["role"]=r[1]
                st.session_state["vecino_id"]=r[2]
                st.session_state["username"]=u
                st.rerun()
            else:
                st.error(T("login_err"))

        st.markdown("</div>",unsafe_allow_html=True)
        st.markdown(f"""
        <div class="card" style="margin-top:1rem;text-align:center;background:#f8fafc">
          <small style="color:#64748b"><b>Demo:</b> admin / admin123 &nbsp;|&nbsp; joan.puig / vecino123</small>
        </div>""",unsafe_allow_html=True)
    st.stop()

# ── SIDEBAR ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown(f"""
    <div style="padding:.8rem 0 .4rem">
      <span style="font-size:1.4rem">💧</span>
      <span style="color:{P['blue']};font-size:1.1rem;font-weight:800"> LES LLOSES</span><br>
      <span style="color:{P['gray']};font-size:.75rem">Booster MT · Aigues</span>
    </div>
    <hr style="border-color:{P['border']};margin:.6rem 0">
    """,unsafe_allow_html=True)

    lang_opt=st.radio("🌐",["Castellano","Català"],horizontal=True,
        index=0 if st.session_state["lang"]=="es" else 1,key="lang_sw")
    st.session_state["lang"]="ca" if lang_opt=="Català" else "es"

    role=st.session_state["role"]
    if role=="admin":
        menu_opts=[T("menu_consums"),T("menu_veins"),T("menu_fact"),T("menu_cfg")]
    else:
        menu_opts=[T("menu_mycons"),T("menu_myfact")]
    menu=st.radio("",menu_opts,label_visibility="collapsed")

    st.markdown(f"<hr style='border-color:{P['border']};margin:.6rem 0'>",unsafe_allow_html=True)
    badge="badge-admin" if role=="admin" else "badge-vecino"
    st.markdown(
        f"<span class='{badge}'>{T('role_'+role)}</span> "
        f"<small style='color:{P['gray']}'> {st.session_state['username']}</small>",
        unsafe_allow_html=True)
    st.markdown("<br>",unsafe_allow_html=True)
    if st.button(T("logout"),use_container_width=True):
        for k in ["role","user_id","vecino_id","username"]:
            st.session_state.pop(k,None)
        st.rerun()

role=st.session_state["role"]

# ══ ADMIN: CONSUMOS ════════════════════════════════════════════════════════════
if role=="admin" and menu==T("menu_consums"):
    st.title(T("menu_consums"))
    with db() as con:
        anyos=pd.read_sql("SELECT DISTINCT anyo FROM consumos ORDER BY anyo DESC",con)["anyo"].tolist()
        vdf=pd.read_sql("SELECT * FROM vecinos",con)
    c1,c2=st.columns(2)
    anyo=c1.selectbox(T("year"),anyos or [date.today().year])
    vista=c2.radio("",[T("view_per_neighbor"),T("view_accumulated")],horizontal=True)
    with db() as con:
        df=pd.read_sql(
            "SELECT v.nombre,c.mes,c.m3 FROM consumos c JOIN vecinos v ON c.vecino_id=v.id WHERE c.anyo=? ORDER BY v.nombre,c.mes",
            con,params=(anyo,))
    MES=months()
    if df.empty:
        st.info(T("no_data"))
    elif vista==T("view_per_neighbor"):
        vdf_f=buscar_vecino(vdf,"consums")
        lista=vdf_f["nombre"].tolist() if not vdf_f.empty else vdf["nombre"].tolist()
        if lista:
            ca,cb=st.columns(2)
            vsel=ca.selectbox(T("neighbor"),lista,key="vsel_c")
            mes_sel=cb.selectbox(T("month"),list(range(1,13)),format_func=lambda x:MES[x-1],key="mes_c")
            df_v=df[df["nombre"]==vsel]
            if not df_v.empty:
                df_vm=df_v[df_v["mes"]==mes_sel]
                total_mes=float(df_vm["m3"].sum()) if not df_vm.empty else 0.0
                media_anual=df_v["m3"].mean()
                k1,k2,k3,k4=st.columns(4)
                k1.metric(f"Total {anyo}",f"{df_v['m3'].sum():.1f} m³")
                k2.metric(T("avg_monthly"),f"{media_anual:.1f} m³/mes")
                k3.metric(MES[mes_sel-1],f"{total_mes:.1f} m³")
                tar=get_tar()
                coste_medio=round(media_anual*float(tar.get("precio_m3",0.85))+float(tar.get("cuota_fija",5.0)),2)
                k4.metric("Coste medio/mes",f"~{coste_medio:.2f} €")
                st.subheader(f"{T('annual_evolution')} – {vsel}")
                fig=chart_line(df_v.copy(),vsel,st.session_state["lang"])
                st.plotly_chart(fig,use_container_width=True)
                ex1,ex2=st.columns(2)
                df_exp=df_v.copy(); df_exp["Mes"]=df_exp["mes"].apply(lambda x:MES[x-1])
                csv=df_exp[["Mes","m3"]].rename(columns={"m3":"m³"}).to_csv(index=False).encode()
                ex1.download_button(T("export_consumption"),data=csv,
                    file_name=f"consumos_{vsel.replace(' ','_')}_{anyo}.csv",mime="text/csv")
                try:
                    import plotly.io as pio
                    img=pio.to_image(fig,format="png",width=900,height=350,scale=2)
                    ex2.download_button(T("export_chart"),data=img,
                        file_name=f"grafica_{vsel.replace(' ','_')}_{anyo}.png",mime="image/png")
                except: pass
        st.subheader(T("bar_chart"))
        mes_b=st.selectbox(T("month"),list(range(1,13)),format_func=lambda x:MES[x-1],key="mes_bar")
        df_mb=df[df["mes"]==mes_b].sort_values("m3",ascending=False)
        if not df_mb.empty:
            st.plotly_chart(chart_bar(df_mb,st.session_state["lang"]),use_container_width=True)
        pivot=df.pivot_table(index="nombre",columns="mes",values="m3",aggfunc="sum")
        pivot.columns=[MES[c-1] for c in pivot.columns]
        pivot["TOTAL"]=pivot.sum(axis=1)
        pivot["MEDIA/MES"]=pivot.drop(columns="TOTAL").mean(axis=1).round(2)
        st.dataframe(pivot.round(2),use_container_width=True)
    else:
        st.subheader(T("view_accumulated"))
        st.plotly_chart(chart_acumulat(df.copy(),st.session_state["lang"]),use_container_width=True)
        tot=df.groupby("mes")["m3"].sum().reset_index().sort_values("mes")
        tot["Mes"]=tot["mes"].apply(lambda x:MES[x-1])
        k1,k2,k3=st.columns(3)
        k1.metric("Total anual",f"{tot['m3'].sum():.1f} m³")
        k2.metric(T("avg_monthly"),f"{tot['m3'].mean():.1f} m³/mes")
        k3.metric("Mes max.",MES[tot.loc[tot["m3"].idxmax(),"mes"]-1])
        tot["Acumulat"]=tot["m3"].cumsum().round(2)
        st.dataframe(tot[["Mes","m3","Acumulat"]].rename(columns={"m3":"m³"}),
                     use_container_width=True,hide_index=True)

# ══ ADMIN: VECINOS ════════════════════════════════════════════════════════════
elif role=="admin" and menu==T("menu_veins"):
    st.title(T("menu_veins"))
    with db() as con:
        vdf=pd.read_sql("SELECT * FROM vecinos",con)

    tabs=st.tabs([T("list"),T("add_neighbor"),T("edit_neighbor"),
                  T("del_neighbor"),T("register_reading"),T("import_csv"),T("users")])
    t1,t2,t_edit,t3,t4,t_csv,t5=tabs

    with t1:
        vdf_f=buscar_vecino(vdf,"list")
        d=vdf_f.copy()
        for col in ["telefono","iban"]:
            if col not in d.columns: d[col]=""
        d["Contacte"]=d["tiene_email"].apply(lambda x:"✅ Email" if x else "✉️ Postal")
        st.dataframe(d[["id","nombre","direccion","email","telefono","iban","Contacte"]].rename(
            columns={"id":"ID","nombre":T("name"),"direccion":T("address"),
                     "email":T("email"),"telefono":T("phone"),"iban":T("iban")}),
            use_container_width=True,hide_index=True)
        excel_buf=BytesIO()
        with pd.ExcelWriter(excel_buf,engine="openpyxl") as writer:
            d[["id","nombre","direccion","email","telefono","iban","Contacte"]].rename(
                columns={"id":"ID","nombre":T("name"),"direccion":T("address"),
                         "email":T("email"),"telefono":T("phone"),"iban":T("iban")}
            ).to_excel(writer,index=False,sheet_name="Vecinos")
        excel_buf.seek(0)
        st.download_button(T("export_list"),data=excel_buf.read(),file_name="vecinos.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    with t2:
        with st.form("fv"):
            n=st.text_input(T("name")); a=st.text_input(T("address"))
            e=st.text_input(T("email")); tlf=st.text_input(T("phone"))
            ib=st.text_input(T("iban"),placeholder=T("iban_hint"))
            if st.form_submit_button(T("add")) and n:
                with db() as con:
                    con.execute(
                        "INSERT INTO vecinos(nombre,direccion,email,tiene_email,telefono,iban) VALUES(?,?,?,?,?,?)",
                        (n,a,e,1 if e else 0,tlf,ib))
                    vid=con.execute("SELECT id FROM vecinos WHERE nombre=?",(n,)).fetchone()[0]
                    uname=n.lower().replace(" ",".").replace("à","a").replace("è","e")\
                        .replace("ú","u").replace("ï","i").replace("ü","u").replace("·","")[:14]
                    con.execute("INSERT OR IGNORE INTO users(username,password,role,vecino_id) VALUES(?,?,?,?)",
                                (uname,hash_pw("vecino123"),"vecino",vid))
                st.success(T("saved_ok")); st.rerun()

    with t_edit:
        with db() as con:
            vdf_e=pd.read_sql("SELECT * FROM vecinos",con)
        for col in ["telefono","iban"]:
            if col not in vdf_e.columns: vdf_e[col]=""
        vsel_e=st.selectbox(T("neighbor"),vdf_e["nombre"].tolist(),key="edit_v")
        row_e=vdf_e[vdf_e["nombre"]==vsel_e].iloc[0]
        with st.form("fedit_v"):
            en=st.text_input(T("name"),value=str(row_e["nombre"]))
            ea=st.text_input(T("address"),value=str(row_e.get("direccion","") or ""))
            ee=st.text_input(T("email"),value=str(row_e.get("email","") or ""))
            et=st.text_input(T("phone"),value=str(row_e.get("telefono","") or ""))
            eib=st.text_input(T("iban"),value=str(row_e.get("iban","") or ""))
            if st.form_submit_button(T("save")) and en:
                with db() as con:
                    con.execute(
                        "UPDATE vecinos SET nombre=?,direccion=?,email=?,tiene_email=?,telefono=?,iban=? WHERE id=?",
                        (en,ea,ee,1 if ee else 0,et,eib,int(row_e["id"])))
                st.success(T("saved_ok")); st.rerun()

    with t3:
        with db() as con:
            vdf2=pd.read_sql("SELECT * FROM vecinos",con)
        vsel_d=st.selectbox(T("neighbor"),vdf2["nombre"].tolist(),key="del_v")
        ok_del=st.checkbox(T("confirm_delete"),key="chk_del")
        if st.button("🗑️ "+T("del_neighbor"),type="primary",disabled=not ok_del):
            vid_d=int(vdf2[vdf2["nombre"]==vsel_d]["id"].values[0])
            with db() as con:
                con.execute("DELETE FROM vecinos WHERE id=?",(vid_d,))
                con.execute("DELETE FROM consumos WHERE vecino_id=?",(vid_d,))
                con.execute("DELETE FROM users WHERE vecino_id=?",(vid_d,))
            st.success(T("saved_ok")); st.rerun()

    with t4:
        with db() as con:
            vdf3=pd.read_sql("SELECT * FROM vecinos",con)
        with st.form("fc"):
            vs=st.selectbox(T("neighbor"),vdf3["nombre"].tolist())
            ay=st.number_input(T("year"),min_value=2020,max_value=2040,value=date.today().year)
            ms=st.selectbox(T("month"),list(range(1,13)),format_func=lambda x:months()[x-1])
            vid_f=int(vdf3[vdf3["nombre"]==vs]["id"].values[0]) if not vdf3.empty else 0
            st.markdown("---")
            with db() as con:
                prev=con.execute(
                    """SELECT m3,lectura_actual FROM consumos WHERE vecino_id=?
                       AND (anyo<? OR (anyo=? AND mes<?)) ORDER BY anyo DESC,mes DESC LIMIT 1""",
                    (vid_f,int(ay),int(ay),int(ms))).fetchone()
            prev_lect=float(prev["lectura_actual"]) if prev and prev["lectura_actual"] else None
            if prev_lect is not None:
                st.info(f"{T('prev_reading')}: **{prev_lect:.2f} m³**")
            curr_lect=st.number_input(T("curr_reading"),min_value=0.0,step=0.1)
            if prev_lect is None:
                m3_manual=st.number_input("m³ (si no hay lectura anterior)",min_value=0.0,step=0.1)
            if st.form_submit_button(T("save")):
                if prev_lect is not None:
                    diff=round(curr_lect-prev_lect,2)
                    if diff<0:
                        st.error(T("reading_error"))
                    else:
                        with db() as con:
                            ex=con.execute("SELECT id FROM consumos WHERE vecino_id=? AND anyo=? AND mes=?",
                                           (vid_f,int(ay),int(ms))).fetchone()
                            if ex:
                                con.execute("UPDATE consumos SET m3=?,lectura_actual=? WHERE id=?",
                                            (diff,curr_lect,ex[0]))
                            else:
                                con.execute("INSERT INTO consumos(vecino_id,anyo,mes,m3,lectura_actual,fuente) VALUES(?,?,?,?,?,?)",
                                            (vid_f,int(ay),int(ms),diff,curr_lect,"manual"))
                        st.success(T("saved_ok"))
                else:
                    with db() as con:
                        ex=con.execute("SELECT id FROM consumos WHERE vecino_id=? AND anyo=? AND mes=?",
                                       (vid_f,int(ay),int(ms))).fetchone()
                        if ex:
                            con.execute("UPDATE consumos SET m3=?,lectura_actual=? WHERE id=?",
                                        (m3_manual,curr_lect,ex[0]))
                        else:
                            con.execute("INSERT INTO consumos(vecino_id,anyo,mes,m3,lectura_actual,fuente) VALUES(?,?,?,?,?,?)",
                                        (vid_f,int(ay),int(ms),m3_manual,curr_lect,"manual"))
                    st.success(T("saved_ok"))

    # ── IMPORTAR EXCEL ────────────────────────────────────────────────────────
    with t_csv:
        st.subheader(T("import_csv"))
        st.info(f"📋 **{T('csv_format')}**\n\n"
                "El campo `nombre` es obligatorio. El resto son opcionales.\n\n"
                "Si el vecino ya existe (mismo nombre), se actualizan sus datos.")
        # Plantilla Excel descargable
        template_df = pd.DataFrame([
            ["Joan Puig",  "Carrer Major 1", "joan@exemple.com",  "600111001", "ES00 0000"],
            ["Maria Sala", "Carrer Major 3", "maria@exemple.com", "600111002", ""],
            ["Pere Font",  "Carrer del Pi 2","",                  "",          ""],
        ], columns=["nombre","direccion","email","telefono","iban"])
        tpl_buf = BytesIO()
        with pd.ExcelWriter(tpl_buf, engine="openpyxl") as writer:
            template_df.to_excel(writer, index=False, sheet_name="Vecinos")
        tpl_buf.seek(0)
        st.download_button("⬇️ Descargar plantilla Excel",
            data=tpl_buf.read(),
            file_name="plantilla_vecinos.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.markdown("---")
        uploaded = st.file_uploader(
            "Seleccionar archivo Excel (.xlsx o .xls)",
            type=["xlsx", "xls"],
            key="excel_upload")
        if uploaded:
            try:
                fname_up = getattr(uploaded, "name", "")
                if fname_up.lower().endswith(".xls") and not fname_up.lower().endswith(".xlsx"):
                    df_prev = pd.read_excel(uploaded, dtype=str, engine="xlrd").fillna("")
                else:
                    df_prev = pd.read_excel(uploaded, dtype=str, engine="openpyxl").fillna("")
                df_prev.columns = [c.strip().lower() for c in df_prev.columns]
                st.write("**Vista previa del Excel:**")
                st.dataframe(df_prev.head(10), use_container_width=True, hide_index=True)
                st.caption(f"{len(df_prev)} filas detectadas")
                uploaded.seek(0)
                if st.button("✅ Confirmar importación", type="primary"):
                    count, err = import_excel(uploaded)
                    if err:
                        st.error(f"Error: {err}")
                    else:
                        st.success(f"✅ {count} {T('import_ok')}")
                        st.rerun()
            except Exception as e:
                st.error(f"No se pudo leer el archivo: {e}")

    with t5:
        with db() as con:
            udf=pd.read_sql("SELECT id,username,role,vecino_id FROM users",con)
        st.dataframe(udf,use_container_width=True,hide_index=True)
        st.markdown("---")
        st.subheader(T("add_user"))
        with st.form("fadd_u"):
            nu=st.text_input("Username"); nr=st.selectbox("Role",["admin","vecino"])
            npw=st.text_input(T("new_pw"),type="password")
            with db() as con:
                vdf4=pd.read_sql("SELECT * FROM vecinos",con)
            nv=st.selectbox(T("neighbor"),["(ninguno)"]+vdf4["nombre"].tolist())
            if st.form_submit_button(T("add")) and nu and npw:
                vid_new=None
                if nv!="(ninguno)":
                    vid_new=int(vdf4[vdf4["nombre"]==nv]["id"].values[0])
                try:
                    with db() as con:
                        con.execute("INSERT INTO users(username,password,role,vecino_id) VALUES(?,?,?,?)",
                                    (nu,hash_pw(npw),nr,vid_new))
                    st.success(T("saved_ok")); st.rerun()
                except Exception as ex:
                    st.error(f"Error: {ex}")
        st.markdown("---")
        st.subheader(T("edit_user"))
        with st.form("fedit_u"):
            usel=st.selectbox(T("users"),udf["username"].tolist(),key="edit_usel")
            new_uname=st.text_input("Nuevo username (vacío = no cambiar)")
            new_role=st.selectbox("Role",["admin","vecino"])
            new_pw=st.text_input(T("new_pw"),type="password")
            if st.form_submit_button(T("save")):
                target=new_uname or usel
                with db() as con:
                    if new_uname:
                        con.execute("UPDATE users SET username=? WHERE username=?",(new_uname,usel))
                    if new_pw:
                        con.execute("UPDATE users SET password=? WHERE username=?",(hash_pw(new_pw),target))
                    con.execute("UPDATE users SET role=? WHERE username=?",(new_role,target))
                st.success(T("saved_ok")); st.rerun()
        st.markdown("---")
        st.subheader(T("del_user"))
        usel_del=st.selectbox(T("users"),udf[udf["username"]!="admin"]["username"].tolist(),key="del_usel")
        chk_u=st.checkbox(T("confirm_delete"),key="chk_del_u")
        if st.button("🗑️ "+T("del_user"),disabled=not chk_u):
            with db() as con:
                con.execute("DELETE FROM users WHERE username=?",(usel_del,))
            st.success(T("saved_ok")); st.rerun()

# ══ ADMIN: FACTURACIÓN ════════════════════════════════════════════════════════
elif role=="admin" and menu==T("menu_fact"):
    st.title(T("menu_fact"))
    tar=get_tar()
    with db() as con:
        vdf=pd.read_sql("SELECT * FROM vecinos",con)
        anyos=pd.read_sql("SELECT DISTINCT anyo FROM consumos ORDER BY anyo DESC",con)["anyo"].tolist()

    c1,c2=st.columns(2)
    af=c1.selectbox(T("year"),anyos or [date.today().year])
    mf=c2.selectbox(T("month"),list(range(1,13)),format_func=lambda x:months()[x-1],
                    index=date.today().month-1)
    MES=months()

    with db() as con:
        df_f=pd.read_sql(
            """SELECT v.id,v.nombre,v.email,v.tiene_email,v.direccion,v.telefono,c.m3
               FROM vecinos v LEFT JOIN consumos c ON c.vecino_id=v.id AND c.anyo=? AND c.mes=?""",
            con,params=(af,mf))

    vdf_f=buscar_vecino(vdf,"fact")
    if not vdf_f.empty and len(vdf_f)<len(vdf):
        df_f=df_f[df_f["nombre"].isin(vdf_f["nombre"])]

    # ── Calcular totales sin generar PDFs ────────────────────────────────────
    pm3=float(tar.get("precio_m3",0.85))
    cf_t=float(tar.get("cuota_fija",5.0))
    iva_t=float(tar.get("iva",0.10))

    def calc_total(m3v):
        base=m3v*pm3+cf_t
        return round(base*(1+iva_t),2)

    df_f=df_f.copy()
    df_f["m3v"]=df_f["m3"].apply(lambda x: float(x) if x else 0.0)
    df_f["total_eur"]=df_f["m3v"].apply(calc_total)
    df_f["contacto"]=df_f["tiene_email"].apply(lambda x:"📧 Email" if x else "✉️ Postal")
    total_mes=df_f["total_eur"].sum()

    st.subheader(f"{T('invoice_period')} {MES[mf-1]} {af}")

    # Tabla resumen — sin generar ningún PDF
    df_show=df_f[["nombre","m3v","total_eur","contacto"]].rename(columns={
        "nombre":T("name"),"m3v":"m³","total_eur":"Total €","contacto":"Contacto"})
    st.dataframe(df_show.reset_index(drop=True),use_container_width=True,hide_index=True)
    st.success(f"**{T('total_billed')}: {total_mes:.2f} €**")

    # ── Descarga individual de PDF (bajo demanda) ─────────────────────────────
    st.markdown("---")
    st.subheader(f"📄 {T('send_single')}")
    nombres_fact=df_f["nombre"].tolist()
    sel_fact=st.selectbox(T("neighbor"), nombres_fact, key="sel_fact_dl")
    if sel_fact:
        row_dl=df_f[df_f["nombre"]==sel_fact].iloc[0]
        col_dl1,col_dl2=st.columns(2)
        if col_dl1.button("📄 Generar y descargar PDF",key="btn_gen_pdf"):
            buf_dl,nf_dl,total_dl,fname_dl=make_pdf(
                row_dl.to_dict(),mf,af,float(row_dl["m3v"]),
                pm3,cf_t,iva_t,
                st.session_state["lang"],
                tar.get("entidad",""),tar.get("direccion",""),tar.get("contacto",""),tar.get("iban",""))
            pdf_dl=buf_dl.read()
            save_factura_disk(pdf_dl,fname_dl,af,mf)
            col_dl2.download_button(T("download_pdf"),data=pdf_dl,
                file_name=fname_dl,mime="application/pdf",key="dl_pdf_individual")

    st.markdown("---")
    # ── EMAIL SIMPLE ──────────────────────────────────────────────────────────
    st.subheader(f"📧 {T('email_cfg')}")
    st.caption(T("email_hint"))
    ec1,ec2=st.columns(2)
    ef=ec1.text_input(T("email_from"),value=tar.get("email_from",""),
                      placeholder="ajuntament@leslloses.cat",key="ef")
    ep=ec2.text_input(T("email_pass"),type="password",key="ep",
                      value=tar.get("email_pass",""),
                      help="Gmail: Ajustes → Seguridad → Contraseñas de aplicación")

    # Configuración SMTP avanzada (opcional)
    with st.expander("⚙️ Configuración SMTP avanzada (solo si el envío falla)"):
        st.caption("Si tu servidor de correo da error de certificado (hosting propio, OVH, Hostalia...) "
                   "especifica aquí el servidor SMTP correcto.")
        sa1,sa2,sa3,sa4=st.columns([3,1,1,2])
        e_host=sa1.text_input("Servidor SMTP",value=tar.get("smtp_host",""),
                               placeholder="mail.tudominio.com",key="e_host")
        e_port=sa2.number_input("Puerto",value=int(tar.get("smtp_port",587) or 587),
                                 min_value=1,max_value=65535,key="e_port")
        e_ssl=sa3.checkbox("SSL",value=bool(tar.get("smtp_ssl",0)),key="e_ssl")
        e_noverify=sa4.checkbox("⚠️ No verificar certificado SSL",
                                value=bool(tar.get("smtp_no_verify",0)),key="e_noverify",
                                help="Activa esto si ves error 'CERTIFICATE_VERIFY_FAILED'. "
                                     "Necesario en muchos hostings compartidos.")
        st.caption("Ejemplos: Hostalia/Nominalia → mail.tudominio.com puerto 465 SSL ✓  |  "
                   "Gmail → dejar vacío  |  OVH → ssl0.ovh.net puerto 465 SSL ✓")

    col_save,col_send=st.columns(2)
    if col_save.button("💾 Guardar configuración email"):
        with db() as con:
            con.execute("UPDATE tarifas SET email_from=?,email_pass=?,smtp_host=?,smtp_port=?,smtp_ssl=?,smtp_no_verify=? WHERE id=1",
                        (ef,ep,e_host,int(e_port),int(e_ssl),int(e_noverify)))
        st.success(T("saved_ok"))

    if col_send.button(T("send_all")):
        if not ef or not ep:
            st.error("Introduce el email y la contraseña primero.")
        else:
            amb=[r for _,r in df_f.iterrows() if r["tiene_email"] and r["email"]]
            prog=st.progress(0); ok_n=0; errs=[]
            for i,row in enumerate(amb):
                m3v=float(row["m3"]) if row["m3"] else 0
                buf2,nf2,_,fname2=make_pdf(
                    row.to_dict(),mf,af,m3v,
                    float(tar.get("precio_m3",0.85)),float(tar.get("cuota_fija",5.0)),float(tar.get("iva",0.10)),
                    st.session_state["lang"],
                    tar.get("entidad",""),tar.get("direccion",""),tar.get("contacto",""),tar.get("iban",""))
                pdf_b=buf2.read()
                # Guardar en disco siempre
                save_factura_disk(pdf_b,fname2,af,mf)
                # Enviar email
                ok,err=send_email(row["email"],row["nombre"],nf2,mf,af,pdf_b,ef,ep,st.session_state["lang"],
                                e_host,int(e_port),bool(e_ssl),bool(e_noverify))
                if ok: ok_n+=1
                else: errs.append(f"{row['nombre']}: {err}")
                prog.progress((i+1)/max(len(amb),1))
            st.success(f"✅ {ok_n} {T('sent_ok')} · {T('facturas_saved')}")
            for e in errs: st.error(e)

    st.markdown("---")
    st.subheader(T("send_single"))
    c1s,c2s=st.columns(2)
    vsel_s=c1s.selectbox(T("neighbor"),vdf["nombre"].tolist(),key="vsel_s")
    mes_s=c2s.selectbox(T("month"),list(range(1,13)),format_func=lambda x:MES[x-1],key="mes_s",index=mf-1)
    with db() as con:
        df_ind=pd.read_sql(
            """SELECT v.id,v.nombre,v.email,v.tiene_email,v.direccion,v.telefono,c.m3
               FROM vecinos v LEFT JOIN consumos c ON c.vecino_id=v.id AND c.anyo=? AND c.mes=?
               WHERE v.nombre=?""",con,params=(af,mes_s,vsel_s))
    if not df_ind.empty:
        row_s=df_ind.iloc[0].to_dict()
        m3_s=float(row_s["m3"]) if row_s.get("m3") else 0
        total_s=calc_total(m3_s)
        st.info(f"{vsel_s} · {MES[mes_s-1]} {af} · {m3_s:.2f} m³ · {total_s:.2f} €")
        col_a,col_b=st.columns(2)
        # Generar PDF solo al pulsar el botón de descarga
        if col_a.button("📄 Generar PDF",key="btn_gen_single"):
            buf_s,nf_s,_,fname_s=make_pdf(
                row_s,mes_s,af,m3_s,pm3,cf_t,iva_t,
                st.session_state["lang"],
                tar.get("entidad",""),tar.get("direccion",""),tar.get("contacto",""),tar.get("iban",""))
            pdf_s=buf_s.read()
            save_factura_disk(pdf_s,fname_s,af,mes_s)
            st.session_state["pdf_single"]=pdf_s
            st.session_state["fname_single"]=fname_s
            st.session_state["nf_single"]=nf_s
            st.session_state["row_single"]=row_s
            st.session_state["mes_single"]=mes_s
        if st.session_state.get("pdf_single") and st.session_state.get("fname_single"):
            col_a.download_button(T("download_pdf"),
                data=st.session_state["pdf_single"],
                file_name=st.session_state["fname_single"],
                mime="application/pdf",key="d_single")
            if row_s.get("tiene_email") and row_s.get("email"):
                if col_b.button(T("send_btn"),key="btn_send_single"):
                    if not ef or not ep:
                        st.error("Configura el email primero.")
                    else:
                        tar_fresh=get_tar()
                        _ef=tar_fresh.get("email_from","") or ef
                        _ep=tar_fresh.get("email_pass","") or ep
                        _host=tar_fresh.get("smtp_host","") or e_host
                        _port=int(tar_fresh.get("smtp_port",587) or 587)
                        _ssl=bool(tar_fresh.get("smtp_ssl",0))
                        _noverify=bool(tar_fresh.get("smtp_no_verify",0))
                        rs=st.session_state["row_single"]
                        ok,err=send_email(rs["email"],rs["nombre"],
                            st.session_state["nf_single"],
                            st.session_state["mes_single"],af,
                            st.session_state["pdf_single"],
                            _ef,_ep,st.session_state["lang"],
                            _host,_port,_ssl,_noverify)
                        if ok: st.success(f"✅ {T('sent_saved')}")
                        else: st.error(err)
            else:
                col_b.warning(T("no_email"))


    st.markdown("---")
    # ── FACTURAS GUARDADAS ────────────────────────────────────────────────────
    with st.expander(T("view_saved")):
        facturas=list_facturas()
        if not facturas:
            st.info(T("no_facturas"))
        else:
            df_fact=pd.DataFrame(facturas)
            st.dataframe(df_fact[["anyo","mes","archivo","size"]],
                         use_container_width=True,hide_index=True)
            st.caption(f"{len(df_fact)} facturas guardadas en {FACTURAS_DIR}")
            # Descarga individual
            sel_f=st.selectbox("Descargar factura",df_fact["archivo"].tolist(),key="sel_fact")
            row_f=df_fact[df_fact["archivo"]==sel_f].iloc[0]
            with open(row_f["ruta"],"rb") as f:
                st.download_button("⬇️ Descargar",data=f.read(),
                    file_name=sel_f,mime="application/pdf",key="dl_fact")

# ══ ADMIN: CONFIGURACIÓN ═════════════════════════════════════════════════════
elif role=="admin" and menu==T("menu_cfg"):
    st.title(T("menu_cfg"))
    tar=get_tar()

    st.subheader("💶 "+T("tariff"))
    with st.form("ftf"):
        c1,c2,c3=st.columns(3)
        p=c1.number_input(T("tariff"),value=float(tar.get("precio_m3",0.85)),step=0.0001,format="%.4f")
        q=c2.number_input(T("fixed_fee"),value=float(tar.get("cuota_fija",5.0)),step=0.5,format="%.2f")
        iv=c3.number_input(T("vat"),value=float(tar.get("iva",0.10))*100,step=1.0,format="%.0f")
        st.markdown("**Datos de la entidad (aparecen en las facturas)**")
        ent=st.text_input("Nombre entidad",value=str(tar.get("entidad","") or ""))
        dir_e=st.text_input("Dirección",value=str(tar.get("direccion","") or ""))
        cont=st.text_input("Contacto / teléfono",value=str(tar.get("contacto","") or ""))
        iban_e=st.text_input(T("iban"),value=str(tar.get("iban","") or ""),placeholder=T("iban_hint"))
        if st.form_submit_button(T("save")):
            with db() as con:
                con.execute(
                    "UPDATE tarifas SET precio_m3=?,cuota_fija=?,iva=?,entidad=?,direccion=?,contacto=?,iban=? WHERE id=1",
                    (p,q,iv/100,ent,dir_e,cont,iban_e))
            st.success(T("saved_ok"))

    st.markdown("---")
    st.subheader("📅 "+T("new_year"))
    with st.form("fa"):
        na=st.number_input(T("year"),min_value=2020,max_value=2040,value=date.today().year+1)
        if st.form_submit_button(T("save")):
            with db() as con:
                vids=pd.read_sql("SELECT id FROM vecinos",con)
                for _,v in vids.iterrows():
                    for mm in range(1,13):
                        ex=con.execute("SELECT id FROM consumos WHERE vecino_id=? AND anyo=? AND mes=?",
                                       (v["id"],int(na),mm)).fetchone()
                        if not ex:
                            con.execute("INSERT INTO consumos(vecino_id,anyo,mes,m3,fuente) VALUES(?,?,?,?,?)",
                                        (v["id"],int(na),mm,0,"manual"))
            st.success(f"✅ {na} {T('saved_ok')}")

    st.markdown("---")
    st.subheader(T("sensors_tab"))

    # ── Configuración MQTT ───────────────────────────────────────────────────
    with st.expander(T("mqtt_cfg")):
        tar2=get_tar()
        with st.form("fmqtt"):
            mc1,mc2=st.columns(2)
            m_broker=mc1.text_input(T("mqtt_broker"),value=str(tar2.get("mqtt_broker","localhost") or "localhost"))
            m_port=mc2.number_input(T("mqtt_port"),value=int(tar2.get("mqtt_port",1883) or 1883),min_value=1,max_value=65535)
            m_topic=st.text_input(T("mqtt_topic"),value=str(tar2.get("mqtt_topic","leslloses/#") or "leslloses/#"))
            m_activo=st.checkbox(T("mqtt_active"),value=bool(tar2.get("mqtt_activo",0)))
            if st.form_submit_button(T("save")):
                with db() as con:
                    con.execute("UPDATE tarifas SET mqtt_broker=?,mqtt_port=?,mqtt_topic=?,mqtt_activo=? WHERE id=1",
                                (m_broker,int(m_port),m_topic,1 if m_activo else 0))
                st.session_state.pop("mqtt_thread_started",None)
                st.success(T("saved_ok"))
                st.rerun()

    # ── Tabla de sensores ────────────────────────────────────────────────────
    with db() as con:
        df_sen=pd.read_sql("""
            SELECT s.id, s.sensor_id, v.nombre as vecino, s.pulsos_m3,
                   s.activo, s.ultima_lectura, s.ultimo_pulso
            FROM sensores s
            LEFT JOIN vecinos v ON v.id=s.vecino_id
            ORDER BY s.sensor_id""", con)
        vecinos_list=pd.read_sql("SELECT id,nombre FROM vecinos ORDER BY nombre",con)["nombre"].tolist()
        vecinos_ids=pd.read_sql("SELECT id,nombre FROM vecinos ORDER BY nombre",con)

    if not df_sen.empty:
        # Calcular estado
        def estado(row):
            if not row["ultima_lectura"]:
                return T("never")
            try:
                dt=datetime.strptime(row["ultima_lectura"],"%Y-%m-%d %H:%M")
                dias=(datetime.now()-dt).days
                if dias>3:
                    return f"{T('no_signal')} ({dias}d)"
                return f"{T('online')} ({row['ultima_lectura']})"
            except:
                return row["ultima_lectura"]

        df_sen["Estado"]=df_sen.apply(estado,axis=1)
        df_show=df_sen[["sensor_id","vecino","pulsos_m3","Estado","ultimo_pulso"]].rename(columns={
            "sensor_id":T("sensor_id"),"vecino":T("neighbor"),
            "pulsos_m3":T("pulsos_m3"),"ultimo_pulso":"Últimos pulsos"})
        st.dataframe(df_show,use_container_width=True,hide_index=True)

        # Eliminar sensor
        del_sen=st.selectbox("Eliminar sensor",["—"]+df_sen["sensor_id"].tolist(),key="del_sen")
        if del_sen!="—":
            if st.button("🗑️ Eliminar",key="btn_del_sen"):
                with db() as con:
                    con.execute("DELETE FROM sensores WHERE sensor_id=?",(del_sen,))
                st.rerun()
    else:
        st.info("No hay sensores configurados todavía.")

    # ── Añadir sensor ────────────────────────────────────────────────────────
    st.markdown(f"**➕ {T('add_sensor')}**")
    with st.form("fadd_sen"):
        col_s1,col_s2,col_s3=st.columns(3)
        new_sid=col_s1.text_input(T("sensor_id"),placeholder="ej. A8B3C2D1")
        new_vec=col_s2.selectbox(T("neighbor"),vecinos_list)
        new_pul=col_s3.number_input(T("pulsos_m3"),value=100,min_value=1,max_value=10000)
        if st.form_submit_button(T("add_sensor"),type="primary"):
            if new_sid and new_vec:
                vid_row=vecinos_ids[vecinos_ids["nombre"]==new_vec]
                if not vid_row.empty:
                    vid=int(vid_row.iloc[0]["id"])
                    try:
                        with db() as con:
                            con.execute(
                                "INSERT INTO sensores(sensor_id,vecino_id,pulsos_m3) VALUES(?,?,?)",
                                (new_sid.strip(),vid,int(new_pul)))
                        st.success(T("sensor_ok"))
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error: {e}")
            else:
                st.warning("Rellena el ID del sensor y selecciona un vecino.")

    st.markdown("---")
    st.subheader("💾 "+T("backup_db"))
    try:
        with open(DB,"rb") as f:
            db_bytes=f.read()
        st.download_button(T("backup_db"),data=db_bytes,
            file_name=f"lloses_{date.today().isoformat()}.db",
            mime="application/octet-stream")
    except Exception as e:
        st.error(str(e))

# ══ VECINO: MIS CONSUMOS ══════════════════════════════════════════════════════
elif role=="vecino" and menu==T("menu_mycons"):
    vid=st.session_state["vecino_id"]
    with db() as con:
        vinfo=con.execute("SELECT nombre FROM vecinos WHERE id=?",(vid,)).fetchone()
        anyos=pd.read_sql("SELECT DISTINCT anyo FROM consumos WHERE vecino_id=? ORDER BY anyo DESC",
                          con,params=(vid,))["anyo"].tolist()
        tar=dict(con.execute("SELECT * FROM tarifas WHERE id=1").fetchone())
    vname=vinfo[0] if vinfo else ""
    st.title(f"👋 {vname}")
    st.subheader(T("menu_mycons"))
    anyo=st.selectbox(T("year"),anyos or [date.today().year])
    with db() as con:
        df_v=pd.read_sql("SELECT mes,m3 FROM consumos WHERE vecino_id=? AND anyo=? ORDER BY mes",
                         con,params=(vid,anyo))
    MES=months()
    if df_v.empty:
        st.info(T("no_data"))
    else:
        pm3=float(tar["precio_m3"]); cf=float(tar["cuota_fija"]); iva=float(tar["iva"])
        df_v["coste"]=df_v["m3"].apply(lambda x:round((x*pm3+cf)*(1+iva),2))
        media=df_v["m3"].mean()
        coste_medio=round((media*pm3+cf)*(1+iva),2)
        k1,k2,k3,k4=st.columns(4)
        k1.metric(f"Total {anyo}",f"{df_v['m3'].sum():.1f} m³")
        k2.metric(T("avg_monthly"),f"{media:.1f} m³/mes")
        k3.metric("Coste medio/mes",f"{coste_medio:.2f} €/mes")
        k4.metric(f"Total {anyo} €",f"{df_v['coste'].sum():.2f} €")
        st.subheader(T("annual_evolution"))
        st.plotly_chart(chart_line(df_v.copy(),vname,st.session_state["lang"]),use_container_width=True)
        df_v["Mes"]=df_v["mes"].apply(lambda x:MES[x-1])
        st.dataframe(df_v[["Mes","m3","coste"]].rename(columns={"m3":"m³","coste":"Coste €"}),
                     use_container_width=True,hide_index=True)
        ex1,ex2=st.columns(2)
        csv_v=df_v[["Mes","m3","coste"]].rename(columns={"m3":"m³","coste":"Coste €"}).to_csv(index=False).encode()
        ex1.download_button(T("export_consumption"),data=csv_v,
            file_name=f"mis_consumos_{anyo}.csv",mime="text/csv")
        try:
            import plotly.io as pio
            fig_v=chart_line(df_v.copy(),vname,st.session_state["lang"])
            img_v=pio.to_image(fig_v,format="png",width=900,height=350,scale=2)
            ex2.download_button(T("export_chart"),data=img_v,
                file_name=f"mi_grafica_{anyo}.png",mime="image/png")
        except: pass

# ══ VECINO: MIS FACTURAS ══════════════════════════════════════════════════════
elif role=="vecino" and menu==T("menu_myfact"):
    vid=st.session_state["vecino_id"]
    with db() as con:
        vinfo=con.execute("SELECT * FROM vecinos WHERE id=?",(vid,)).fetchone()
        tar=dict(con.execute("SELECT * FROM tarifas WHERE id=1").fetchone())
        anyos=pd.read_sql("SELECT DISTINCT anyo FROM consumos WHERE vecino_id=? ORDER BY anyo DESC",
                          con,params=(vid,))["anyo"].tolist()
    vrow={"id":vinfo[0],"nombre":vinfo[1],"direccion":vinfo[2],
          "email":vinfo[3],"telefono":vinfo[5] if len(vinfo)>5 else ""}
    st.title(T("menu_myfact"))
    anyo_f=st.selectbox(T("year"),anyos or [date.today().year])
    with db() as con:
        cons=pd.read_sql("SELECT mes,m3 FROM consumos WHERE vecino_id=? AND anyo=? ORDER BY mes",
                         con,params=(vid,anyo_f))
    MES=months()
    # Consumo medio del año
    if not cons.empty:
        media_f=cons["m3"].mean()
        coste_m=round((media_f*float(tar.get("precio_m3",0.85))+float(tar.get("cuota_fija",5.0)))*
                      (1+float(tar.get("iva",0.10))),2)
        k1,k2=st.columns(2)
        k1.metric(T("avg_monthly"),f"{media_f:.1f} m³/mes")
        k2.metric("Coste medio/mes",f"{coste_m:.2f} €/mes")
        st.markdown("---")
    for _,crow in cons.iterrows():
        m3v=float(crow["m3"]) if crow["m3"] else 0
        buf,nf,total,fname=make_pdf(
            vrow,int(crow["mes"]),anyo_f,m3v,
            float(tar.get("precio_m3",0.85)),float(tar.get("cuota_fija",5.0)),float(tar.get("iva",0.10)),
            st.session_state["lang"],
            tar.get("entidad",""),tar.get("direccion",""),tar.get("contacto",""),tar.get("iban",""))
        with st.expander(f"{MES[int(crow['mes'])-1]} {anyo_f} – {m3v:.2f} m³ – {total:.2f} €"):
            st.download_button(T("download_pdf"),data=buf.read(),
                               file_name=fname,mime="application/pdf",key=f"vf{crow['mes']}")
            st.metric(T("total"),f"{total:.2f} €")
