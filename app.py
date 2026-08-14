from flask import Flask, render_template, request, jsonify, send_file, session
import xmlrpc.client
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from datetime import datetime, date
import io, os, json, re, base64
import pandas as pd
import psycopg2
from psycopg2.extras import RealDictCursor

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "vauma-maquinaria-2026-secreto")

class DateEncoder(json.JSONEncoder):
    def default(self, obj):
        if isinstance(obj, (datetime, date)):
            return obj.isoformat()
        return super().default(obj)

app.json_encoder = DateEncoder

# VAUMA no usa Odoo — gastos manuales únicamente
ODOO_URL  = ""
ODOO_DB   = ""
ODOO_USER = ""
ODOO_PASS = ""
# Múltiples administradores VAUMA
ADMINS = {
    os.environ.get("ADMIN_USER_1", "Luis.Alfaro"):  os.environ.get("ADMIN_PASS_1", "Alfaro2026"),
    os.environ.get("ADMIN_USER_2", "Cristhian.Lobo"): os.environ.get("ADMIN_PASS_2", "Lobo2026"),
}
# Compatibilidad legacy
ADMIN_USER = os.environ.get("ADMIN_USER_1", "Luis.Alfaro")
ADMIN_PASS = os.environ.get("ADMIN_PASS_1", "Alfaro2026")

# ── VAUMA — Vargas Ulloa Maquinaria S.A. ──────────────
# Colores: Gris #3A3A3A | Amarillo #FFE500 | Verde azulado #1B4B5A

# ── Proyectos base (siempre disponibles) ──────────────
PROYECTOS_BASE = {}

# ── Base de datos PostgreSQL ─────────────────────────
DATABASE_URL = os.environ.get("DATABASE_URL", "")

# ── SendGrid (notificaciones por email) ──────────────────
SENDGRID_API_KEY  = os.environ.get("SENDGRID_API_KEY", "")
SENDGRID_FROM     = os.environ.get("SENDGRID_FROM", "notificaciones@vauma.cr")
SENDGRID_FROM_NAME = os.environ.get("SENDGRID_FROM_NAME", "Vargas Ulloa Maquinaria S.A.")

def get_db():
    """Retorna conexión a PostgreSQL."""
    return psycopg2.connect(DATABASE_URL)

def init_db():
    """Crea tablas si no existen."""
    if not DATABASE_URL:
        print("[DB] Sin DATABASE_URL — modo solo memoria")
        return
    try:
        conn = get_db()
        cur = conn.cursor()
        cur.execute("""
            CREATE TABLE IF NOT EXISTS proyectos (
                codigo VARCHAR(50) PRIMARY KEY,
                nombre TEXT NOT NULL,
                cliente TEXT NOT NULL,
                partidas JSONB NOT NULL,
                email_cliente TEXT DEFAULT '',
                telefono_cliente TEXT DEFAULT '',
                creado_en TIMESTAMP DEFAULT NOW()
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS cronograma (
                id SERIAL PRIMARY KEY,
                proyecto_codigo VARCHAR(50) NOT NULL REFERENCES proyectos(codigo) ON DELETE CASCADE,
                capitulo VARCHAR(200) NOT NULL,
                orden INTEGER DEFAULT 0,
                fecha_inicio_plan DATE,
                fecha_fin_plan DATE,
                fecha_inicio_real DATE,
                fecha_fin_real DATE,
                duracion_semanas NUMERIC(5,1) DEFAULT 0,
                pct_avance INTEGER DEFAULT 0,
                estado VARCHAR(20) DEFAULT 'pendiente',
                creado_en TIMESTAMP DEFAULT NOW()
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS bitacora (
                id SERIAL PRIMARY KEY,
                proyecto_codigo VARCHAR(50) NOT NULL,
                fecha DATE NOT NULL DEFAULT CURRENT_DATE,
                autor VARCHAR(100) DEFAULT 'Admin',
                titulo VARCHAR(200),
                contenido TEXT NOT NULL,
                capitulo VARCHAR(200),
                creado_en TIMESTAMP DEFAULT NOW()
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS fotos_obra (
                id SERIAL PRIMARY KEY,
                proyecto_codigo VARCHAR(50) NOT NULL,
                fecha DATE NOT NULL DEFAULT CURRENT_DATE,
                capitulo VARCHAR(200),
                autor VARCHAR(100) DEFAULT 'Admin',
                nota TEXT,
                imagen_b64 TEXT,
                pct_estimado INTEGER,
                pct_confirmado INTEGER,
                aprobado BOOLEAN DEFAULT FALSE,
                creado_en TIMESTAMP DEFAULT NOW()
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS gastos_manuales (
                id SERIAL PRIMARY KEY,
                proyecto_codigo VARCHAR(50) NOT NULL,
                fecha DATE NOT NULL DEFAULT CURRENT_DATE,
                partida_codigo VARCHAR(50) NOT NULL,
                partida_nombre VARCHAR(200),
                descripcion TEXT NOT NULL,
                monto NUMERIC(15,2) NOT NULL,
                tipo VARCHAR(50) DEFAULT 'efectivo',
                comprobante_b64 TEXT,
                comprobante_tipo VARCHAR(10),
                creado_en TIMESTAMP DEFAULT NOW()
            )
        """)
        conn.commit()
        cur.close()
        conn.close()
        print("[DB] Conectado a PostgreSQL OK")
    except Exception as e:
        print(f"[DB] Error init: {e}")

def get_proyectos_db():
    """Lee proyectos dinámicos desde PostgreSQL."""
    if not DATABASE_URL:
        return {}
    try:
        conn = get_db()
        cur = conn.cursor(cursor_factory=RealDictCursor)
        cur.execute("SELECT * FROM proyectos ORDER BY creado_en")
        rows = cur.fetchall()
        cur.close(); conn.close()
        result = {}
        for row in rows:
            partidas_raw = row['partidas']
            if isinstance(partidas_raw, str):
                partidas_raw = json.loads(partidas_raw)
            result[row['codigo']] = {
                'nombre':   row['nombre'],
                'cliente':  row['cliente'],
                'partidas': [tuple(p) for p in partidas_raw],
            }
        return result
    except Exception as e:
        print(f"[DB] Error read: {e}")
        return {}

def save_proyecto_db(codigo, nombre, cliente, partidas, email_cliente="", telefono_cliente=""):
    """Guarda un proyecto en PostgreSQL."""
    if not DATABASE_URL:
        return False
    try:
        conn = get_db()
        cur = conn.cursor()
        cur.execute("""
            INSERT INTO proyectos (codigo, nombre, cliente, partidas, email_cliente, telefono_cliente)
            VALUES (%s, %s, %s, %s, %s, %s)
            ON CONFLICT (codigo) DO UPDATE
            SET nombre=EXCLUDED.nombre, cliente=EXCLUDED.cliente, partidas=EXCLUDED.partidas,
                email_cliente=EXCLUDED.email_cliente, telefono_cliente=EXCLUDED.telefono_cliente
        """, (codigo, nombre, cliente, json.dumps(partidas), email_cliente, telefono_cliente))
        conn.commit()
        cur.close(); conn.close()
        return True
    except Exception as e:
        print(f"[DB] Error save: {e}")
        return False

def delete_proyecto_db(codigo):
    """Elimina un proyecto de PostgreSQL."""
    if not DATABASE_URL:
        return False
    try:
        conn = get_db()
        cur = conn.cursor()
        cur.execute("DELETE FROM proyectos WHERE codigo = %s", (codigo,))
        conn.commit()
        cur.close(); conn.close()
        return True
    except Exception as e:
        print(f"[DB] Error delete: {e}")
        return False

# Inicializar DB al arrancar
init_db()

def get_proyectos():
    """Retorna todos los proyectos: base + dinámicos desde DB."""
    return {**PROYECTOS_BASE, **get_proyectos_db()}

def get_email_cliente(codigo_proyecto):
    """Obtiene el email del cliente desde la base de datos."""
    try:
        conn = get_db()
        cur = conn.cursor()
        cur.execute("SELECT email_cliente, cliente FROM proyectos WHERE codigo = %s", (codigo_proyecto,))
        row = cur.fetchone()
        cur.close(); conn.close()
        if row:
            return row[0], row[1]
        return None, None
    except:
        return None, None

def enviar_notificacion_email(email_destino, nombre_cliente, nombre_proyecto, tipo, detalle=""):
    """
    Envía notificación por email via SendGrid.
    tipo: 'avance', 'foto', 'bitacora'
    """
    if not SENDGRID_API_KEY or not email_destino:
        print(f"[EMAIL] Sin API key o email — notificación omitida")
        return False

    temas = {
        "avance":   f"📅 Nuevo avance registrado en su proyecto",
        "foto":     f"📷 Nueva foto de obra disponible",
        "bitacora": f"📝 Nueva entrada en bitácora de obra",
    }
    asunto = temas.get(tipo, "Actualización de su proyecto") + f" — {nombre_proyecto}"

    html_body = f"""
<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"></head>
<body style="margin:0;padding:0;background:#f4f4f4;font-family:Arial,sans-serif">
  <table width="100%" cellpadding="0" cellspacing="0" style="background:#f4f4f4;padding:30px 0">
    <tr><td align="center">
      <table width="580" cellpadding="0" cellspacing="0" style="background:#ffffff;border-radius:12px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,0.1)">
        <!-- Header -->
        <tr><td style="background:#080e0a;padding:28px 32px;text-align:center">
          <div style="font-size:22px;font-weight:700;color:#E8A020;letter-spacing:2px">URBANISTYKA</div>
          <div style="font-size:11px;color:#94a896;letter-spacing:3px;margin-top:4px">CONSTRUCTORA</div>
        </td></tr>
        <!-- Body -->
        <tr><td style="padding:32px">
          <p style="font-size:15px;color:#333;margin:0 0 16px">Estimado/a <strong>{nombre_cliente}</strong>,</p>
          <p style="font-size:14px;color:#555;margin:0 0 20px">Le informamos que hay una nueva actualización en su proyecto:</p>
          <div style="background:#f8f9fa;border-left:4px solid #E8A020;border-radius:4px;padding:16px 20px;margin:0 0 24px">
            <div style="font-size:12px;color:#888;text-transform:uppercase;letter-spacing:1px;margin-bottom:6px">Proyecto</div>
            <div style="font-size:16px;font-weight:700;color:#1a1a1a">{nombre_proyecto}</div>
            {f'<div style="font-size:14px;color:#555;margin-top:8px">{detalle}</div>' if detalle else ''}
          </div>
          <p style="font-size:14px;color:#555;margin:0 0 24px">Puede revisar el estado completo de su proyecto ingresando al portal:</p>
          <div style="text-align:center;margin:0 0 24px">
            <a href="https://urbanistykaconstruction.up.railway.app" style="background:#E8A020;color:#000;text-decoration:none;padding:12px 28px;border-radius:6px;font-weight:700;font-size:14px;display:inline-block">
              Ver mi proyecto →
            </a>
          </div>
          <p style="font-size:12px;color:#999;margin:0">Si tiene alguna consulta, no dude en contactarnos.</p>
        </td></tr>
        <!-- Footer -->
        <tr><td style="background:#f0f0f0;padding:16px 32px;text-align:center">
          <p style="font-size:11px;color:#999;margin:0">© 2026 Urbanistyka Constructora · Costa Rica</p>
          <p style="font-size:11px;color:#999;margin:4px 0 0">Este es un mensaje automático del sistema de control de obra.</p>
        </td></tr>
      </table>
    </td></tr>
  </table>
</body>
</html>"""

    try:
        payload = json_lib.dumps({
            "personalizations": [{"to": [{"email": email_destino, "name": nombre_cliente}]}],
            "from": {"email": SENDGRID_FROM, "name": SENDGRID_FROM_NAME},
            "subject": asunto,
            "content": [{"type": "text/html", "value": html_body}]
        }).encode("utf-8")

        req = urllib.request.Request(
            "https://api.sendgrid.com/v3/mail/send",
            data=payload,
            headers={
                "Content-Type": "application/json",
                "Authorization": f"Bearer {SENDGRID_API_KEY}"
            },
            method="POST"
        )
        with urllib.request.urlopen(req, timeout=15) as resp:
            status = resp.status
        print(f"[EMAIL] Enviado a {email_destino} — status {status}")
        return True
    except Exception as e:
        print(f"[EMAIL] Error enviando a {email_destino}: {e}")
        return False

def conectar_odoo():
    common = xmlrpc.client.ServerProxy(f"{ODOO_URL}/xmlrpc/2/common")
    uid = common.authenticate(ODOO_DB, ODOO_USER, ODOO_PASS, {})
    if not uid:
        raise Exception("Error de autenticación con Odoo")
    models = xmlrpc.client.ServerProxy(f"{ODOO_URL}/xmlrpc/2/object")
    return uid, models

def crear_cuentas_odoo(nombre_proyecto, partidas):
    """VAUMA no usa Odoo — no crea cuentas analíticas."""
    return 0

def obtener_datos_proyecto(nombre_proyecto, partidas):
    """VAUMA no usa Odoo — retorna gastos vacíos, los gastos manuales se suman aparte."""
    gastos  = {c: 0 for c, _, _ in partidas}
    detalle = {c: [] for c, _, _ in partidas}
    return gastos, detalle

def generar_excel(nombre_proyecto, cliente, partidas, gastos, detalle):
    wb = Workbook()
    thin = Side(style="thin", color="B8CCE4")
    ba = Border(left=thin, right=thin, top=thin, bottom=thin)
    def fill(h): return PatternFill("solid", start_color=h, fgColor=h)
    def ca():    return Alignment(horizontal="center", vertical="center", wrap_text=True)
    def la():    return Alignment(horizontal="left",   vertical="center", wrap_text=True)
    ws = wb.active; ws.title = "Resumen"
    for col, w in zip("ABCDEFG", [8,34,16,16,16,10,14]):
        ws.column_dimensions[col].width = w
    ws.merge_cells("A1:G1")
    ws["A1"] = f"CONTROL DE OBRA — {nombre_proyecto}"
    ws["A1"].font = Font(name="Arial", bold=True, size=13, color="FFFFFF")
    ws["A1"].fill = fill("1F3864"); ws["A1"].alignment = ca()
    ws.merge_cells("A2:G2")
    ws["A2"] = f"Cliente: {cliente}  |  {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    ws["A2"].font = Font(name="Arial", italic=True, size=9, color="595959")
    ws["A2"].fill = fill("D6E4F0"); ws["A2"].alignment = ca()
    for col, h in enumerate(["#","Partida","Presupuesto","Gastado","Saldo","% Ejec.","Estado"], 1):
        c = ws.cell(row=3, column=col, value=h)
        c.font = Font(name="Arial", bold=True, size=10, color="FFFFFF")
        c.fill = fill("1F3864"); c.alignment = ca(); c.border = ba
    total_p = total_g = 0
    for i, (codigo, nombre, presup) in enumerate(partidas):
        r = i+4; g = gastos.get(codigo,0); s = presup-g
        p = (g/presup*100) if presup>0 else 0
        bg = "FCE4D6" if p>=100 else ("FFF2CC" if p>=80 else ("E2EFDA" if p>0 else ("DEEAF1" if i%2==0 else "FFFFFF")))
        estado = "EXCEDIDO" if p>=100 else ("ALERTA" if p>=80 else ("EN CURSO" if p>0 else "SIN INICIO"))
        for col,(v,fmt,aln) in enumerate(zip([codigo,nombre,presup,g,s,p/100,estado],
            [None,None,"₡#,##0","₡#,##0","₡#,##0","0.0%",None],[ca(),la(),ca(),ca(),ca(),ca(),ca()]),1):
            c = ws.cell(row=r,column=col,value=v)
            c.font=Font(name="Arial",size=10); c.fill=fill(bg); c.alignment=aln; c.border=ba
            if fmt: c.number_format=fmt
        total_p+=presup; total_g+=g
    tr=len(partidas)+4; ws.merge_cells(f"A{tr}:B{tr}")
    ws.cell(row=tr,column=1,value="TOTAL").font=Font(name="Arial",bold=True,size=11,color="FFFFFF")
    ws.cell(row=tr,column=1).fill=fill("1F3864"); ws.cell(row=tr,column=1).alignment=ca()
    for col,v,fmt in [(3,total_p,"₡#,##0"),(4,total_g,"₡#,##0"),
                      (5,total_p-total_g,"₡#,##0"),(6,(total_g/total_p if total_p else 0),"0.0%")]:
        c=ws.cell(row=tr,column=col,value=v)
        c.font=Font(name="Arial",bold=True,size=11,color="FFFFFF")
        c.fill=fill("1F3864"); c.alignment=ca(); c.number_format=fmt
    output=io.BytesIO(); wb.save(output); output.seek(0)
    return output

def generar_excel_detallado_planos(nombre_proyecto, detalle):
    """
    Genera el Excel detallado de presupuesto (estilo Casa Itaca):
    una fila por actividad, con M.O. y Materiales separados, agrupado
    por capítulo, con totales generales al final.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Presupuesto"

    thin = Side(style="thin", color="B8CCE4")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    def fill(hex_color):
        return PatternFill("solid", start_color=hex_color, fgColor=hex_color)

    def center():
        return Alignment(horizontal="center", vertical="center", wrap_text=True)

    def left():
        return Alignment(horizontal="left", vertical="center", wrap_text=True)

    widths = {"A": 6, "B": 38, "C": 10, "D": 12, "E": 16, "F": 16, "G": 18, "H": 18, "I": 18}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    ws.merge_cells("A1:I1")
    ws["A1"] = f"PRESUPUESTO DETALLADO — {nombre_proyecto.upper()}"
    ws["A1"].font = Font(name="Arial", bold=True, size=14, color="FFFFFF")
    ws["A1"].fill = fill("1F3864")
    ws["A1"].alignment = center()

    ws.merge_cells("A2:I2")
    ws["A2"] = f"Generado con asistencia de IA a partir de planos · {datetime.now().strftime('%d/%m/%Y')}"
    ws["A2"].font = Font(name="Arial", italic=True, size=9, color="595959")
    ws["A2"].fill = fill("D6E4F0")
    ws["A2"].alignment = center()

    headers = ["#", "Partida / Actividad", "Unid.", "Cantidad", "P.U. M.O. (₡)",
               "P.U. Mat. (₡)", "Subtotal M.O. (₡)", "Subtotal Mat. (₡)", "Total Partida (₡)"]
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=3, column=col, value=h)
        c.font = Font(name="Arial", bold=True, size=10, color="FFFFFF")
        c.fill = fill("1F3864")
        c.alignment = center()
        c.border = border

    row = 4
    num = 1
    total_mo_general = 0
    total_mat_general = 0

    capitulos_orden = []
    for item in detalle:
        if item["capitulo"] not in capitulos_orden:
            capitulos_orden.append(item["capitulo"])

    for cap_nombre in capitulos_orden:
        ws.merge_cells(f"A{row}:I{row}")
        ws.cell(row=row, column=1, value=cap_nombre.upper())
        c = ws.cell(row=row, column=1)
        c.font = Font(name="Arial", bold=True, size=11, color="1F3864")
        c.fill = fill("D6E4F0")
        c.alignment = left()
        row += 1

        items_cap = [d for d in detalle if d["capitulo"] == cap_nombre]
        for item in items_cap:
            bg = "FFFFFF" if num % 2 == 0 else "F2F7FB"
            valores = [num, item["nombre"], item["unidad"], item["cantidad"],
                       item["pu_mo"], item["pu_mat"], item["sub_mo"], item["sub_mat"], item["total"]]
            formatos = [None, None, None, "#,##0.00", "₡#,##0", "₡#,##0", "₡#,##0", "₡#,##0", "₡#,##0"]
            alineaciones = [center(), left(), center(), center(), center(), center(), center(), center(), center()]

            for col, (val, fmt, aln) in enumerate(zip(valores, formatos, alineaciones), 1):
                c = ws.cell(row=row, column=col, value=val)
                c.font = Font(name="Arial", size=10)
                c.fill = fill(bg)
                c.alignment = aln
                c.border = border
                if fmt:
                    c.number_format = fmt

            total_mo_general += item["sub_mo"]
            total_mat_general += item["sub_mat"]
            row += 1
            num += 1

    ws.merge_cells(f"A{row}:F{row}")
    c = ws.cell(row=row, column=1, value="TOTAL COSTO DIRECTO")
    c.font = Font(name="Arial", bold=True, size=12, color="FFFFFF")
    c.fill = fill("1F3864")
    c.alignment = Alignment(horizontal="right", vertical="center")

    costo_directo = total_mo_general + total_mat_general
    for col, val in zip([7, 8, 9], [total_mo_general, total_mat_general, costo_directo]):
        c = ws.cell(row=row, column=col, value=val)
        c.font = Font(name="Arial", bold=True, size=12, color="FFFFFF")
        c.fill = fill("1F3864")
        c.alignment = center()
        c.number_format = "₡#,##0"
    row += 1

    monto_admin = round(costo_directo * 0.08)
    monto_util = round(costo_directo * 0.07)
    filas_resumen = [
        ("Administración (8%)", monto_admin, "FFF2CC"),
        ("Utilidad (7%)", monto_util, "E2EFDA"),
    ]
    for nombre_fila, monto_fila, bg in filas_resumen:
        ws.merge_cells(f"A{row}:H{row}")
        c = ws.cell(row=row, column=1, value=nombre_fila)
        c.font = Font(name="Arial", bold=True, size=11)
        c.fill = fill(bg)
        c.alignment = Alignment(horizontal="right", vertical="center")
        c = ws.cell(row=row, column=9, value=monto_fila)
        c.font = Font(name="Arial", bold=True, size=11)
        c.fill = fill(bg)
        c.alignment = center()
        c.number_format = "₡#,##0"
        row += 1

    ws.merge_cells(f"A{row}:H{row}")
    c = ws.cell(row=row, column=1, value="PRECIO TOTAL (SIN IVA)")
    c.font = Font(name="Arial", bold=True, size=13, color="FFFFFF")
    c.fill = fill("E8A020")
    c.alignment = Alignment(horizontal="right", vertical="center")
    c = ws.cell(row=row, column=9, value=costo_directo + monto_admin + monto_util)
    c.font = Font(name="Arial", bold=True, size=13, color="FFFFFF")
    c.fill = fill("E8A020")
    c.alignment = center()
    c.number_format = "₡#,##0"

    ws.freeze_panes = "A4"

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

def leer_excel_presupuesto(file_bytes, extension):
    """Lee un archivo Excel y extrae partidas con montos."""
    try:
        if extension in ['xls']:
            df = pd.read_excel(io.BytesIO(file_bytes), engine='xlrd', header=None)
        else:
            df = pd.read_excel(io.BytesIO(file_bytes), header=None)

        partidas = []
        seen = set()
        for _, row in df.iterrows():
            for col_idx in range(len(row)-1):
                val = row.iloc[col_idx]
                if isinstance(val, str) and len(val.strip()) > 3:
                    nombre = val.strip()
                    for offset in range(1, min(6, len(row)-col_idx)):
                        monto = row.iloc[col_idx+offset]
                        if isinstance(monto, (int, float)) and not pd.isna(monto) and monto > 10000:
                            key = nombre.upper()[:30]
                            if key not in seen:
                                seen.add(key)
                                partidas.append({"nombre": nombre[:50], "monto": int(monto)})
                            break

        partidas = [p for p in partidas if 50000 <= p["monto"] <= 50000000]
        partidas.sort(key=lambda x: x["monto"], reverse=True)
        return partidas[:30]
    except Exception as e:
        return []

# ── RUTAS ──────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/login", methods=["POST"])
def login():
    data     = request.json
    usuario  = data.get("usuario","").strip()
    password = data.get("password","").strip()
    codigo   = data.get("codigo","").strip().upper()
    if usuario and password:
        if usuario in ADMINS and ADMINS[usuario] == password:
            session["tipo"] = "admin"; session["admin"] = True; session["admin_user"] = usuario
            return jsonify({"ok": True, "admin": True})
        return jsonify({"ok": False, "msg": "Usuario o contraseña incorrectos"}), 401
    proyectos = get_proyectos()
    if codigo in proyectos:
        session["tipo"] = "cliente"; session["codigo"] = codigo; session["admin"] = False
        return jsonify({"ok": True, "admin": False,
                        "nombre": proyectos[codigo]["nombre"],
                        "cliente": proyectos[codigo]["cliente"]})
    return jsonify({"ok": False, "msg": "Código inválido"}), 401

@app.route("/logout")
def logout():
    session.clear(); return jsonify({"ok": True})

@app.route("/api/proyectos")
def api_proyectos():
    if "tipo" not in session:
        return jsonify({"error": "No autorizado"}), 401
    proyectos = get_proyectos()
    codigos = list(proyectos.keys()) if session.get("admin") else [session.get("codigo")]
    resultado = []
    for cod in codigos:
        if not cod or cod not in proyectos: continue
        p = proyectos[cod]
        try:
            gastos, detalle = obtener_datos_proyecto(p["nombre"], p["partidas"])
        except:
            gastos  = {c: 0 for c,_,_ in p["partidas"]}
            detalle = {c: [] for c,_,_ in p["partidas"]}

        # Sumar gastos manuales por partida e incluir en detalle
        try:
            conn = get_db()
            cur = conn.cursor(cursor_factory=RealDictCursor)
            cur.execute("""
                SELECT partida_codigo, id, fecha, descripcion, monto, tipo, comprobante_tipo
                FROM gastos_manuales WHERE proyecto_codigo = %s
                ORDER BY fecha DESC
            """, (cod,))
            gastos_manuales_rows = cur.fetchall()
            cur.close(); conn.close()
            gastos_manuales = {}
            detalle_manuales = {}
            for r in gastos_manuales_rows:
                pc = r["partida_codigo"]
                gastos_manuales[pc] = gastos_manuales.get(pc, 0) + float(r["monto"])
                if pc not in detalle_manuales:
                    detalle_manuales[pc] = []
                detalle_manuales[pc].append({
                    "fecha": r["fecha"].isoformat() if hasattr(r["fecha"], "isoformat") else str(r["fecha"]),
                    "descripcion": r["descripcion"],
                    "proveedor": f"Manual ({r['tipo']})" + (" 📎" if r["comprobante_tipo"] else ""),
                    "monto": float(r["monto"]),
                    "es_manual": True,
                    "gasto_id": r["id"],
                    "comprobante_tipo": r["comprobante_tipo"]
                })
        except Exception as e:
            print(f"[GASTOS] Error leyendo manuales: {e}")
            gastos_manuales = {}
            detalle_manuales = {}

        partidas_data = []
        for codigo, nombre, presup in p["partidas"]:
            g_odoo = gastos.get(codigo, 0)
            g_manual = gastos_manuales.get(codigo, 0)
            g_total = g_odoo + g_manual
            pct = round(g_total/presup*100, 1) if presup > 0 else 0
            det_odoo = detalle.get(codigo, [])
            det_manual = detalle_manuales.get(codigo, [])
            partidas_data.append({"codigo":codigo,"nombre":nombre,"presupuesto":presup,
                "gastado":g_total,"gastado_odoo":g_odoo,"gastado_manual":g_manual,
                "saldo":presup-g_total,"pct":pct,"detalle":det_odoo + det_manual})
        total_p = sum(p2 for _,_,p2 in p["partidas"])
        total_g = sum(pt["gastado"] for pt in partidas_data)
        resultado.append({"codigo":cod,"nombre":p["nombre"],"cliente":p["cliente"],
            "total_presup":total_p,"total_gastado":total_g,
            "pct_global":round(total_g/total_p*100,1) if total_p>0 else 0,
            "partidas":partidas_data,"actualizado":datetime.now().strftime("%d/%m/%Y %H:%M")})
    return jsonify(resultado)

@app.route("/api/excel/<codigo_proyecto>")
def descargar_excel(codigo_proyecto):
    if "tipo" not in session: return "No autorizado", 401
    proyectos = get_proyectos()
    if not session.get("admin") and session.get("codigo") != codigo_proyecto:
        return "No autorizado", 401
    if codigo_proyecto not in proyectos: return "Proyecto no encontrado", 404
    p = proyectos[codigo_proyecto]
    try:
        gastos, detalle = obtener_datos_proyecto(p["nombre"], p["partidas"])
    except:
        gastos  = {c: 0 for c,_,_ in p["partidas"]}
        detalle = {c: [] for c,_,_ in p["partidas"]}
    excel = generar_excel(p["nombre"], p["cliente"], p["partidas"], gastos, detalle)
    return send_file(excel, as_attachment=True,
        download_name=f"Control_{codigo_proyecto}_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ── PANEL ADMIN: NUEVO PROYECTO ──────────────────────

@app.route("/api/admin/parsear-excel", methods=["POST"])
def parsear_excel():
    """Lee un Excel subido y retorna las partidas detectadas."""
    if not session.get("admin"):
        return jsonify({"error": "No autorizado"}), 401
    if "archivo" not in request.files:
        return jsonify({"error": "No se recibió archivo"}), 400
    archivo = request.files["archivo"]
    extension = archivo.filename.rsplit(".", 1)[-1].lower()
    file_bytes = archivo.read()
    partidas = leer_excel_presupuesto(file_bytes, extension)
    return jsonify({"ok": True, "partidas": partidas, "total": len(partidas)})

@app.route("/api/admin/crear-proyecto", methods=["POST"])
def crear_proyecto():
    """Crea un proyecto nuevo en DB y en Odoo."""
    if not session.get("admin"):
        return jsonify({"error": "No autorizado"}), 401
    data = request.json
    codigo  = data.get("codigo","").strip().upper()
    nombre  = data.get("nombre","").strip()
    cliente = data.get("cliente","").strip()
    partidas_raw = data.get("partidas", [])

    if not codigo or not nombre or not cliente or not partidas_raw:
        return jsonify({"ok": False, "msg": "Faltan datos obligatorios"}), 400

    proyectos = get_proyectos()
    if codigo in proyectos:
        return jsonify({"ok": False, "msg": f"El código '{codigo}' ya existe"}), 400

    prefijo = re.sub(r'[^A-Z0-9]', '', codigo[:3])
    partidas = []
    for i, p in enumerate(partidas_raw, 1):
        cod_partida = f"{prefijo}-{i:02d}"
        partidas.append((cod_partida, p["nombre"][:50], int(p["monto"])))

    # Guardar en PostgreSQL
    save_proyecto_db(codigo, nombre, cliente, partidas)

    # Crear en Odoo
    try:
        n_creadas = crear_cuentas_odoo(nombre, partidas)
        odoo_msg = f"{n_creadas} cuentas analíticas creadas en Odoo"
    except Exception as e:
        odoo_msg = f"Proyecto creado (Odoo: {str(e)[:50]})"

    return jsonify({"ok": True, "msg": odoo_msg, "codigo": codigo,
                    "partidas": len(partidas)})

@app.route("/api/admin/proyectos-lista")
def lista_proyectos():
    """Lista todos los proyectos para el panel admin."""
    if not session.get("admin"):
        return jsonify({"error": "No autorizado"}), 401
    proyectos = get_proyectos()
    result = []
    for cod, p in proyectos.items():
        total = sum(m for _,_,m in p["partidas"])
        # Un proyecto es dinámico si NO está en PROYECTOS_BASE
        es_dinamico = cod not in PROYECTOS_BASE
        result.append({"codigo": cod, "nombre": p["nombre"],
                       "cliente": p["cliente"], "total": total,
                       "partidas": len(p["partidas"]),
                       "dinamico": es_dinamico})
    return jsonify(result)

@app.route("/api/admin/actualizar-contacto", methods=["POST"])
def actualizar_contacto():
    """Actualiza email y teléfono del cliente de cualquier proyecto (base o dinámico)."""
    if not session.get("admin"):
        return jsonify({"error": "No autorizado"}), 401

    data = request.json
    codigo = data.get("codigo", "").upper()
    email = data.get("email_cliente", "").strip()
    telefono = data.get("telefono_cliente", "").strip()

    if not codigo:
        return jsonify({"ok": False, "msg": "Falta código"}), 400

    try:
        conn = get_db()
        cur = conn.cursor()
        # Upsert — funciona para proyectos base y dinámicos
        cur.execute("""
            INSERT INTO proyectos (codigo, nombre, cliente, partidas, email_cliente, telefono_cliente)
            VALUES (%s, %s, %s, %s::jsonb, %s, %s)
            ON CONFLICT (codigo) DO UPDATE
            SET email_cliente = EXCLUDED.email_cliente,
                telefono_cliente = EXCLUDED.telefono_cliente
        """, (codigo,
              get_proyectos().get(codigo, {}).get("nombre", codigo),
              get_proyectos().get(codigo, {}).get("cliente", ""),
              json.dumps(get_proyectos().get(codigo, {}).get("partidas", [])),
              email, telefono))
        conn.commit()
        cur.close(); conn.close()
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"ok": False, "msg": str(e)}), 500


@app.route("/api/admin/eliminar-proyecto/<codigo>", methods=["DELETE"])
def eliminar_proyecto(codigo):
    """Elimina un proyecto dinámico."""
    if not session.get("admin"):
        return jsonify({"error": "No autorizado"}), 401
    if codigo in PROYECTOS_BASE:
        return jsonify({"ok": False, "msg": "No se pueden eliminar proyectos base"}), 400
    if delete_proyecto_db(codigo):
        return jsonify({"ok": True})
    return jsonify({"ok": False, "msg": "Proyecto no encontrado"}), 404

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)

# ── ENDPOINT: PARSEO CON CLAUDE ───────────────────────
import urllib.request
import urllib.parse
import json as json_lib

@app.route("/api/admin/parsear-con-claude", methods=["POST"])
def parsear_con_claude():
    """Usa Claude API para interpretar un Excel de presupuesto."""
    if not session.get("admin"):
        return jsonify({"error": "No autorizado"}), 401
    if "archivo" not in request.files:
        return jsonify({"error": "No se recibió archivo"}), 400

    archivo = request.files["archivo"]
    extension = archivo.filename.rsplit(".", 1)[-1].lower()
    file_bytes = archivo.read()

    try:
        if extension == 'xls':
            df_dict = pd.read_excel(io.BytesIO(file_bytes), engine='xlrd',
                                    header=None, sheet_name=None)
        else:
            df_dict = pd.read_excel(io.BytesIO(file_bytes), header=None,
                                    sheet_name=None)

        texto_excel = ""
        for sheet_name, df in df_dict.items():
            texto_hoja = f"=== {sheet_name} ===\n"
            for _, row in df.iterrows():
                vals = [str(v).strip() for v in row if str(v).strip() not in ['nan','NaN','None','']]
                if not vals:
                    continue
                row_text = " | ".join(vals)
                tiene_numero_capitulo = any(
                    str(v).strip() in [str(i) for i in range(1, 30)] or
                    (str(v).strip().replace('.','').isdigit() and len(str(v).strip()) <= 2)
                    for v in row
                )
                tiene_monto_grande = any(
                    isinstance(v, (int, float)) and v > 100000
                    for v in row
                )
                texto_corto = len(row_text) < 80
                if (tiene_numero_capitulo or texto_corto) and tiene_monto_grande:
                    texto_hoja += row_text + "\n"
                elif tiene_monto_grande and len(vals) <= 4:
                    texto_hoja += row_text + "\n"
            texto_excel += texto_hoja
        texto_excel = texto_excel[:12000]
    except Exception as e:
        return jsonify({"ok": False, "msg": f"Error leyendo archivo: {str(e)}"}), 400

    prompt = f"""Eres un ingeniero de construcción costarricense experto en presupuestos de obra.

Analizá este presupuesto y extraé los CAPÍTULOS PRINCIPALES con sus totales para usarlos como partidas de control de obra.

ARCHIVO:
{texto_excel}

INSTRUCCIONES:
1. El presupuesto tiene una jerarquía: capítulos (1, 2, 3...) y sub-ítems (1.1, 1.2, 2.1...)
2. Incluí TODOS los capítulos principales numerados (1, 2, 3, 4... hasta el último)
3. Para cada capítulo usá su monto TOTAL (el que aparece junto al número del capítulo)
4. Si un capítulo no tiene total explícito, sumá sus sub-ítems directos
5. NO incluyas materiales individuales (varillas, cemento, bloques, cables, etc.)
6. NO incluyas sub-ítems si ya tenés el capítulo padre
7. Redondea los montos a números enteros
8. Incluí entre 8 y 15 partidas

Devolvé ÚNICAMENTE un JSON válido sin texto adicional ni backticks:
{{"partidas": [{{"nombre": "Nombre del capítulo", "monto": 1234567}}, ...]}}"""

    try:
        payload = json_lib.dumps({
            "model": "claude-sonnet-4-5",
            "max_tokens": 2000,
            "messages": [{"role": "user", "content": prompt}]
        }).encode('utf-8')

        ANTHROPIC_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
        if not ANTHROPIC_KEY:
            raise Exception("ANTHROPIC_API_KEY no configurada")
        req = urllib.request.Request(
            "https://api.anthropic.com/v1/messages",
            data=payload,
            headers={
                "Content-Type": "application/json",
                "x-api-key": ANTHROPIC_KEY,
                "anthropic-version": "2023-06-01"
            },
            method="POST"
        )
        with urllib.request.urlopen(req, timeout=30) as resp:
            result = json_lib.loads(resp.read().decode('utf-8'))

        text = result["content"][0]["text"].strip()
        text = text.replace("```json","").replace("```","").strip()
        parsed = json_lib.loads(text)
        partidas = parsed.get("partidas", [])
        partidas = [p for p in partidas
                    if p.get("nombre") and isinstance(p.get("monto"), (int,float)) and p["monto"] > 0]
        return jsonify({"ok": True, "partidas": partidas, "total": len(partidas)})

    except Exception as e:
        print(f"[CLAUDE] ERROR: {str(e)}")
        partidas = leer_excel_presupuesto(file_bytes, extension)
        return jsonify({"ok": True, "partidas": partidas, "total": len(partidas),
                        "fallback": True})

# ── ENDPOINT: GENERAR PRESUPUESTO DESDE PLANOS ───────────

@app.route("/api/admin/generar-desde-planos", methods=["POST"])
def generar_desde_planos():
    """
    Recibe hasta 10 imágenes/PDFs de planos, los envía a Claude API con visión,
    y devuelve un presupuesto detallado (estilo Casa Itaca: M.O. + Materiales
    separados por actividad) más el resumen agrupado por capítulo.
    """
    if not session.get("admin"):
        return jsonify({"error": "No autorizado"}), 401

    archivos = request.files.getlist("planos")
    if not archivos or len(archivos) == 0:
        return jsonify({"ok": False, "msg": "No se recibieron planos"}), 400
    if len(archivos) > 10:
        return jsonify({"ok": False, "msg": "Máximo 10 planos por análisis"}), 400

    nombre_proyecto = request.form.get("nombre", "Proyecto sin nombre")
    area_m2 = request.form.get("area_m2", "")

    content_blocks = []
    extensiones_validas = {"png", "jpg", "jpeg", "webp", "pdf"}

    for archivo in archivos:
        ext = archivo.filename.rsplit(".", 1)[-1].lower()
        if ext not in extensiones_validas:
            continue
        file_bytes = archivo.read()
        b64_data = base64.b64encode(file_bytes).decode("utf-8")

        if ext == "pdf":
            media_type = "application/pdf"
            block_type = "document"
        else:
            media_type = f"image/{'jpeg' if ext == 'jpg' else ext}"
            block_type = "image"

        content_blocks.append({
            "type": block_type,
            "source": {"type": "base64", "media_type": media_type, "data": b64_data}
        })

    if len(content_blocks) == 0:
        return jsonify({"ok": False, "msg": "Ningún archivo válido (use PNG, JPG, WEBP o PDF)"}), 400

    area_txt = f"\nÁREA APROXIMADA: {area_m2} m²" if area_m2 else ""

    prompt_texto = f"""Eres un ingeniero de costos costarricense experto en presupuestos de construcción residencial.

Analiza estos planos de construcción del proyecto "{nombre_proyecto}"{area_txt} y genera un presupuesto detallado de obra civil, igual al formato estándar de control de obra en Costa Rica.

INSTRUCCIONES:
1. Identifica TODAS las actividades constructivas visibles o inferibles de los planos: movimiento de tierras, fundaciones, estructura, mampostería, techos, acabados, instalaciones eléctricas, hidráulicas, sanitarias, pintura, etc.
2. Para cada actividad, estima: unidad de medida (m2, m3, m, kg, ml, unidad, global), cantidad, precio unitario de MANO DE OBRA (₡), precio unitario de MATERIALES (₡).
3. Usa precios de mercado costarricense actuales (Gran Área Metropolitana, 2026).
4. Agrupa las actividades en CAPÍTULOS (ej: "Preliminares", "Fundaciones", "Estructura", "Mampostería", "Techos", "Acabados", "Instalaciones Eléctricas", "Instalaciones Hidrosanitarias", "Pintura", etc.)
5. Genera entre 40 y 70 líneas de actividad en total, distribuidas en los capítulos.
6. Sé razonable y conservador con las cantidades — basate en lo que se puede inferir de los planos, y si algo no es claro, usa una estimación típica para una vivienda residencial de ese tamaño.
7. INCLUÍ OBLIGATORIAMENTE un capítulo llamado "Gastos Indirectos" con actividades propias (no como porcentaje), por ejemplo: alquiler de andamios, alquiler de mezcladora, alquiler de compactadora, herramienta menor, bodega temporal/caseta de obra, limpieza final de obra, transporte de materiales, energía y agua temporal de obra, rotulación y señalización, EPP (equipo de protección personal) de cuadrilla. Estimá cantidades y precios razonables para estas actividades según la duración típica de una obra de ese tamaño.

Devuelve ÚNICAMENTE un JSON válido sin texto adicional ni backticks, con esta estructura exacta:
{{
  "capitulos": [
    {{
      "nombre": "Nombre del capítulo",
      "actividades": [
        {{"nombre": "Nombre de la actividad", "unidad": "m2", "cantidad": 120, "pu_mo": 5000, "pu_mat": 8000}}
      ]
    }}
  ]
}}"""

    content_blocks.append({"type": "text", "text": prompt_texto})

    try:
        payload = json_lib.dumps({
            "model": "claude-sonnet-4-6",
            "max_tokens": 8000,
            "messages": [{"role": "user", "content": content_blocks}]
        }).encode("utf-8")

        ANTHROPIC_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
        if not ANTHROPIC_KEY:
            raise Exception("ANTHROPIC_API_KEY no configurada")

        req = urllib.request.Request(
            "https://api.anthropic.com/v1/messages",
            data=payload,
            headers={
                "Content-Type": "application/json",
                "x-api-key": ANTHROPIC_KEY,
                "anthropic-version": "2023-06-01"
            },
            method="POST"
        )
        with urllib.request.urlopen(req, timeout=120) as resp:
            result = json_lib.loads(resp.read().decode("utf-8"))

        text = result["content"][0]["text"].strip()
        text = text.replace("```json", "").replace("```", "").strip()
        parsed = json_lib.loads(text)

        capitulos = parsed.get("capitulos", [])
        if not capitulos:
            return jsonify({"ok": False, "msg": "Claude no pudo extraer actividades de los planos"}), 400

        resumen_capitulos = []
        detalle_completo = []
        total_mo_proyecto = 0
        total_mat_proyecto = 0

        for cap in capitulos:
            nombre_cap = cap.get("nombre", "Sin nombre")
            total_mo_cap = 0
            total_mat_cap = 0
            for act in cap.get("actividades", []):
                cantidad = float(act.get("cantidad", 0) or 0)
                pu_mo = float(act.get("pu_mo", 0) or 0)
                pu_mat = float(act.get("pu_mat", 0) or 0)
                sub_mo = cantidad * pu_mo
                sub_mat = cantidad * pu_mat
                total_mo_cap += sub_mo
                total_mat_cap += sub_mat
                detalle_completo.append({
                    "capitulo": nombre_cap,
                    "nombre": act.get("nombre", ""),
                    "unidad": act.get("unidad", ""),
                    "cantidad": cantidad,
                    "pu_mo": pu_mo,
                    "pu_mat": pu_mat,
                    "sub_mo": round(sub_mo),
                    "sub_mat": round(sub_mat),
                    "total": round(sub_mo + sub_mat)
                })
            total_mo_proyecto += total_mo_cap
            total_mat_proyecto += total_mat_cap
            resumen_capitulos.append({
                "nombre": nombre_cap,
                "monto": round(total_mo_cap + total_mat_cap)
            })

        PCT_ADMINISTRACION = 0.08
        PCT_UTILIDAD = 0.07

        costo_directo = total_mo_proyecto + total_mat_proyecto
        monto_administracion = round(costo_directo * PCT_ADMINISTRACION)
        monto_utilidad = round(costo_directo * PCT_UTILIDAD)
        total_con_admin_utilidad = costo_directo + monto_administracion + monto_utilidad

        return jsonify({
            "ok": True,
            "detalle": detalle_completo,
            "resumen_capitulos": resumen_capitulos,
            "total_mo": round(total_mo_proyecto),
            "total_mat": round(total_mat_proyecto),
            "costo_directo": round(costo_directo),
            "monto_administracion": monto_administracion,
            "monto_utilidad": monto_utilidad,
            "total_general": round(total_con_admin_utilidad)
        })

    except Exception as e:
        print(f"[PLANOS] ERROR: {str(e)}")
        return jsonify({"ok": False, "msg": f"Error al analizar planos: {str(e)[:200]}"}), 500


@app.route("/api/admin/descargar-presupuesto-planos", methods=["POST"])
def descargar_presupuesto_planos():
    """Genera el Excel detallado (estilo Casa Itaca) a partir del JSON de detalle."""
    if not session.get("admin"):
        return jsonify({"error": "No autorizado"}), 401

    data = request.json
    detalle = data.get("detalle", [])
    nombre_proyecto = data.get("nombre", "Proyecto")

    if not detalle:
        return jsonify({"ok": False, "msg": "Sin datos de detalle"}), 400

    excel_io = generar_excel_detallado_planos(nombre_proyecto, detalle)
    return send_file(excel_io, as_attachment=True,
        download_name=f"Presupuesto_Detallado_{nombre_proyecto.replace(' ','_')}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ═══════════════════════════════════════════════════════════════
# MÓDULO: CONTROL DE OBRA — CRONOGRAMA + BITÁCORA + FOTOS
# ═══════════════════════════════════════════════════════════════

@app.route("/api/cronograma/generar", methods=["POST"])
def generar_cronograma():
    """
    Genera el cronograma inicial de un proyecto usando Claude.
    Claude estima duración y solapamientos lógicos por capítulo.
    """
    if "tipo" not in session:
        return jsonify({"error": "No autorizado"}), 401

    data = request.json
    codigo = data.get("codigo", "").strip().upper()
    fecha_inicio = data.get("fecha_inicio", "")  # YYYY-MM-DD

    if not codigo or not fecha_inicio:
        return jsonify({"ok": False, "msg": "Falta código o fecha de inicio"}), 400

    proyectos = get_proyectos()
    if codigo not in proyectos:
        return jsonify({"ok": False, "msg": "Proyecto no encontrado"}), 404

    p = proyectos[codigo]
    # Agrupar partidas por capítulo (primeras 2 letras del código, ej CI, CC, POR)
    capitulos_raw = {}
    for cod_part, nom_part, monto in p["partidas"]:
        # Usar el nombre de la partida directamente como capítulo
        capitulos_raw[nom_part] = capitulos_raw.get(nom_part, 0) + monto

    capitulos_lista = [{"nombre": n, "monto": m} for n, m in capitulos_raw.items()]
    total_proyecto = sum(m for m in capitulos_raw.values())

    prompt = f"""Eres un ingeniero de construcción costarricense experto en programación de obras.

Proyecto: {p['nombre']}
Total presupuesto: ₡{total_proyecto:,}
Fecha de inicio: {fecha_inicio}

Capítulos del proyecto:
{json_lib.dumps(capitulos_lista, ensure_ascii=False, indent=2)}

Generá un cronograma realista con solapamientos lógicos de obra (como en Costa Rica: estructura puede iniciar cuando fundaciones van al 70%, acabados van escalonados, etc.).

Para cada capítulo indicá:
- duracion_semanas: duración estimada (puede ser decimal, ej: 2.5)
- inicio_offset_semanas: cuántas semanas desde el inicio del proyecto arranca este capítulo (0 = arranca desde el día 1)
- considera que capítulos que dependen de otros arrancan con el offset correcto

Devolvé ÚNICAMENTE JSON válido sin backticks:
{{
  "capitulos": [
    {{
      "nombre": "nombre exacto del capítulo",
      "duracion_semanas": 3.0,
      "inicio_offset_semanas": 0
    }}
  ],
  "duracion_total_semanas": 24
}}"""

    try:
        payload = json_lib.dumps({
            "model": "claude-sonnet-4-6",
            "max_tokens": 3000,
            "messages": [{"role": "user", "content": prompt}]
        }).encode("utf-8")

        ANTHROPIC_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
        req = urllib.request.Request(
            "https://api.anthropic.com/v1/messages",
            data=payload,
            headers={"Content-Type": "application/json", "x-api-key": ANTHROPIC_KEY, "anthropic-version": "2023-06-01"},
            method="POST"
        )
        with urllib.request.urlopen(req, timeout=60) as resp:
            result = json_lib.loads(resp.read().decode("utf-8"))

        text = result["content"][0]["text"].strip().replace("```json","").replace("```","").strip()
        parsed = json_lib.loads(text)

        from datetime import date, timedelta
        fecha_ini = date.fromisoformat(fecha_inicio)

        # Guardar en BD y devolver estructura
        conn = get_db()
        cur = conn.cursor()

        # Limpiar cronograma anterior si existe
        cur.execute("DELETE FROM cronograma WHERE proyecto_codigo = %s", (codigo,))

        capitulos_resultado = []
        for i, cap in enumerate(parsed.get("capitulos", [])):
            offset = float(cap.get("inicio_offset_semanas", 0))
            duracion = float(cap.get("duracion_semanas", 1))
            fi_plan = fecha_ini + timedelta(weeks=offset)
            ff_plan = fi_plan + timedelta(weeks=duracion)

            cur.execute("""
                INSERT INTO cronograma (proyecto_codigo, capitulo, orden, fecha_inicio_plan, fecha_fin_plan, duracion_semanas, pct_avance, estado)
                VALUES (%s, %s, %s, %s, %s, %s, 0, 'pendiente')
                RETURNING id
            """, (codigo, cap["nombre"], i, fi_plan, ff_plan, duracion))
            row_id = cur.fetchone()[0]

            capitulos_resultado.append({
                "id": row_id,
                "capitulo": cap["nombre"],
                "orden": i,
                "fecha_inicio_plan": fi_plan.isoformat(),
                "fecha_fin_plan": ff_plan.isoformat(),
                "duracion_semanas": duracion,
                "pct_avance": 0,
                "estado": "pendiente"
            })

        conn.commit()
        cur.close(); conn.close()

        return jsonify({
            "ok": True,
            "capitulos": capitulos_resultado,
            "duracion_total_semanas": parsed.get("duracion_total_semanas", 0),
            "fecha_inicio": fecha_inicio
        })

    except Exception as e:
        print(f"[CRONOGRAMA] ERROR: {e}")
        return jsonify({"ok": False, "msg": str(e)[:200]}), 500


@app.route("/api/cronograma/<codigo>")
def get_cronograma(codigo):
    """Lee el cronograma actual de un proyecto."""
    if "tipo" not in session:
        return jsonify({"error": "No autorizado"}), 401
    try:
        conn = get_db()
        cur = conn.cursor(cursor_factory=RealDictCursor)
        cur.execute("""
            SELECT id, capitulo, orden, fecha_inicio_plan, fecha_fin_plan,
                   fecha_inicio_real, fecha_fin_real, duracion_semanas,
                   pct_avance, estado
            FROM cronograma WHERE proyecto_codigo = %s ORDER BY orden
        """, (codigo,))
        rows = cur.fetchall()
        cur.close(); conn.close()
        return jsonify([dict(r) for r in rows])
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/cronograma/actualizar", methods=["POST"])
def actualizar_cronograma():
    """Actualiza % de avance, fechas reales y estado de un capítulo."""
    if "tipo" not in session:
        return jsonify({"error": "No autorizado"}), 401

    data = request.json
    cap_id = data.get("id")
    pct = data.get("pct_avance")
    estado = data.get("estado")
    fecha_inicio_real = data.get("fecha_inicio_real")
    fecha_fin_real = data.get("fecha_fin_real")

    try:
        conn = get_db()
        cur = conn.cursor()
        cur.execute("""
            UPDATE cronograma SET
                pct_avance = COALESCE(%s, pct_avance),
                estado = COALESCE(%s, estado),
                fecha_inicio_real = COALESCE(%s::date, fecha_inicio_real),
                fecha_fin_real = COALESCE(%s::date, fecha_fin_real)
            WHERE id = %s
            RETURNING proyecto_codigo, capitulo
        """, (pct, estado, fecha_inicio_real, fecha_fin_real, cap_id))
        row = cur.fetchone()
        conn.commit()
        cur.close(); conn.close()

        # Notificar al cliente
        if row:
            codigo_proy, capitulo = row
            email, nombre_cli = get_email_cliente(codigo_proy)
            if email:
                proyectos = get_proyectos()
                nombre_proy = proyectos.get(codigo_proy, {}).get("nombre", codigo_proy)
                enviar_notificacion_email(email, nombre_cli, nombre_proy, "avance",
                    f"Se actualizó el avance del capítulo <strong>{capitulo}</strong> a <strong>{pct}%</strong>.")

        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"ok": False, "msg": str(e)}), 500


@app.route("/api/bitacora/<codigo>")
def get_bitacora(codigo):
    """Lee entradas de bitácora de un proyecto."""
    if "tipo" not in session:
        return jsonify({"error": "No autorizado"}), 401
    try:
        conn = get_db()
        cur = conn.cursor(cursor_factory=RealDictCursor)
        cur.execute("""
            SELECT id, fecha, autor, titulo, contenido, capitulo, creado_en
            FROM bitacora WHERE proyecto_codigo = %s ORDER BY fecha DESC, creado_en DESC
        """, (codigo,))
        rows = cur.fetchall()
        cur.close(); conn.close()
        return jsonify([dict(r) for r in rows])
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/bitacora/agregar", methods=["POST"])
def agregar_bitacora():
    """Agrega una entrada de bitácora."""
    if "tipo" not in session:
        return jsonify({"error": "No autorizado"}), 401

    data = request.json
    codigo = data.get("codigo", "").upper()
    fecha = data.get("fecha", "")
    autor = data.get("autor", "Admin")
    titulo = data.get("titulo", "")
    contenido = data.get("contenido", "").strip()
    capitulo = data.get("capitulo", "")

    if not codigo or not contenido:
        return jsonify({"ok": False, "msg": "Faltan datos"}), 400
    try:
        conn = get_db()
        cur = conn.cursor()
        cur.execute("""
            INSERT INTO bitacora (proyecto_codigo, fecha, autor, titulo, contenido, capitulo)
            VALUES (%s, %s::date, %s, %s, %s, %s) RETURNING id
        """, (codigo, fecha or "today", autor, titulo, contenido, capitulo))
        new_id = cur.fetchone()[0]
        conn.commit()
        cur.close(); conn.close()

        # Notificar al cliente
        email, nombre_cli = get_email_cliente(codigo)
        if email:
            proyectos = get_proyectos()
            nombre_proy = proyectos.get(codigo, {}).get("nombre", codigo)
            detalle = f"<strong>{titulo}</strong><br>{contenido[:200]}{'...' if len(contenido)>200 else ''}"
            enviar_notificacion_email(email, nombre_cli, nombre_proy, "bitacora", detalle)

        return jsonify({"ok": True, "id": new_id})
    except Exception as e:
        return jsonify({"ok": False, "msg": str(e)}), 500


@app.route("/api/fotos/<codigo>")
def get_fotos(codigo):
    """Lee fotos de obra de un proyecto."""
    if "tipo" not in session:
        return jsonify({"error": "No autorizado"}), 401
    solo_aprobadas = request.args.get("aprobadas", "false") == "true"
    try:
        conn = get_db()
        cur = conn.cursor(cursor_factory=RealDictCursor)
        query = "SELECT id, fecha, autor, capitulo, nota, pct_estimado, pct_confirmado, aprobado FROM fotos_obra WHERE proyecto_codigo = %s"
        if solo_aprobadas:
            query += " AND aprobado = TRUE"
        query += " ORDER BY fecha DESC, id DESC"
        cur.execute(query, (codigo,))
        rows = cur.fetchall()
        cur.close(); conn.close()
        return jsonify([dict(r) for r in rows])
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/fotos/imagen/<int:foto_id>")
def get_foto_imagen(foto_id):
    """Devuelve la imagen base64 de una foto específica."""
    if "tipo" not in session:
        return jsonify({"error": "No autorizado"}), 401
    try:
        conn = get_db()
        cur = conn.cursor()
        cur.execute("SELECT imagen_b64 FROM fotos_obra WHERE id = %s", (foto_id,))
        row = cur.fetchone()
        cur.close(); conn.close()
        if not row or not row[0]:
            return jsonify({"error": "Sin imagen"}), 404
        return jsonify({"imagen_b64": row[0]})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/fotos/subir", methods=["POST"])
def subir_foto():
    """Sube una foto de obra, Claude estima % de avance del capítulo."""
    if "tipo" not in session:
        return jsonify({"error": "No autorizado"}), 401

    codigo = request.form.get("codigo", "").upper()
    capitulo = request.form.get("capitulo", "")
    nota = request.form.get("nota", "")
    autor = request.form.get("autor", "Admin")
    fecha = request.form.get("fecha", "")
    archivo = request.files.get("foto")

    if not archivo or not codigo:
        return jsonify({"ok": False, "msg": "Faltan datos"}), 400

    file_bytes = archivo.read()
    ext = archivo.filename.rsplit(".", 1)[-1].lower()
    media_type = f"image/{'jpeg' if ext in ['jpg','jpeg'] else ext}"
    b64_data = base64.b64encode(file_bytes).decode("utf-8")

    # Claude estima % de avance
    pct_estimado = None
    try:
        prompt_vision = f"""Sos un inspector de obras costarricense. Analizá esta foto de la actividad "{capitulo}" del proyecto y estimá el porcentaje de avance de esa actividad específica basándote en lo que ves.

Devolvé ÚNICAMENTE un JSON sin backticks:
{{"pct_avance": 65, "observacion": "Se observan las columnas levantadas hasta el 60% de altura, falta completar la estructura de la losa."}}

El pct_avance debe ser un número entero entre 0 y 100."""

        payload = json_lib.dumps({
            "model": "claude-sonnet-4-6",
            "max_tokens": 500,
            "messages": [{"role": "user", "content": [
                {"type": "image", "source": {"type": "base64", "media_type": media_type, "data": b64_data}},
                {"type": "text", "text": prompt_vision}
            ]}]
        }).encode("utf-8")

        ANTHROPIC_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
        req = urllib.request.Request(
            "https://api.anthropic.com/v1/messages",
            data=payload,
            headers={"Content-Type": "application/json", "x-api-key": ANTHROPIC_KEY, "anthropic-version": "2023-06-01"},
            method="POST"
        )
        with urllib.request.urlopen(req, timeout=60) as resp:
            result = json_lib.loads(resp.read().decode("utf-8"))
        text = result["content"][0]["text"].strip().replace("```json","").replace("```","").strip()
        parsed = json_lib.loads(text)
        pct_estimado = parsed.get("pct_avance")
        observacion_claude = parsed.get("observacion", "")
        if nota:
            nota = f"{nota}\n[Claude]: {observacion_claude}"
        else:
            nota = f"[Claude]: {observacion_claude}"
    except Exception as e:
        print(f"[FOTOS] Error Claude: {e}")

    try:
        conn = get_db()
        cur = conn.cursor()
        cur.execute("""
            INSERT INTO fotos_obra (proyecto_codigo, fecha, autor, capitulo, nota, imagen_b64, pct_estimado, aprobado)
            VALUES (%s, %s::date, %s, %s, %s, %s, %s, FALSE) RETURNING id
        """, (codigo, fecha or "today", autor, capitulo, nota, b64_data, pct_estimado))
        new_id = cur.fetchone()[0]
        conn.commit()
        cur.close(); conn.close()
        return jsonify({"ok": True, "id": new_id, "pct_estimado": pct_estimado, "nota_claude": nota})
    except Exception as e:
        return jsonify({"ok": False, "msg": str(e)}), 500


@app.route("/api/fotos/eliminar/<int:foto_id>", methods=["DELETE"])
def eliminar_foto(foto_id):
    """Elimina una foto de obra."""
    if not session.get("admin"):
        return jsonify({"error": "No autorizado"}), 401
    try:
        conn = get_db()
        cur = conn.cursor()
        cur.execute("DELETE FROM fotos_obra WHERE id = %s", (foto_id,))
        conn.commit()
        cur.close(); conn.close()
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"ok": False, "msg": str(e)}), 500


@app.route("/api/bitacora/eliminar/<int:entrada_id>", methods=["DELETE"])
def eliminar_bitacora(entrada_id):
    """Elimina una entrada de bitácora."""
    if not session.get("admin"):
        return jsonify({"error": "No autorizado"}), 401
    try:
        conn = get_db()
        cur = conn.cursor()
        cur.execute("DELETE FROM bitacora WHERE id = %s", (entrada_id,))
        conn.commit()
        cur.close(); conn.close()
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"ok": False, "msg": str(e)}), 500


@app.route("/api/fotos/aprobar", methods=["POST"])
def aprobar_foto():
    """Admin aprueba una foto y confirma el % de avance."""
    if not session.get("admin"):
        return jsonify({"error": "No autorizado"}), 401

    data = request.json
    foto_id = data.get("id")
    pct_confirmado = data.get("pct_confirmado")

    try:
        conn = get_db()
        cur = conn.cursor()
        cur.execute("""
            UPDATE fotos_obra SET aprobado = TRUE, pct_confirmado = %s WHERE id = %s
            RETURNING proyecto_codigo, capitulo
        """, (pct_confirmado, foto_id))
        row = cur.fetchone()
        conn.commit()

        # Actualizar cronograma (separado para no bloquear si falla)
        if row and pct_confirmado is not None:
            try:
                cur.execute("""
                    UPDATE cronograma SET pct_avance = %s,
                        estado = CASE WHEN %s >= 100 THEN 'completado'
                                      WHEN %s > 0 THEN 'en_curso'
                                      ELSE estado END
                    WHERE proyecto_codigo = %s AND LOWER(capitulo) = LOWER(%s)
                """, (pct_confirmado, pct_confirmado, pct_confirmado, row[0], row[1]))
                conn.commit()
            except Exception as e2:
                print(f"[APROBAR] No se pudo actualizar cronograma: {e2}")

        cur.close(); conn.close()

        # Notificar al cliente (no bloquea si falla)
        if row:
            try:
                codigo_proy, capitulo = row
                email, nombre_cli = get_email_cliente(codigo_proy)
                if email:
                    proyectos = get_proyectos()
                    nombre_proy = proyectos.get(codigo_proy, {}).get("nombre", codigo_proy)
                    enviar_notificacion_email(email, nombre_cli, nombre_proy, "foto",
                        f"Nueva foto aprobada del capítulo <strong>{capitulo}</strong> con avance de <strong>{pct_confirmado}%</strong>.")
            except Exception as e3:
                print(f"[APROBAR] Error en notificación: {e3}")

        return jsonify({"ok": True})
    except Exception as e:
        print(f"[APROBAR] Error: {e}")
        return jsonify({"ok": False, "msg": str(e)}), 500

# ═══════════════════════════════════════════════════════════════
# MÓDULO: GASTOS MANUALES (sin factura)
# ═══════════════════════════════════════════════════════════════

@app.route("/api/gastos/<codigo>")
def get_gastos(codigo):
    """Lee gastos manuales de un proyecto."""
    if "tipo" not in session:
        return jsonify({"error": "No autorizado"}), 401
    try:
        conn = get_db()
        cur = conn.cursor(cursor_factory=RealDictCursor)
        cur.execute("""
            SELECT id, fecha, partida_codigo, partida_nombre, descripcion,
                   monto, tipo, comprobante_tipo, creado_en
            FROM gastos_manuales
            WHERE proyecto_codigo = %s
            ORDER BY fecha DESC, id DESC
        """, (codigo,))
        rows = cur.fetchall()
        cur.close(); conn.close()
        return jsonify([dict(r) for r in rows])
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/gastos/comprobante/<int:gasto_id>")
def get_comprobante(gasto_id):
    """Devuelve el comprobante base64 de un gasto específico."""
    if "tipo" not in session:
        return jsonify({"error": "No autorizado"}), 401
    try:
        conn = get_db()
        cur = conn.cursor()
        cur.execute("SELECT comprobante_b64, comprobante_tipo FROM gastos_manuales WHERE id = %s", (gasto_id,))
        row = cur.fetchone()
        cur.close(); conn.close()
        if not row or not row[0]:
            return jsonify({"error": "Sin comprobante"}), 404
        return jsonify({"comprobante_b64": row[0], "tipo": row[1]})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/gastos/agregar", methods=["POST"])
def agregar_gasto():
    """Registra un gasto manual con comprobante opcional."""
    if not session.get("admin"):
        return jsonify({"error": "No autorizado"}), 401

    codigo = request.form.get("codigo", "").upper()
    fecha = request.form.get("fecha", "")
    partida_codigo = request.form.get("partida_codigo", "")
    partida_nombre = request.form.get("partida_nombre", "")
    descripcion = request.form.get("descripcion", "").strip()
    monto = request.form.get("monto", "0")
    tipo = request.form.get("tipo", "efectivo")
    archivo = request.files.get("comprobante")

    if not codigo or not descripcion or not partida_codigo:
        return jsonify({"ok": False, "msg": "Faltan datos obligatorios"}), 400

    try:
        monto_num = float(monto.replace(",","").replace("₡",""))
    except:
        return jsonify({"ok": False, "msg": "Monto inválido"}), 400

    comprobante_b64 = None
    comprobante_tipo = None

    if archivo and archivo.filename:
        ext = archivo.filename.rsplit(".", 1)[-1].lower()
        file_bytes = archivo.read()
        comprobante_b64 = base64.b64encode(file_bytes).decode("utf-8")
        comprobante_tipo = "pdf" if ext == "pdf" else "img"

    try:
        conn = get_db()
        cur = conn.cursor()
        cur.execute("""
            INSERT INTO gastos_manuales
                (proyecto_codigo, fecha, partida_codigo, partida_nombre,
                 descripcion, monto, tipo, comprobante_b64, comprobante_tipo)
            VALUES (%s, %s::date, %s, %s, %s, %s, %s, %s, %s)
            RETURNING id
        """, (codigo, fecha or "today", partida_codigo, partida_nombre,
              descripcion, monto_num, tipo, comprobante_b64, comprobante_tipo))
        new_id = cur.fetchone()[0]
        conn.commit()
        cur.close(); conn.close()

        # Notificar al cliente
        try:
            email, nombre_cli = get_email_cliente(codigo)
            if email:
                proyectos = get_proyectos()
                nombre_proy = proyectos.get(codigo, {}).get("nombre", codigo)
                enviar_notificacion_email(email, nombre_cli, nombre_proy, "avance",
                    f"Se registró un gasto de <strong>₡{monto_num:,.0f}</strong> en la partida <strong>{partida_nombre}</strong>.")
        except Exception as e:
            print(f"[GASTOS] Error notificación: {e}")

        return jsonify({"ok": True, "id": new_id})
    except Exception as e:
        print(f"[GASTOS] Error: {e}")
        return jsonify({"ok": False, "msg": str(e)}), 500


@app.route("/api/gastos/eliminar/<int:gasto_id>", methods=["DELETE"])
def eliminar_gasto(gasto_id):
    """Elimina un gasto manual."""
    if not session.get("admin"):
        return jsonify({"error": "No autorizado"}), 401
    try:
        conn = get_db()
        cur = conn.cursor()
        cur.execute("DELETE FROM gastos_manuales WHERE id = %s", (gasto_id,))
        conn.commit()
        cur.close(); conn.close()
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"ok": False, "msg": str(e)}), 500


@app.route("/api/gastos/totales/<codigo>")
def get_gastos_totales(codigo):
    """Devuelve el total de gastos manuales agrupado por partida."""
    if "tipo" not in session:
        return jsonify({"error": "No autorizado"}), 401
    try:
        conn = get_db()
        cur = conn.cursor(cursor_factory=RealDictCursor)
        cur.execute("""
            SELECT partida_codigo, SUM(monto) as total
            FROM gastos_manuales
            WHERE proyecto_codigo = %s
            GROUP BY partida_codigo
        """, (codigo,))
        rows = cur.fetchall()
        cur.close(); conn.close()
        return jsonify({r["partida_codigo"]: float(r["total"]) for r in rows})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# ═══════════════════════════════════════════════════════════════
# MÓDULO: INFORMES PDF PARA CLIENTE
# ═══════════════════════════════════════════════════════════════
from reportlab.lib.pagesizes import letter, A4
from reportlab.lib import colors
from reportlab.lib.units import cm, mm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_LEFT, TA_JUSTIFY
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
    TableStyle, PageBreak, HRFlowable, Image as RLImage, KeepTogether)
from reportlab.graphics.shapes import Drawing, Rect, String, Line
from reportlab.graphics.charts.barcharts import VerticalBarChart
from reportlab.graphics import renderPDF

# Colores corporativos Urbanistyka
URB_BLACK  = colors.HexColor("#2E2E2E")  # Gris oscuro VAUMA
URB_DARK   = colors.HexColor("#0d1410")
URB_GOLD   = colors.HexColor("#FFE500")  # Amarillo VAUMA
URB_GOLD2  = colors.HexColor("#FFD000")  # Amarillo VAUMA oscuro
URB_GREEN  = colors.HexColor("#22c55e")
URB_RED    = colors.HexColor("#ef4444")
URB_YELLOW = colors.HexColor("#f59e0b")
URB_GRAY   = colors.HexColor("#1B4B5A")  # Verde azulado VAUMA
URB_LIGHT  = colors.HexColor("#F5F5F5")  # Blanco VAUMA
URB_BORDER = colors.HexColor("#1a2d1e")

def fmt_colones(n):
    """Formato colones Costa Rica: C/54.924.000"""
    num = f"{int(n):,}".replace(",", ".")
    return f"C/{num}"

def get_informe_data(codigo_proyecto):
    """Obtiene todos los datos necesarios para el informe."""
    proyectos = get_proyectos()
    if codigo_proyecto not in proyectos:
        return None, None, None, None

    p = proyectos[codigo_proyecto]

    # Datos financieros
    try:
        gastos_odoo, detalle = obtener_datos_proyecto(p["nombre"], p["partidas"])
    except:
        gastos_odoo = {c: 0 for c,_,_ in p["partidas"]}
        detalle = {c: [] for c,_,_ in p["partidas"]}

    # Gastos manuales
    try:
        conn = get_db()
        cur = conn.cursor(cursor_factory=RealDictCursor)
        cur.execute("""
            SELECT partida_codigo, SUM(monto) as total
            FROM gastos_manuales WHERE proyecto_codigo = %s
            GROUP BY partida_codigo
        """, (codigo_proyecto,))
        gastos_manuales = {r["partida_codigo"]: float(r["total"]) for r in cur.fetchall()}
        cur.close(); conn.close()
    except:
        gastos_manuales = {}

    # Cronograma
    try:
        conn = get_db()
        cur = conn.cursor(cursor_factory=RealDictCursor)
        cur.execute("SELECT * FROM cronograma WHERE proyecto_codigo = %s ORDER BY orden", (codigo_proyecto,))
        cronograma = [dict(r) for r in cur.fetchall()]
        cur.close(); conn.close()
    except:
        cronograma = []

    # Fotos aprobadas (últimas 6)
    try:
        conn = get_db()
        cur = conn.cursor(cursor_factory=RealDictCursor)
        cur.execute("""
            SELECT id, fecha, capitulo, nota, imagen_b64, pct_confirmado
            FROM fotos_obra WHERE proyecto_codigo = %s AND aprobado = TRUE
            ORDER BY fecha DESC LIMIT 6
        """, (codigo_proyecto,))
        fotos = [dict(r) for r in cur.fetchall()]
        cur.close(); conn.close()
    except:
        fotos = []

    # Bitácora (últimas 5 entradas)
    try:
        conn = get_db()
        cur = conn.cursor(cursor_factory=RealDictCursor)
        cur.execute("""
            SELECT fecha, autor, titulo, contenido, capitulo
            FROM bitacora WHERE proyecto_codigo = %s
            ORDER BY fecha DESC LIMIT 5
        """, (codigo_proyecto,))
        bitacora = [dict(r) for r in cur.fetchall()]
        cur.close(); conn.close()
    except:
        bitacora = []

    # Construir partidas con totales
    partidas_data = []
    for cod, nom, presup in p["partidas"]:
        g = gastos_odoo.get(cod, 0) + gastos_manuales.get(cod, 0)
        partidas_data.append({
            "codigo": cod, "nombre": nom, "presupuesto": presup,
            "gastado": g, "saldo": presup - g,
            "pct": round(g/presup*100, 1) if presup > 0 else 0
        })

    return p, partidas_data, cronograma, fotos, bitacora


def build_header(canvas_obj, doc, proyecto_nombre, tipo_informe, logo_b64=None):
    """Header y footer en cada página."""
    canvas_obj.saveState()
    w, h = A4

    # Banda superior negra
    canvas_obj.setFillColor(URB_BLACK)
    canvas_obj.rect(0, h-40*mm, w, 40*mm, fill=1, stroke=0)

    # Logo (PNG compuesto sobre fondo negro)
    if logo_b64:
        try:
            logo_bytes = base64.b64decode(logo_b64)
            logo_buf = io.BytesIO(logo_bytes)
            # Logo es 679x188px — ratio 3.6:1, usamos 5cm de ancho → 1.4cm alto
            logo_img = RLImage(logo_buf, width=5*cm, height=1.4*cm)
            logo_img.drawOn(canvas_obj, 1.2*cm, h-32*mm)
        except Exception as e:
            print(f"[PDF] Error logo: {e}")
            # Fallback texto
            canvas_obj.setFillColor(URB_GOLD)
            canvas_obj.setFont("Helvetica-Bold", 15)
            canvas_obj.drawString(1.5*cm, h-18*mm, "VAUMA")
            canvas_obj.setFillColor(URB_GRAY)
            canvas_obj.setFont("Helvetica", 7)
            canvas_obj.drawString(1.5*cm, h-25*mm, "VARGAS ULLOA MAQUINARIA S.A.")

    # Tipo de informe y proyecto (derecha)
    canvas_obj.setFillColor(URB_LIGHT)
    canvas_obj.setFont("Helvetica-Bold", 10)
    canvas_obj.drawRightString(w-1.5*cm, h-18*mm, tipo_informe.upper())
    canvas_obj.setFillColor(URB_GRAY)
    canvas_obj.setFont("Helvetica", 8)
    canvas_obj.drawRightString(w-1.5*cm, h-26*mm, proyecto_nombre)

    # Línea dorada
    canvas_obj.setStrokeColor(URB_GOLD)
    canvas_obj.setLineWidth(2)
    canvas_obj.line(0, h-40*mm, w, h-40*mm)

    # Footer
    canvas_obj.setFillColor(URB_BLACK)
    canvas_obj.rect(0, 0, w, 10*mm, fill=1, stroke=0)
    canvas_obj.setFillColor(URB_GRAY)
    canvas_obj.setFont("Helvetica", 7)
    canvas_obj.drawString(1.5*cm, 3*mm,
        f"Generado el {datetime.now().strftime('%d/%m/%Y %H:%M')}  |  Urbanistyka Constructora  |  Costa Rica")
    canvas_obj.drawRightString(w-1.5*cm, 3*mm, f"Pag. {canvas_obj.getPageNumber()}")

    canvas_obj.restoreState()


def generar_pdf_informe(codigo_proyecto, tipo="completo"):
    """
    Genera el PDF del informe.
    tipo: 'financiero' | 'avance' | 'completo'
    """
    result = get_informe_data(codigo_proyecto)
    if result[0] is None:
        return None

    p, partidas_data, cronograma, fotos, bitacora = result

    nombres_tipo = {
        "financiero": "Informe Financiero",
        "avance":     "Informe de Avance de Obra",
        "completo":   "Informe de Proyecto"
    }
    nombre_tipo = nombres_tipo.get(tipo, "Informe de Proyecto")

    output = io.BytesIO()
    doc = SimpleDocTemplate(output, pagesize=A4,
        leftMargin=1.5*cm, rightMargin=1.5*cm,
        topMargin=45*mm, bottomMargin=18*mm)

    # Logo VAUMA embebido
    logo_b64 = "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Rpv/AEFrn/vhKf8A8Ko0b/oI3n/jlR/aWHD+wsX/ACnl/wAtHy16knwr8P8A/LS+v/8AvtP/AIipU+GPhxP+Wl4//bb/AOwo/tSiaf6v4rseUcetHHrXrqfDfwqn34Jn/wC2z1Knw98Jf9Arf/22f/4uo/tSkX/q7izx75aTivZ4fBPhWH/V6ND/AMDd3qwnhXw4n3NDs/8AgcKVH9rwNo8O1n9o8Q4o/Cvek0fSofuaVZp/uQpVhIYIf9XGif7iVH9sf3TWHDM/tyPBUsL2b/V2kz/7iVah8OeIJvuaHf8A/fl69zoqP7Yl/KbQ4Zh9uZ4wngvxXN/q9Hm/4G6JVqL4ceKn/wBZBDD/AL81eu0Vh/a9U3hw7hzzCL4W67IP9Iv7NP8Ac3vVyH4Syf8ALfXP++Lb/wCzr0Oisv7SxB0RyPCQ+ycTD8K9DT/X314/+5sStCH4deFYf9ZYzTf78z101FYzxlaf2jphluEh9gyYfCXhm2/1eh2f/A031oQ2dpbf6i1hh/3E2VNRWU685nRChRp/BAKKKKzubWQUUUUDCiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKAP//Z"

    def on_page(canvas_obj, doc_obj):
        build_header(canvas_obj, doc_obj, p["nombre"], nombre_tipo, logo_b64)

    # Estilos
    styles = getSampleStyleSheet()
    st_titulo = ParagraphStyle("titulo", fontName="Helvetica-Bold", fontSize=14,
        textColor=URB_GOLD, spaceAfter=4*mm)
    st_subtitulo = ParagraphStyle("subtitulo", fontName="Helvetica-Bold", fontSize=10,
        textColor=URB_LIGHT, spaceAfter=3*mm)
    st_normal = ParagraphStyle("normal", fontName="Helvetica", fontSize=9,
        textColor=colors.HexColor("#555555"), spaceAfter=2*mm, leading=14)
    st_bold = ParagraphStyle("bold", fontName="Helvetica-Bold", fontSize=9,
        textColor=URB_BLACK, spaceAfter=2*mm)
    st_center = ParagraphStyle("center", fontName="Helvetica", fontSize=9,
        alignment=TA_CENTER, textColor=colors.HexColor("#333333"))

    story = []

    # ── INFO GENERAL ────────────────────────────────────────
    story.append(Paragraph(p["nombre"], st_titulo))
    story.append(Paragraph(f"Cliente: {p['cliente']}", st_normal))
    story.append(Paragraph(f"Código: {codigo_proyecto}  ·  Fecha del informe: {datetime.now().strftime('%d de %B de %Y')}", st_normal))
    story.append(HRFlowable(width="100%", thickness=1, color=URB_GOLD, spaceAfter=5*mm))

    # ── SECCIÓN FINANCIERA ────────────────────────────────────
    if tipo in ("financiero", "completo"):
        story.append(Paragraph("CONTROL FINANCIERO", st_titulo))

        total_presup = sum(pt["presupuesto"] for pt in partidas_data)
        total_gast   = sum(pt["gastado"] for pt in partidas_data)
        total_saldo  = total_presup - total_gast
        pct_global   = round(total_gast/total_presup*100, 1) if total_presup > 0 else 0

        # KPIs en tabla
        kpi_data = [
            ["PRESUPUESTO TOTAL", "EJECUTADO", "SALDO DISPONIBLE", "% EJECUCIÓN"],
            [fmt_colones(total_presup), fmt_colones(total_gast),
             fmt_colones(total_saldo), f"{pct_global}%"]
        ]
        kpi_table = Table(kpi_data, colWidths=[4.5*cm, 4.5*cm, 4.5*cm, 4*cm])
        kpi_style = TableStyle([
            ("BACKGROUND",   (0,0), (-1,0), URB_BLACK),
            ("TEXTCOLOR",    (0,0), (-1,0), URB_GRAY),
            ("FONTNAME",     (0,0), (-1,0), "Helvetica"),
            ("FONTSIZE",     (0,0), (-1,0), 7),
            ("BACKGROUND",   (0,1), (-1,1), colors.HexColor("#f8f9fa")),
            ("FONTNAME",     (0,1), (-1,1), "Helvetica-Bold"),
            ("FONTSIZE",     (0,1), (-1,1), 12),
            ("TEXTCOLOR",    (0,1), (0,1), URB_GOLD),
            ("TEXTCOLOR",    (1,1), (1,1), URB_YELLOW if pct_global>=80 else URB_GOLD),
            ("TEXTCOLOR",    (2,1), (2,1), URB_GREEN),
            ("TEXTCOLOR",    (3,1), (3,1), URB_RED if pct_global>=100 else URB_YELLOW if pct_global>=80 else URB_GOLD),
            ("ALIGN",        (0,0), (-1,-1), "CENTER"),
            ("VALIGN",       (0,0), (-1,-1), "MIDDLE"),
            ("ROWHEIGHT",    (0,0), (-1,0), 8*mm),
            ("ROWHEIGHT",    (0,1), (-1,1), 14*mm),
            ("GRID",         (0,0), (-1,-1), 0.5, colors.HexColor("#dddddd")),
            ("ROUNDEDCORNERS", [3]),
        ])
        kpi_table.setStyle(kpi_style)
        story.append(kpi_table)
        story.append(Spacer(1, 5*mm))

        # Barra de avance global
        bar_w = 17*cm
        bar_h = 8*mm
        fill_w = bar_w * min(pct_global/100, 1)
        d = Drawing(bar_w, bar_h + 10)
        # Fondo
        d.add(Rect(0, 5, bar_w, bar_h, fillColor=colors.HexColor("#e5e7eb"), strokeColor=None))
        # Relleno
        bar_color = URB_RED if pct_global >= 100 else URB_YELLOW if pct_global >= 80 else URB_GOLD
        d.add(Rect(0, 5, fill_w, bar_h, fillColor=bar_color, strokeColor=None))
        # Texto
        d.add(String(bar_w/2, 8, f"Avance Global: {pct_global}%",
            fontName="Helvetica-Bold", fontSize=9, fillColor=colors.white,
            textAnchor="middle"))
        story.append(d)
        story.append(Spacer(1, 6*mm))

        # Tabla de partidas
        story.append(Paragraph("Detalle por Partida", st_subtitulo))
        part_headers = [["CÓD", "PARTIDA", "PRESUPUESTO", "EJECUTADO", "SALDO", "%", "ESTADO"]]
        part_rows = []
        for pt in partidas_data:
            estado = "EXCEDIDO" if pt["pct"]>=100 else "ALERTA" if pt["pct"]>=80 else "EN CURSO" if pt["gastado"]>0 else "SIN INICIO"
            part_rows.append([
                pt["codigo"],
                Paragraph(pt["nombre"], ParagraphStyle("pn", fontName="Helvetica", fontSize=8, leading=10)),
                fmt_colones(pt["presupuesto"]),
                fmt_colones(pt["gastado"]),
                fmt_colones(pt["saldo"]),
                f"{pt['pct']}%",
                estado
            ])
        part_rows.append([
            "TOTAL", Paragraph("<b>PROYECTO</b>", ParagraphStyle("pt", fontName="Helvetica-Bold", fontSize=8)),
            fmt_colones(total_presup), fmt_colones(total_gast),
            fmt_colones(total_saldo), f"{pct_global}%", ""
        ])

        col_w = [1.4*cm, 5.5*cm, 2.8*cm, 2.8*cm, 2.8*cm, 1.2*cm, 1.8*cm]
        part_table = Table(part_headers + part_rows, colWidths=col_w, repeatRows=1)
        ts = TableStyle([
            ("BACKGROUND",  (0,0), (-1,0), URB_BLACK),
            ("TEXTCOLOR",   (0,0), (-1,0), URB_GRAY),
            ("FONTNAME",    (0,0), (-1,0), "Helvetica-Bold"),
            ("FONTSIZE",    (0,0), (-1,0), 7),
            ("ALIGN",       (0,0), (-1,-1), "CENTER"),
            ("ALIGN",       (1,0), (1,-1), "LEFT"),
            ("VALIGN",      (0,0), (-1,-1), "MIDDLE"),
            ("FONTNAME",    (0,1), (-1,-2), "Helvetica"),
            ("FONTSIZE",    (0,1), (-1,-1), 8),
            ("ROWHEIGHT",   (0,0), (-1,0), 8*mm),
            ("GRID",        (0,0), (-1,-1), 0.3, colors.HexColor("#dddddd")),
            ("BACKGROUND",  (0,-1), (-1,-1), URB_BLACK),
            ("TEXTCOLOR",   (0,-1), (-1,-1), URB_LIGHT),
            ("FONTNAME",    (0,-1), (-1,-1), "Helvetica-Bold"),
        ])
        # Colorear filas por estado
        for i, pt in enumerate(partidas_data):
            r = i + 1
            bg = colors.HexColor("#fef2f2") if pt["pct"]>=100 else \
                 colors.HexColor("#fffbeb") if pt["pct"]>=80 else \
                 colors.HexColor("#f0fdf4") if pt["gastado"]>0 else \
                 colors.HexColor("#f8f9fa") if i%2==0 else colors.white
            ts.add("BACKGROUND", (0,r), (-1,r), bg)
        part_table.setStyle(ts)
        story.append(part_table)
        story.append(Spacer(1, 6*mm))

        # Gráfico de barras por partida
        story.append(Paragraph("Ejecucion Presupuestaria por Partida", st_subtitulo))
        if partidas_data:
            chart_w = 17 * cm
            chart_h = 6 * cm
            margin_b = 18
            d = Drawing(chart_w, chart_h + margin_b)
            max_val = max((pt["presupuesto"] for pt in partidas_data), default=1)
            n = len(partidas_data)
            slot_w = chart_w / n

            for i, pt in enumerate(partidas_data):
                x = i * slot_w
                bw = slot_w * 0.8
                bx = x + (slot_w - bw) / 2

                bh = (pt["presupuesto"] / max_val) * chart_h
                d.add(Rect(bx, margin_b, bw, bh,
                    fillColor=colors.HexColor("#d1d5db"), strokeColor=None))

                if pt["gastado"] > 0:
                    be = (pt["gastado"] / max_val) * chart_h
                    bc = URB_RED if pt["pct"] >= 100 else URB_YELLOW if pt["pct"] >= 80 else URB_GOLD
                    d.add(Rect(bx, margin_b, bw * 0.5, be, fillColor=bc, strokeColor=None))
                    d.add(String(bx + bw * 0.25, margin_b + be + 2,
                        f"{pt['pct']}%",
                        fontName="Helvetica-Bold", fontSize=5.5,
                        fillColor=bc, textAnchor="middle"))

                label = pt["codigo"].split("-")[-1]
                d.add(String(bx + bw / 2, 4, label,
                    fontName="Helvetica", fontSize=6,
                    fillColor=colors.HexColor("#555555"), textAnchor="middle"))

            story.append(d)

            ley_data = [["  Presupuesto", "  Ejecutado"]]
            ley_t = Table(ley_data, colWidths=[4*cm, 4*cm])
            ley_t.setStyle(TableStyle([
                ("BACKGROUND", (0,0), (0,0), colors.HexColor("#d1d5db")),
                ("BACKGROUND", (1,0), (1,0), URB_GOLD),
                ("FONTSIZE", (0,0), (-1,-1), 7),
                ("FONTNAME", (0,0), (-1,-1), "Helvetica"),
                ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
                ("ROWHEIGHT", (0,0), (-1,-1), 5*mm),
                ("TEXTCOLOR", (0,0), (0,0), colors.HexColor("#333333")),
                ("TEXTCOLOR", (1,0), (1,0), colors.black),
            ]))
            story.append(ley_t)
        story.append(Spacer(1, 4*mm))

    # ── SECCIÓN AVANCE DE OBRA ────────────────────────────────
    if tipo in ("avance", "completo"):
        if tipo == "completo":
            story.append(PageBreak())
        story.append(Paragraph("AVANCE DE OBRA", st_titulo))

        # Cronograma
        if cronograma:
            pct_obra = round(sum(c["pct_avance"] or 0 for c in cronograma)/len(cronograma), 1)
            completados = sum(1 for c in cronograma if c["estado"]=="completado")
            story.append(Paragraph(
                f"Avance general: <b>{pct_obra}%</b>  ·  Capítulos completados: <b>{completados}/{len(cronograma)}</b>",
                ParagraphStyle("inf", fontName="Helvetica", fontSize=9, textColor=colors.HexColor("#333333"), spaceAfter=4*mm)
            ))
            story.append(Paragraph("Cronograma de Capítulos", st_subtitulo))

            cron_headers = [["CAPÍTULO", "INICIO PLAN", "FIN PLAN", "SEMANAS", "AVANCE", "ESTADO"]]
            cron_rows = []
            for cap in cronograma:
                fi = cap["fecha_inicio_plan"].strftime("%d/%m/%Y") if cap.get("fecha_inicio_plan") else "—"
                ff = cap["fecha_fin_plan"].strftime("%d/%m/%Y") if cap.get("fecha_fin_plan") else "—"
                estado = {"pendiente":"Pendiente","en_curso":"En curso","completado":"Completado","atrasado":"Atrasado"}.get(cap.get("estado",""), "—")
                cron_rows.append([
                    Paragraph(cap["capitulo"], ParagraphStyle("cn", fontName="Helvetica", fontSize=8, leading=10)),
                    fi, ff, f"{cap.get('duracion_semanas','—')}s",
                    f"{cap.get('pct_avance',0)}%", estado
                ])

            cron_table = Table(cron_headers + cron_rows,
                colWidths=[5.5*cm, 2.5*cm, 2.5*cm, 2*cm, 2*cm, 2.8*cm], repeatRows=1)
            cts = TableStyle([
                ("BACKGROUND", (0,0), (-1,0), URB_BLACK),
                ("TEXTCOLOR",  (0,0), (-1,0), URB_GRAY),
                ("FONTNAME",   (0,0), (-1,0), "Helvetica-Bold"),
                ("FONTSIZE",   (0,0), (-1,0), 7),
                ("ALIGN",      (1,0), (-1,-1), "CENTER"),
                ("VALIGN",     (0,0), (-1,-1), "MIDDLE"),
                ("FONTNAME",   (0,1), (-1,-1), "Helvetica"),
                ("FONTSIZE",   (0,1), (-1,-1), 8),
                ("GRID",       (0,0), (-1,-1), 0.3, colors.HexColor("#dddddd")),
            ])
            for i, cap in enumerate(cronograma):
                r = i+1
                pct = cap.get("pct_avance",0) or 0
                bg = colors.HexColor("#f0fdf4") if cap.get("estado")=="completado" else \
                     colors.HexColor("#fef2f2") if cap.get("estado")=="atrasado" else \
                     colors.HexColor("#fffbeb") if pct>0 else \
                     colors.HexColor("#f8f9fa") if i%2==0 else colors.white
                cts.add("BACKGROUND", (0,r), (-1,r), bg)
            cron_table.setStyle(cts)
            story.append(cron_table)
            story.append(Spacer(1, 6*mm))

        # Fotos aprobadas — 2 por fila, más compactas con nota completa
        if fotos:
            story.append(Paragraph("Registro Fotografico de Obra", st_subtitulo))
            foto_rows = []
            foto_pair = []
            for f in fotos:
                try:
                    img_data = f.get("imagen_b64") or f.get("imagen_b64".lower())
                    if not img_data:
                        print(f"[PDF] Foto {f.get('id')} sin imagen_b64, keys: {list(f.keys())}")
                        continue
                    img_bytes = base64.b64decode(img_data)
                    img = RLImage(io.BytesIO(img_bytes), width=7.8*cm, height=5.2*cm)
                    fecha_obj = f.get("fecha")
                    fecha_str = fecha_obj.strftime("%d/%m/%Y") if hasattr(fecha_obj, "strftime") else str(fecha_obj or "")
                    capitulo = f.get("capitulo") or "General"
                    pct = f.get("pct_confirmado")
                    nota = (f.get("nota") or "").replace("[Claude]: ", "").strip()

                    st_cap = ParagraphStyle("fc", fontName="Helvetica-Bold", fontSize=8,
                        textColor=URB_BLACK, spaceAfter=1*mm)
                    st_meta = ParagraphStyle("fm", fontName="Helvetica", fontSize=7,
                        textColor=URB_GRAY, spaceAfter=1*mm)
                    st_pct = ParagraphStyle("fp", fontName="Helvetica-Bold", fontSize=9,
                        textColor=URB_GOLD, spaceAfter=1*mm)
                    st_nota = ParagraphStyle("fn", fontName="Helvetica", fontSize=7,
                        textColor=colors.HexColor("#555555"), leading=10)

                    cell_items = [
                        img,
                        Paragraph(capitulo, st_cap),
                        Paragraph(fecha_str, st_meta),
                    ]
                    if pct is not None:
                        cell_items.append(Paragraph(f"{pct}% avance confirmado", st_pct))
                    if nota:
                        cell_items.append(Paragraph(nota, st_nota))

                    foto_pair.append(cell_items)
                except Exception as e:
                    print(f"[PDF] Error foto: {e}")
                    continue
                if len(foto_pair) == 2:
                    foto_rows.append(foto_pair)
                    foto_pair = []
            if foto_pair:
                foto_rows.append(foto_pair + [[Spacer(1, 1)]])

            for row in foto_rows:
                ft = Table([row], colWidths=[8.5*cm, 8.5*cm])
                ft.setStyle(TableStyle([
                    ("VALIGN",  (0,0), (-1,-1), "TOP"),
                    ("TOPPADDING", (0,0), (-1,-1), 2*mm),
                    ("BOTTOMPADDING", (0,0), (-1,-1), 2*mm),
                    ("LEFTPADDING", (0,0), (-1,-1), 2*mm),
                    ("RIGHTPADDING", (0,0), (-1,-1), 2*mm),
                    ("LINEBELOW", (0,0), (-1,-1), 0.3, colors.HexColor("#eeeeee")),
                ]))
                story.append(ft)
                story.append(Spacer(1, 4*mm))

        # Bitácora — texto completo, sin truncar
        if bitacora:
            story.append(Spacer(1, 4*mm))
            story.append(Paragraph("Bitacora de Obra", st_subtitulo))
            for entrada in bitacora:
                fecha_obj = entrada.get("fecha")
                fecha_str = fecha_obj.strftime("%d/%m/%Y") if hasattr(fecha_obj, "strftime") else str(fecha_obj or "")
                meta = f"{fecha_str}  |  {entrada.get('autor','Admin')}"
                if entrada.get("capitulo"):
                    meta += f"  |  {entrada['capitulo']}"
                contenido_txt = entrada.get("contenido", "").replace("<", "&lt;").replace(">", "&gt;")

                items = [Paragraph(meta,
                    ParagraphStyle("bm", fontName="Helvetica", fontSize=7,
                        textColor=URB_GRAY, spaceAfter=2*mm))]
                if entrada.get("titulo"):
                    items.append(Paragraph(entrada["titulo"],
                        ParagraphStyle("bt", fontName="Helvetica-Bold", fontSize=9,
                            textColor=URB_BLACK, spaceAfter=2*mm)))
                items.append(Paragraph(contenido_txt,
                    ParagraphStyle("bc", fontName="Helvetica", fontSize=8,
                        textColor=colors.HexColor("#333333"), leading=13, spaceAfter=0)))

                box = Table([[items]], colWidths=[17*cm])
                box.setStyle(TableStyle([
                    ("BACKGROUND",    (0,0), (-1,-1), colors.HexColor("#f8f9fa")),
                    ("BOX",           (0,0), (-1,-1), 0.5, colors.HexColor("#dddddd")),
                    ("LEFTPADDING",   (0,0), (-1,-1), 8),
                    ("RIGHTPADDING",  (0,0), (-1,-1), 8),
                    ("TOPPADDING",    (0,0), (-1,-1), 6),
                    ("BOTTOMPADDING", (0,0), (-1,-1), 6),
                    ("VALIGN",        (0,0), (-1,-1), "TOP"),
                ]))
                story.append(KeepTogether([box, Spacer(1, 3*mm)]))

    # ── FIRMA ─────────────────────────────────────────────────
    story.append(Spacer(1, 12*mm))
    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#dddddd")))
    story.append(Spacer(1, 4*mm))
    firma_data = [[
        Paragraph("_________________________\n<b>Ing. Luis Alfaro Arguedas</b>\nCFIA IC-25146\nUrbanistyka Constructora",
            ParagraphStyle("firma", fontName="Helvetica", fontSize=8, leading=14, textColor=colors.HexColor("#333333"))),
        Paragraph(f"<b>San Rafael de Alajuela, {datetime.now().strftime('%d de %B de %Y')}</b>",
            ParagraphStyle("fecha_firma", fontName="Helvetica", fontSize=8, textColor=URB_GRAY, alignment=TA_RIGHT))
    ]]
    firma_table = Table(firma_data, colWidths=[10*cm, 7*cm])
    firma_table.setStyle(TableStyle([("VALIGN", (0,0), (-1,-1), "BOTTOM")]))
    story.append(firma_table)

    doc.build(story, onFirstPage=on_page, onLaterPages=on_page)
    output.seek(0)
    return output


@app.route("/api/informe/<codigo_proyecto>/<tipo>")
def descargar_informe(codigo_proyecto, tipo):
    """Genera y descarga el informe PDF."""
    if "tipo" not in session:
        return "No autorizado", 401
    if not session.get("admin") and session.get("codigo") != codigo_proyecto:
        return "No autorizado", 401
    if tipo not in ("financiero", "avance", "completo"):
        return "Tipo inválido", 400

    pdf = generar_pdf_informe(codigo_proyecto, tipo)
    if not pdf:
        return "Proyecto no encontrado", 404

    nombres = {
        "financiero": "Informe_Financiero",
        "avance":     "Informe_Avance",
        "completo":   "Informe_Completo"
    }
    return send_file(pdf, as_attachment=True,
        download_name=f"{nombres[tipo]}_{codigo_proyecto}_{datetime.now().strftime('%Y%m%d')}.pdf",
        mimetype="application/pdf")
