import os
import io
import base64
import unicodedata
import re
from functools import wraps

from flask import Flask, render_template, request, redirect, url_for, flash, send_file
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from docxtpl import DocxTemplate

app = Flask(__name__, template_folder='templates', static_folder='static')
app.secret_key = os.environ.get('SECRET_KEY', 'cambia_esto_por_un_clave_segura')

# — Usuarios y permisos —
USERS = {
    os.environ.get('BASIC_USER',   'admin'):    os.environ.get('BASIC_PASS',   'password'),
    os.environ.get('BASIC_USER2','usuario2'): os.environ.get('BASIC_PASS2','pass2')
}
PERMISSIONS = {
    'usuario2': ['Hora de Llegada','CSV_Cust_Email','CSV_Cust_Phone1','cuantosAcomp']
}

def check_auth(u, p):
    return USERS.get(u) == p

def authenticate():
    return ('Autorización requerida.'), 401, {
        'WWW-Authenticate': 'Basic realm="Login Required"'
    }

def requires_auth(f):
    @wraps(f)
    def deco(*args, **kwargs):
        auth = request.authorization
        if not auth or not check_auth(auth.username, auth.password):
            return authenticate()
        return f(*args, **kwargs)
    return deco

def normalize_string(s: str) -> str:
    """Convierte nombres de columna a snake_case sin acentos."""
    s = ''.join(c for c in unicodedata.normalize('NFD', s)
                if unicodedata.category(c) != 'Mn')
    s = re.sub(r'[^A-Za-z0-9]', '_', s).lower()
    return re.sub(r'__+', '_', s).strip('_')  # <-- comilla corregida aquí

# — Etiquetas legibles —
display_names = {
    'ID_Reserva':'Clave de reservación',
    'Pms_Confirm_No':'Clave de Central',
    'CSV_Guest_NM':'Nombre huésped',
    'CSV_Cust_Phone1':'Teléfono',
    'CSV_Cust_Email':'Email',
    'CSV_Arrival_Date':'Check-in',
    'CSV_Depart_Date':'Check-out',
    'CSV_Nights_Qty':'No. de noches',
    'CSV_Status':'Estatus reserva',
    'Which_Date':'Fecha reserva',
    'Hora de Llegada':'Hora de llegada',
    'cuantosAcomp':'No. de personas',
}

# — Google Sheets setup —
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly"
]
if 'GOOGLE_SHEETS_JSON_B64' in os.environ:
    raw       = base64.b64decode(os.environ['GOOGLE_SHEETS_JSON_B64'])
    CRED_FILE = '/tmp/credentials.json'
    with open(CRED_FILE,'wb') as f: f.write(raw)
else:
    BASE_DIR  = os.path.dirname(os.path.abspath(__file__))
    CRED_FILE = os.path.join(BASE_DIR,'credentials.json')

creds     = ServiceAccountCredentials.from_json_keyfile_name(CRED_FILE, SCOPES)
client    = gspread.authorize(creds)
worksheet = client.open_by_key(
    '1LDhajDpQTzi0RLw8BXLTzmA1m9yRlTX_SrxC9aKLKYg'
).worksheet('hoja')

@app.route('/', methods=['GET'])
@requires_auth
def index():
    user = request.authorization.username
    return render_template('index.html',
                           display_names=display_names,
                           user=user)

@app.route('/search', methods=['POST'])
@requires_auth
def search():
    auth    = request.authorization; user = auth.username
    allowed = PERMISSIONS.get(user)

    # Recolectar criterios no vacíos
    fields = ['ID_Reserva','Pms_Confirm_No','CSV_Guest_NM',
              'CSV_Cust_Email','CSV_Arrival_Date','Which_Date']
    criteria = {
        f: request.form.get(f).strip()
        for f in fields
        if request.form.get(f) and request.form.get(f).strip()
    }
    if not criteria:
        flash('Debes indicar al menos un criterio de búsqueda.', 'error')
        return redirect(url_for('index'))

    # Índices de columnas
    headers     = worksheet.row_values(1)
    col_indices = {h: i+1 for i, h in enumerate(headers)}

    # Búsqueda por criterio
    row_sets = []
    for field, val in criteria.items():
        col = col_indices.get(field)
        if not col:
            row_sets.append(set())
            continue
        col_vals = worksheet.col_values(col)
        matches  = {
            idx for idx, cell in enumerate(col_vals[1:], start=2)
            if cell.strip().lower() == val.lower()
        }
        row_sets.append(matches)

    matches = set.intersection(*row_sets) if row_sets else set()
    if not matches:
        flash('No se encontraron reservas que cumplan los criterios.', 'error')
        return redirect(url_for('index'))

    if len(matches) == 1:
        return redirect(url_for('edit_record', row_idx=matches.pop()))

    # Varias coincidencias: listar
    records = []
    for r in sorted(matches):
        vals = worksheet.row_values(r)
        rec  = dict(zip(headers, vals))
        records.append({
            'CSV_Guest_NM':   rec.get('CSV_Guest_NM',''),
            'CSV_Cust_Email': rec.get('CSV_Cust_Email',''),
            'Pms_Confirm_No': rec.get('Pms_Confirm_No',''),
            'row_idx':        r
        })

    return render_template('index.html',
                           records=records,
                           display_names=display_names,
                           user=user)

@app.route('/edit/<int:row_idx>', methods=['GET'])
@requires_auth
def edit_record(row_idx):
    auth      = request.authorization; user = auth.username
    allowed   = PERMISSIONS.get(user)
    headers   = worksheet.row_values(1)
    vals      = worksheet.row_values(row_idx)
    record    = dict(zip(headers, vals))
    record['row_idx'] = row_idx

    # Reservas previas por email y nombre
    email     = record.get('CSV_Cust_Email','').strip().lower()
    name      = record.get('CSV_Guest_NM','').strip().lower()
    prev_rows = set()
    if email:
        col = headers.index('CSV_Cust_Email') + 1
        prev_rows |= {c.row for c in worksheet.findall(email, in_column=col)}
    if name:
        col = headers.index('CSV_Guest_NM') + 1
        prev_rows |= {c.row for c in worksheet.findall(name, in_column=col)}
    prev_rows.discard(row_idx)

    previous_records = []
    for pr in sorted(prev_rows):
        v = worksheet.row_values(pr)
        d = dict(zip(headers, v))
        previous_records.append({
            'row_idx':        pr,
            'CSV_Cust_Email': d.get('CSV_Cust_Email',''),
            'Which_Date':     d.get('Which_Date','')
        })

    # Cargar faltantes si se solicita
    filled_fields = []
    src = request.args.get('source')
    if src and request.args.get('fill') == '1':
        src_idx  = int(src)
        src_vals = worksheet.row_values(src_idx)
        src_rec  = dict(zip(headers, src_vals))
        for h in headers:
            if not record.get(h) and src_rec.get(h):
                record[h] = src_rec[h]
                filled_fields.append(h)

    return render_template('edit.html',
                           record=record,
                           headers=headers,
                           user=user,
                           allowed=allowed,
                           display_names=display_names,
                           previous_records=previous_records,
                           filled_fields=filled_fields)

@app.route('/update', methods=['POST'])
@requires_auth
def update():
    auth      = request.authorization; user = auth.username
    allowed   = PERMISSIONS.get(user)
    row_idx   = int(request.form.get('row_idx'))
    headers   = worksheet.row_values(1)

    # 1) Guardar en Sheets
    for i, h in enumerate(headers, start=1):
        if allowed is not None and h not in allowed:
            continue
        val = request.form.get(h, '')
        print(f"Updating row {row_idx}, col {i} ({h}) => '{val}'")
        try:
            worksheet.update_cell(row_idx, i, val)
        except Exception as e:
            flash(f"Error al actualizar «{h}»: {e}", 'error')
            return redirect(url_for('edit_record', row_idx=row_idx))

    # 2) Si clic en “Guardar y exportar Word”
    if request.form.get('save_export'):
        context = {}
        for h in headers:
            context[normalize_string(h)] = request.form.get(h, '')
        tpl_path = os.path.join(
            os.path.dirname(__file__),
            'templates_docx',
            'itinerary_template.docx'
        )
        doc = DocxTemplate(tpl_path)
        doc.render(context)
        bio = io.BytesIO()
        doc.save(bio); bio.seek(0)
        fname = f"Itinerary_{context.get('pms_confirm_no', row_idx)}.docx"
        return send_file(
            bio,
            as_attachment=True,
            download_name=fname,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )

    # 3) Guardado normal
    flash('Registro actualizado con éxito.', 'success')
    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)
