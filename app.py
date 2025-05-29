import os
import json
import io
import base64
import re
from flask import Flask, render_template, request, redirect, url_for, flash, send_file
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from functools import wraps
from docxtpl import DocxTemplate
from unidecode import unidecode

# —– Helper: normalizar cabeceras a claves válidas para docxtpl —–
def normalize(header: str) -> str:
    """
    Convierte una cabecera como "CSV Guest NM" o
    "PRE ARRIVAL NOTAS (BORRADOR)" en un identificador
    tipo "csv_guest_nm" o "pre_arrival_notas_borrador".
    """
    h = header.strip()
    h = unidecode(h)                # quita acentos
    h = h.lower()                   # minúsculas
    h = re.sub(r'[^a-z0-9]+', '_', h)  # no alfanumérico → '_'
    return h.strip('_')

# —– Configuración de Flask —–
app = Flask(__name__, template_folder='templates', static_folder='static')
app.secret_key = os.environ.get('SECRET_KEY', 'cambia_esto_por_un_valor_segura')

# —– Usuarios y permisos —–
USERS = {
    os.environ.get('BASIC_USER',   'admin'):    os.environ.get('BASIC_PASS',   'password'),
    os.environ.get('BASIC_USER2', 'usuario2'): os.environ.get('BASIC_PASS2', 'pass2'),
}
PERMISSIONS = {
    'usuario2': [
        'ITINERARIO',
        'AMENIDAD',
        'PRE ARRIVAL NOTAS (BORRADOR)',
        'REGISTRO DE CONTACTO',
        'Transfer'
    ]
}

def check_auth(u, p):
    return USERS.get(u) == p

def authenticate():
    return ('Autorización requerida.'), 401, {
        'WWW-Authenticate': 'Basic realm="Login Required"'
    }

def requires_auth(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        auth = request.authorization
        if not auth or not check_auth(auth.username, auth.password):
            return authenticate()
        return f(*args, **kwargs)
    return decorated

# —– Google Sheets setup —–
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly"
]
if 'GOOGLE_SHEETS_JSON_B64' in os.environ:
    raw = base64.b64decode(os.environ['GOOGLE_SHEETS_JSON_B64'])
    CRED_FILE = '/tmp/credentials.json'
    with open(CRED_FILE, 'wb') as f:
        f.write(raw)
else:
    BASE_DIR  = os.path.dirname(os.path.abspath(__file__))
    CRED_FILE = os.path.join(BASE_DIR, 'credentials.json')

creds  = ServiceAccountCredentials.from_json_keyfile_name(CRED_FILE, SCOPES)
client = gspread.authorize(creds)

# —– Tu hoja y worksheet —–
SHEET_ID  = '1LDhajDpQTzi0RLw8BXLTzmA1m9yRlTX_SrxC9aKLKYg'
worksheet = client.open_by_key(SHEET_ID).worksheet('hoja')

# —– Rutas —–
@app.route('/', methods=['GET', 'POST'])
def index():
    user    = None
    allowed = None

    if request.method == 'POST':
        auth    = request.authorization
        user    = auth.username if auth else None
        allowed = PERMISSIONS.get(user)

        search_id = request.form.get('search_id', '').strip()
        if not search_id:
            flash('El campo ID no puede estar vacío.', 'error')
        else:
            try:
                cell    = worksheet.find(search_id, in_column=3)
                row_idx = cell.row
                headers = worksheet.row_values(1)
                values  = worksheet.row_values(row_idx)
                record  = dict(zip(headers, values))
                record['row_idx'] = row_idx

                return render_template(
                    'edit.html',
                    record=record,
                    headers=headers,
                    user=user,
                    allowed=allowed
                )
            except Exception:
                flash('ID no encontrado. Try again.', 'error')

    return render_template('index.html')

@app.route('/update', methods=['POST'])
@requires_auth
def update():
    auth    = request.authorization
    user    = auth.username
    allowed = PERMISSIONS.get(user)

    row_idx = int(request.form.get('row_idx'))
    headers = worksheet.row_values(1)

    # — Guardar cambios en Google Sheets —
    try:
        for col_idx, header in enumerate(headers, start=1):
            if allowed is not None and header not in allowed:
                continue
            new_val = request.form.get(header, '')
            worksheet.update_cell(row_idx, col_idx, new_val)
    except Exception:
        flash('Error al guardar cambios. Intenta de nuevo.', 'error')
        record = {h: request.form.get(h, '') for h in headers}
        record['row_idx'] = row_idx
        return render_template('edit.html',
                               record=record,
                               headers=headers,
                               user=user,
                               allowed=allowed)

    # — Exportar a Word usando plantilla docxtpl —
    if request.form.get('export'):
        # 1) Leer toda la fila original
        raw_record = dict(zip(
            headers,
            worksheet.row_values(row_idx)
        ))
        # 2) Construir contexto normalizado
        context = { normalize(h): raw_record[h] for h in headers }
        # 3) Cargar y renderizar plantilla
        BASE_DIR = os.path.dirname(os.path.abspath(__file__))
        tpl_path = os.path.join(BASE_DIR, 'templates_docx', 'itinerary_template.docx')
        doc = DocxTemplate(tpl_path)
        doc.render(context)
        # 4) Enviar archivo
        bio = io.BytesIO()
        doc.save(bio)
        bio.seek(0)
        filename = f"Itinerary_{ context.get('pms_confirm_no', row_idx) }.docx"
        return send_file(
            bio,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    flash('Registro actualizado con éxito.', 'success')
    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)
