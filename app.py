import os
import json
import io
import base64
import unicodedata

from flask import Flask, render_template, request, redirect, url_for, flash, send_file
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from functools import wraps
from docxtpl import DocxTemplate

# —– Configuración de Flask —–
app = Flask(__name__, template_folder='templates', static_folder='static')
app.secret_key = os.environ.get('SECRET_KEY', 'cambia_esto_por_un_clave_segura')

# —– Usuarios y permisos —–
USERS = {
    os.environ.get('BASIC_USER',   'admin'):    os.environ.get('BASIC_PASS',   'password'),
    os.environ.get('BASIC_USER2', 'usuario2'): os.environ.get('BASIC_PASS2', 'pass2')
}

# Columnas que usuario2 puede editar
PERMISSIONS = {
    'usuario2': [
        'ITINERARIO',
        'AMENIDAD',
        'PRE ARRIVAL NOTAS (BORRADOR)',
        'REGISTRO DE CONTACTO',
        'Transfer'
    ]
}

def check_auth(username, password):
    return USERS.get(username) == password

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

# Si recibimos JSON en Base64, volcamos a disco
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

# ID de tu hoja y nombre de pestaña
SHEET_ID   = '1LDhajDpQTzi0RLw8BXLTzmA1m9yRlTX_SrxC9aKLKYg'
worksheet  = client.open_by_key(SHEET_ID).worksheet('hoja')


# —– Función auxiliar: normaliza una cadena de texto para que sirva como clave —–
def normalize_string(s: str) -> str:
    """
    1) convierte a minúsculas
    2) elimina acentos
    3) sustituye cualquier carácter no alfanumérico (incluyendo espacios) por '_'
    4) quita múltiples guiones bajos consecutivos
    """
    # 1) a minúsculas
    s = s.lower()
    # 2) eliminar acentos y signos diacríticos
    s = ''.join(
        c for c in unicodedata.normalize('NFD', s)
        if unicodedata.category(c) != 'Mn'
    )
    # 3) reemplazar cualquier cosa que no sea letra o número por guión bajo
    import re
    s = re.sub(r'[^a-z0-9]', '_', s)
    # 4) colapsar guiones bajos repetidos a uno solo
    s = re.sub(r'__+', '_', s)
    # 5) quitar guiones bajos al principio o final
    s = s.strip('_')
    return s


# —– Rutas —–
@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
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

                user    = request.authorization.username if request.authorization else None
                allowed = PERMISSIONS.get(user, None)

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
    allowed = PERMISSIONS.get(user, None)

    row_idx = int(request.form.get('row_idx'))
    headers = worksheet.row_values(1)

    # — Guardar cambios en Google Sheets —
    try:
        for col_idx, header in enumerate(headers, start=1):
            # si hay permisos limitados y la columna NO está permitida, la saltamos
            if allowed is not None and header not in allowed:
                continue
            new_val = request.form.get(header, '')
            worksheet.update_cell(row_idx, col_idx, new_val)
    except Exception:
        flash('Error al guardar cambios. Intenta de nuevo.', 'error')
        record = {h: request.form.get(h, '') for h in headers}
        record['row_idx'] = row_idx
        return render_template(
            'edit.html',
            record=record,
            headers=headers,
            user=user,
            allowed=allowed
        )

    # — Exportar a Word usando plantilla docxtpl —
    if request.form.get('export'):
        # Creamos un diccionario “normalizado” de context:
        #   { 'csv_guest_nm': 'Juan Pérez', 'csv_nights_qty': '3', … }
        context = {}
        for h in headers:
            key_norm = normalize_string(h)  # ej: "Preferencia Café y Lácteos" -> "preferencia_cafe_y_lacteos"
            context[key_norm] = request.form.get(h, '')

        # Ruta a tu plantilla .docx
        BASE_DIR = os.path.dirname(os.path.abspath(__file__))
        tpl_path = os.path.join(
            BASE_DIR,
            'templates_docx',
            'itinerary_template.docx'
        )
        doc = DocxTemplate(tpl_path)
        doc.render(context)

        bio = io.BytesIO()
        doc.save(bio)
        bio.seek(0)

        # Usamos, por ejemplo, el “PMS Confirm No.” como nombre de archivo. Si no existe,
        # caemos al row_idx:
        filename = f"Itinerary_{context.get('pms_confirm_no', row_idx)}.docx"
        return send_file(
            bio,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )

    flash('Registro actualizado con éxito.', 'success')
    return redirect(url_for('index'))


if __name__ == '__main__':
    app.run(debug=True)
